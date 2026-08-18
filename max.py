
from __future__ import annotations

import logging
import mimetypes
import os
import re
import time
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

import truststore

truststore.inject_into_ssl()

import requests
from dateutil.relativedelta import relativedelta
from docx.opc.part import PartFactory
from docx.package import Package
from docx.parts.document import DocumentPart
import docxtpl.template as docxtpl_template
from dotenv import load_dotenv
from core import misc as misc

API_URL = "https://platform-api2.max.ru"
BOT_DIR = Path(__file__).resolve().parent
load_dotenv(BOT_DIR / ".env")
MESSAGES_DIR = BOT_DIR / "messages"
IMAGES_DIR = BOT_DIR / "images"
TOKEN = os.getenv("MAX_BOT_TOKEN", "").strip()
if TOKEN.casefold().startswith("bearer "):
    raise RuntimeError(
        "MAX_BOT_TOKEN должен содержать только токен, без префикса Bearer"
    )
try:
    STAFF = {
        int(item.strip())
        for item in os.getenv("MAX_STAFF_IDS", "").split(",")
        if item.strip()
    }
except ValueError as exc:
    raise RuntimeError(
        "MAX_STAFF_IDS должен содержать ID сотрудников через запятую"
    ) from exc
# if not STAFF:
#     raise RuntimeError("Не задана переменная окружения MAX_STAFF_IDS")
START_BUTTON_TEXT = "🚀 Начать"
RESTART_BUTTON_TEXT = "🔄 Перезапустить"
START_WORDS = {
    "/start",
    START_BUTTON_TEXT.casefold(),
    RESTART_BUTTON_TEXT.casefold(),
    "старт",
    "начать",
    "начало",
    "начни",
    "перезапустить",
    "начать заново",
}
RECENT_ACTIONS: dict[int, tuple[str, float]] = {}
ACTION_DEBOUNCE_SECONDS = 1.5
IMAGES_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%d-%m-%Y %H:%M:%S",
)
log = logging.getLogger("max-bot")
ms = misc.Miscellaneous(bot_dir=BOT_DIR, trust_env=False)
DIRECT_HTTP = requests.Session()
DIRECT_HTTP.trust_env = False

# кортеж сессий пользователей бота, чтобы они не нарушалась работа переменных
user_data = {}

def open_word_document(filename: str | Path):
    return Package.open(filename).main_document_part.document

PartFactory.part_type_for[ms.DOCM_MAIN_CONTENT_TYPE] = DocumentPart
docxtpl_template.Document = open_word_document

class MaxAuthenticationError(RuntimeError):
    """MAX отклонил токен чат-бота."""


class MaxAPI:
    def __init__(self, token: str):
        if not token:
            raise RuntimeError("Не задана переменная окружения MAX_BOT_TOKEN")
        self.session = requests.Session()
        # MAX всегда использует прямое соединение и не наследует HTTP(S)_PROXY.
        self.session.trust_env = False
        self.session.headers["Authorization"] = token

    def request(self, method: str, path: str, **kwargs) -> dict[str, Any]:
        response = self.session.request(
            method, f"{API_URL}{path}", timeout=kwargs.pop("timeout", 40), **kwargs
        )

        if response.status_code == 401:
            log.error(
                "MAX отклонил авторизацию | method=%s | path=%s | status=401",
                method,
                path,
            )
            raise MaxAuthenticationError(
                "MAX отклонил MAX_BOT_TOKEN. Получите актуальный токен в "
                "настройках чат-бота и перезапустите программу."
            )

        if not response.ok:
            detail = response.text[:500].replace("\r", " ").replace("\n", " ")
            log.error(
                "Ошибка MAX API | method=%s | path=%s | status=%s | response=%s",
                method,
                path,
                response.status_code,
                detail,
            )

        response.raise_for_status()
        return response.json() if response.content else {}

    def send(
        self,
        user_id: int,
        text: str = "",
        attachments: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        body: dict[str, Any] = {"text": text, "format": "html"}
        if attachments:
            body["attachments"] = attachments
        return self.request("POST", "messages", params={"user_id": user_id}, json=body)

    def me(self) -> dict[str, Any]:
        return self.request("GET", "/me")

    def updates(self, marker: int | None) -> dict[str, Any]:
        params: dict[str, Any] = {
            "limit": 100,
            "timeout": 30,
            "types": "bot_started,message_created,message_callback",
        }
        if marker is not None:
            params["marker"] = marker
        return self.request("GET", "/updates", params=params, timeout=40)

    def upload_file(self, path: Path) -> dict[str, Any]:
        upload = self.request(
            "POST",
            "/uploads",
            params={"type": "file"},
        )

        suffix = path.suffix.lower()

        mime_types = {
            ".docm": "application/vnd.ms-word.document.macroEnabled.12",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".doc": "application/msword",
            ".pdf": "application/pdf",
        }

        content_type = mime_types.get(
            path.suffix.lower(),
            "application/octet-stream",
        )

        with path.open("rb") as stream:
            response = DIRECT_HTTP.post(
                upload["url"],
                files={
                    "data": (
                        path.name,
                        stream,
                        content_type,
                    )
                },
                timeout=180,
            )

        response.raise_for_status()

        payload = response.json()

        log.info(
            "Файл загружен в MAX | filename=%s | mime=%s | bytes=%d",
            path.name,
            content_type,
            path.stat().st_size,
        )

        return {
            "type": "file",
            "payload": payload,
        }


bot = MaxAPI(TOKEN)


def session(user_id: int) -> dict[str, Any]:
    return user_data.setdefault(
        user_id,
        {
            "cost": 0,
            "lastname": None,
            "ending": None,
            "complects": 0,
            "count_print": 0,
            "count_sending": 0,
            "stage": "idle",
        },
    )


def is_repeated_action(user_id: int, action: str) -> bool:
    """Подавляет повтор одного действия, сохраняя отзывчивость других кнопок."""
    now = time.monotonic()
    previous_action, previous_at = RECENT_ACTIONS.get(user_id, ("", 0.0))
    RECENT_ACTIONS[user_id] = (action, now)
    return (
        previous_action == action
        and now - previous_at < ACTION_DEBOUNCE_SECONDS
    )


def allowed(user_id: int) -> bool:
    if user_id in STAFF:
        return True
    log.warning("Попытка запуска пользователем %s", user_id)
    bot.send(
        user_id,
        "❌ <b>Вам запрещён доступ к боту!</b> Обратитесь к техническому "
        "специалисту СпецЭнергоразвития для получения доступа.",
    )
    return False


def landing(user_id: int) -> None:
    if not allowed(user_id):
        return
    user_data[user_id] = {
        "cost": 0,
        "lastname": None,
        "ending": None,
        "complects": 0,
        "count_print": 0,
        "count_sending": 0,
        "stage": "name",
    }
    hello_path = MESSAGES_DIR / "hello.txt"
    start_message = (
        hello_path.read_text(encoding="utf-8", errors="ignore")
        if hello_path.exists()
        else "Здравствуйте! Введите фамилию."
    )
    log.info("Запущен бот пользователем %s", user_id)
    bot.send(user_id, start_message, [control_keyboard()])


def control_keyboard() -> dict[str, Any]:
    return {
        "type": "inline_keyboard",
        "payload": {
            "buttons": [
                [
                    {
                        "type": "callback",
                        "text": START_BUTTON_TEXT,
                        "payload": "start",
                    },
                    {
                        "type": "callback",
                        "text": RESTART_BUTTON_TEXT,
                        "payload": "restart",
                    },
                ]
            ]
        },
    }


def show_start_menu(user_id: int, text: str | None = None) -> None:
    if not allowed(user_id):
        return
    user_data[user_id] = {
        "cost": 0,
        "lastname": None,
        "ending": None,
        "complects": 0,
        "count_print": 0,
        "count_sending": 0,
        "stage": "idle",
    }
    bot.send(
        user_id,
        text or "Нажмите кнопку, чтобы сформировать новый проект договора.",
        [control_keyboard()],
    )


def source_keyboard() -> dict[str, Any]:
    return {
        "type": "inline_keyboard",
        "payload": {
            "buttons": [
                [
                    {"type": "callback", "text": "🧾 Документ", "payload": "doc"},
                    {"type": "callback", "text": "🖼️ Картинка", "payload": "pic"},
                ],
                [
                    {"type": "callback", "text": "📄 PDF скана", "payload": "pdf"},
                    {"type": "callback", "text": "📢 Сообщением", "payload": "mes"},
                ],
                [
                    {
                        "type": "callback",
                        "text": RESTART_BUTTON_TEXT,
                        "payload": "restart",
                    }
                ],
            ]
        },
    }


def message_text(message: dict[str, Any]) -> str:
    return str((message.get("body") or {}).get("text") or "").strip()


def sender_id(message: dict[str, Any], update: dict[str, Any]) -> int:
    sender = message.get("sender") or update.get("user") or {}
    value = sender.get("user_id")
    if value is None:
        raise ValueError("В обновлении MAX отсутствует sender.user_id")
    return int(value)


def attachment_url(attachment: dict[str, Any]) -> str | None:
    """Находит URL оригинала в payload вложения разных версий MAX API."""
    preferred = ("url", "download_url")

    def walk(value: Any) -> str | None:
        if isinstance(value, dict):
            for key in preferred:
                candidate = value.get(key)
                if isinstance(candidate, str) and candidate.startswith(("http://", "https://")):
                    return candidate
            for nested in value.values():
                found = walk(nested)
                if found:
                    return found
        elif isinstance(value, list):
            for nested in value:
                found = walk(nested)
                if found:
                    return found
        return None

    return walk(attachment.get("payload") or attachment)


def attachment_filename(attachment: dict[str, Any]) -> str | None:
    """Находит исходное имя файла в разных вариантах payload MAX."""
    preferred = ("filename", "file_name", "name")

    def walk(value: Any) -> str | None:
        if isinstance(value, dict):
            for key in preferred:
                candidate = value.get(key)
                if isinstance(candidate, str) and candidate.strip():
                    return candidate.strip()
            for nested in value.values():
                found = walk(nested)
                if found:
                    return found
        elif isinstance(value, list):
            for nested in value:
                found = walk(nested)
                if found:
                    return found
        return None

    return walk(attachment.get("payload") or attachment)


def detect_attachment_suffix(
    attachment: dict[str, Any],
    url: str,
    response: requests.Response,
    expected: str,
) -> str:
    """Определяет формат независимо от расширения подписанного URL MAX."""
    original_name = attachment_filename(attachment)
    disposition = response.headers.get("Content-Disposition", "")
    disposition_match = re.search(
        r"filename\*?=(?:UTF-8''|\")?([^\";]+)",
        disposition,
        flags=re.IGNORECASE,
    )
    disposition_name = (
        unquote(disposition_match.group(1).strip()) if disposition_match else None
    )

    document_suffixes = set(ms.ALLOWED_EXTENSIONS)
    image_suffixes = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
    accepted_suffixes = image_suffixes if expected == "image" else document_suffixes

    for candidate in (original_name, disposition_name, unquote(urlparse(url).path)):
        suffix = Path(candidate or "").suffix.lower()
        if suffix in accepted_suffixes:
            return suffix

    content_type = response.headers.get("Content-Type", "").split(";", 1)[0].lower()
    known_mime_types = {
        "application/pdf": ".pdf",
        "application/msword": ".doc",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
        "image/jpeg": ".jpg",
        "image/png": ".png",
        "image/gif": ".gif",
        "image/bmp": ".bmp",
        "image/tiff": ".tiff",
        "image/webp": ".webp",
    }
    if (
        content_type in known_mime_types
        and known_mime_types[content_type] in accepted_suffixes
    ):
        return known_mime_types[content_type]
    guessed = mimetypes.guess_extension(content_type) if content_type else None
    if guessed and guessed.lower() in accepted_suffixes:
        return guessed.lower()

    content = response.content
    if expected == "file" and content.lstrip().startswith(b"%PDF-"):
        return ".pdf"
    if expected == "file" and content.startswith(
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    ):
        return ".doc"
    if expected == "file" and content.startswith(b"PK"):
        try:
            with zipfile.ZipFile(BytesIO(content)) as archive:
                if "word/document.xml" in archive.namelist():
                    return ".docx"
        except (OSError, zipfile.BadZipFile):
            pass

    image_signatures = (
        (b"\xff\xd8\xff", ".jpg"),
        (b"\x89PNG\r\n\x1a\n", ".png"),
        (b"GIF87a", ".gif"),
        (b"GIF89a", ".gif"),
        (b"BM", ".bmp"),
        (b"II*\x00", ".tiff"),
        (b"MM\x00*", ".tiff"),
    )
    for signature, suffix in image_signatures:
        if expected == "image" and content.startswith(signature):
            return suffix

    if expected == "image":
        return ".jpg"
    raise misc.DocumentExtractionError(
        "MAX передал файл без имени и распознаваемого формата. "
        "Отправьте его повторно как DOC, DOCX или PDF."
    )


def download_attachment(message: dict[str, Any], expected: str) -> Path:
    attachments = (message.get("body") or {}).get("attachments") or []
    selected = next((a for a in attachments if a.get("type") == expected), None)
    if selected is None:
        raise ValueError("Сообщение не содержит нужного вложения")
    url = attachment_url(selected)
    if not url:
        raise ValueError("MAX не прислал URL вложения")
    response = DIRECT_HTTP.get(url, timeout=180)
    response.raise_for_status()
    if len(response.content) > ms.MAX_FILE_BYTES:
        raise misc.DocumentExtractionError(
            f"Файл больше {ms.MAX_FILE_BYTES // (1024 * 1024)} МБ"
        )
    suffix = detect_attachment_suffix(selected, url, response, expected)
    if expected == "file" and suffix not in ms.ALLOWED_EXTENSIONS:
        raise misc.DocumentExtractionError(
            f"Формат {suffix} не поддерживается. Отправьте DOC, DOCX или PDF."
        )
    target = IMAGES_DIR / f"max_{int(time.time() * 1000)}{suffix}"
    target.write_bytes(response.content)
    log.info(
        "Вложение MAX загружено | expected=%s | filename=%s | extension=%s | "
        "content_type=%s | bytes=%d",
        expected,
        attachment_filename(selected) or "<unknown>",
        suffix,
        response.headers.get("Content-Type", "<unknown>"),
        len(response.content),
    )
    return target


def finish_document(user_id: int, company_data: Any, source_path: Path | None) -> None:
    data = session(user_id)
    ms.validate_company_data(company_data)
    bot.send(
        user_id,
        "🧾 Реквизиты извлечены. Рассчитываю суммы и заполняю шаблон…",
        [control_keyboard()],
    )
    number_contract = ms.get_bot_doc_num(user_data, user_id)
    number_invoice = ms.get_bot_count_num(user_data, user_id)
    texted_costs = ms.integer_texted(data["cost"])
    texted_sending = ms.integer_texted(data["count_sending"])
    local_doc = Path(
        ms.bot_insert_req(
            user_data,
            user_id,
            company_data,
            number_contract,
            number_invoice,
            texted_costs,
            texted_sending,
            str(source_path) if source_path else None,
        )
    )
    try:
        bot.send(user_id, "📤 Договор готов. Загружаю файл…", [control_keyboard()])
        attachment = bot.upload_file(local_doc)
        delays = (1, 2, 4, 8, 12)

        for attempt, delay in enumerate(delays, start=1):
            try:
                bot.send(user_id, "Ваши реквизиты", [attachment])
                break
            except requests.HTTPError as exc:
                response = exc.response
                try:
                    error = response.json() if response is not None else {}
                except ValueError:
                    error = {}
                if error.get("code") != "attachment.not.ready":
                    raise
                log.info(
                    "Файл MAX ещё обрабатывается; попытка %s/%s через %s сек.",
                    attempt,
                    len(delays),
                    delay,
                )
                time.sleep(delay)
            else:
                raise RuntimeError("MAX не подготовил загруженный договор за отведённое время")
        #bot.send(user_id, "Ваши реквизиты", [bot.upload_file(local_doc)])
    finally:
        ms.delete_local_doc(str(local_doc))
        if source_path and source_path.exists():
            source_path.unlink(missing_ok=True)
    show_start_menu(user_id, "✅ Договор отправлен. Можно начать новый сценарий.")


def handle_input(user_id: int, message: dict[str, Any]) -> None:
    if not allowed(user_id):
        return
    data = session(user_id)
    text = message_text(message)
    stage = data["stage"]

    if text.casefold() in START_WORDS:
        landing(user_id)
        return

    if stage == "idle":
        show_start_menu(
            user_id,
            "⚠️ Активный шаг не найден. Возможно, бот был перезапущен. "
            "Нажмите «Начать».",
        )
        return

    if stage == "name":
        if len(text) < 2:
            bot.send(user_id, "❌ Фамилия должна состоять более одной буквы!")
            return
        data["lastname"] = text[:2].upper()
        ending = datetime.now() + relativedelta(years=1)
        data["ending"] = f"{ending.day} {ms.GENITIVUS[ending.month]} {ending.year} года"
        data["stage"] = "cost"
        bot.send(
            user_id,
            f"✅ Фамилия принята. Код документов: <b>{data['lastname']}</b>\n"
            "❔ Введите стоимость разовой услуги",
            [control_keyboard()],
        )
        return

    if stage in {"cost", "complects"}:
        if not text.isdigit():
            bot.send(user_id, "❌ Число введено некорректно!")
            return
        value = int(text)
        if stage == "cost":
            data["cost"] = value
            data["stage"] = "complects"
            bot.send(
                user_id,
                f"✅ Стоимость принята: <b>{value:,} ₽</b>\n"
                "❔ Введите количество комплектов".replace(",", " "),
                [control_keyboard()],
            )
            return
        data["complects"] = value
        data["count_print"] = 3900 * value if value else 0
        data["count_sending"] = data["count_print"] + 1000 if value else 0
        data["stage"] = "choice"
        bot.send(
            user_id,
            f"✅ Количество комплектов принято: <b>{value}</b>\n"
            "❔ Выберите источник реквизитов",
            [source_keyboard()],
        )
        return

    source: Path | None = None
    normalized: Path | None = None
    try:
        if stage == "doc":
            source = download_attachment(message, "file")
            bot.send(
                user_id,
                "✅ Документ получен. Извлекаю текст и передаю реквизиты ИИ — "
                "это может занять до минуты.",
                [control_keyboard()],
            )
            finish_document(user_id, ms.sent_doc_to_ai(str(source)), source)
        elif stage == "pdf":
            source = download_attachment(message, "file")
            bot.send(
                user_id,
                "✅ PDF-скан получен. Распознаю страницы и извлекаю реквизиты — "
                "пожалуйста, не отправляйте файл повторно.",
                [control_keyboard()],
            )
            finish_document(user_id, ms.sent_pdf_scan_to_ai(str(source)), source)
        elif stage == "pic":
            source = download_attachment(message, "image")
            normalized = ms.normalize_jpeg(source)
            bot.send(
                user_id,
                "✅ Изображение получено. Распознаю текст и извлекаю реквизиты — "
                "это может занять до минуты.",
                [control_keyboard()],
            )
            finish_document(
                user_id,
                ms.sent_image_to_ai(str(normalized), "JPEG"),
                source,
            )
        elif stage == "mes":
            if not text:
                raise ValueError("Отправьте текстовое сообщение")
            bot.send(
                user_id,
                "✅ Сообщение принято. Извлекаю реквизиты…",
                [control_keyboard()],
            )
            finish_document(user_id, ms.sent_message_to_ai(text), None)
        else:
            show_start_menu(user_id)
    except misc.MissingCompanyDetailsError as exc:
        log.warning(
            "Формирование договора остановлено: не хватает реквизитов | "
            "user_id=%s | missing=%s",
            user_id,
            ", ".join(exc.missing_fields),
        )
        fields = "\n".join(f"• <b>{field}</b>" for field in exc.missing_fields)
        bot.send(
            user_id,
            "⚠️ Договор пока нельзя сформировать. Не хватает обязательных "
            f"реквизитов:\n\n{fields}\n\n"
            "Уточните их у заказчика и отправьте исправленные реквизиты повторно.",
            [control_keyboard()],
        )
    except misc.DocumentExtractionError as exc:
        log.warning(
            "Изображение или документ отклонены | user_id=%s | reason=%s",
            user_id,
            exc,
        )
        bot.send(
            user_id,
            f"❌ {exc}\n\nОтправьте другой файл или изображение.",
            [control_keyboard()],
        )
    except Exception as exc:
        log.exception("Ошибка обработки сообщения пользователя %s", user_id)
        bot.send(
            user_id,
            f"❌ Не удалось обработать данные: {exc}",
            [control_keyboard()],
        )
    finally:
        if normalized:
            normalized.unlink(missing_ok=True)
        if source:
            source.unlink(missing_ok=True)


def handle_update(update: dict[str, Any]) -> None:
    update_type = update.get("update_type")
    if update_type == "bot_started":
        user = update.get("user") or {}
        show_start_menu(int(user["user_id"]))
        return

    if update_type == "message_callback":
        callback = update.get("callback") or {}
        message = update.get("message") or callback.get("message") or {}
        user = callback.get("user") or update.get("user") or message.get("sender") or {}
        user_id = int(user["user_id"])
        if not allowed(user_id):
            return
        payload = callback.get("payload")
        if is_repeated_action(user_id, f"callback:{payload}"):
            log.debug(
                "Повторный callback подавлен | user_id=%s | payload=%s",
                user_id,
                payload,
            )
            return
        if payload in {"start", "restart"}:
            landing(user_id)
            return
        prompts = {
            "pic": ("pic", "<b>▶️ Вышлите картинку</b>"),
            "doc": ("doc", "<b>▶️ Вышлите документ</b>"),
            "pdf": ("pdf", "<b>▶️ Вышлите PDF-скан документа</b>"),
            "mes": ("mes", "<b>▶️ Отправьте или перешлите сообщение</b>"),
        }
        if payload in prompts:
            if session(user_id)["stage"] != "choice":
                show_start_menu(
                    user_id,
                    "⚠️ Предыдущий сценарий недоступен. Нажмите «Начать».",
                )
                return
            session(user_id)["stage"] = prompts[payload][0]
            source_names = {
                "pic": "изображение",
                "doc": "документ",
                "pdf": "PDF-скан",
                "mes": "текст сообщения",
            }
            bot.send(
                user_id,
                f"✅ Выбран источник: <b>{source_names[payload]}</b>\n"
                f"{prompts[payload][1]}",
                [control_keyboard()],
            )
        # POST /answers нужен только для изменения исходного сообщения или
        # показа одноразового уведомления. Для простого перехода сценария
        # отвечать на callback отдельным запросом не требуется. Пустое тело
        # /answers некоторые версии MAX API отклоняют с HTTP 400.
        return

    if update_type == "message_created":
        message = update.get("message") or {}
        user_id = sender_id(message, update)
        normalized_text = message_text(message).casefold()
        if normalized_text in START_WORDS:
            if is_repeated_action(user_id, f"message:{normalized_text}"):
                return
            landing(user_id)
        else:
            handle_input(user_id, message)


def main() -> None:
    marker: int | None = None
    try:
        identity = bot.me()
    except MaxAuthenticationError as exc:
        log.critical("MAX-бот остановлен: %s", exc)
        return
    log.info(
        "Авторизация MAX подтверждена | bot_id=%s | username=%s",
        identity.get("user_id"),
        identity.get("username"),
    )
    log.info("MAX-бот запущен")
    for user_id in STAFF:
        try:
            show_start_menu(
                user_id,
                "✅ Бот доступен после запуска. Можно начать новый сценарий.",
            )
        except requests.RequestException:
            log.warning(
                "Не удалось показать стартовую кнопку пользователю %s",
                user_id,
            )
    while True:
        try:
            page = bot.updates(marker)
            # Сначала запоминаем указатель страницы. Иначе исключение внутри
            # одного обработчика приводит к повторной выдаче всей страницы и
            # бесконечному повторению уже отправленных сообщений.
            next_marker = page.get("marker", marker)
            for update in page.get("updates", []):
                try:
                    handle_update(update)
                except Exception:
                    log.exception(
                        "Ошибка обработки update_type=%s; событие пропущено",
                        update.get("update_type"),
                    )
            marker = next_marker
        except MaxAuthenticationError as exc:
            log.critical("MAX-бот остановлен: %s", exc)
            return
        except requests.RequestException:
            log.exception("Ошибка MAX API; повтор через 3 секунды")
            time.sleep(3)
        except Exception:
            log.exception("Необработанная ошибка обновления")


if __name__ == "__main__":
    main()
