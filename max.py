"""MAX-интерфейс генератора договоров."""

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
from uuid import uuid4

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
LOGS_DIR = BOT_DIR / "logs"
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
ACTION_DEBOUNCE_SECONDS = 1.5
RECENT_ACTIONS: dict[int, tuple[str, float]] = {}

SER_QUESTIONS = [
    ("invoice_number", "❔ Введите номер счёта"),
    ("invoice_date", "❔ Введите дату счёта"),
    ("cost_month", "❔ Введите стоимость обслуживания в месяц"),
    ("advance", "❔ Введите сумму аванса"),
    ("object_address", "❔ Введите адрес объекта"),
    ("object_name", "❔ Введите наименование объекта"),
    ("service_period", "❔ Введите период тех. обслуживания"),
    ("email", "❔ Введите электронную почту заказчика"),
]
MONTHS_COUNT_OPTIONS = ["3", "2", "1"]
ADVANCE_PERIOD_OPTIONS = [
    "за один месяц",
    "за два месяца",
    "за три месяца",
]
VISITS_FREQUENCY_OPTIONS = [
    "Ежеквартально",
    "Ежемесячно",
    "Один раз в два месяца",
]
TERMINATION_PERIOD_OPTIONS = [
    "90 (девяносто) календарных дней",
    "60 (шестьдесят) календарных дней",
    "30 (тридцать) календарных дней",
]

IMAGES_DIR.mkdir(parents=True, exist_ok=True)
misc.configure_logging(LOGS_DIR, os.getenv("LOG_LEVEL", "INFO"))
log = logging.getLogger("max-bot")


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise misc.ConfigurationError(
            f"Не задана обязательная переменная окружения {name}"
        )
    return value


def parse_staff_ids(value: str) -> set[int]:
    try:
        result = {int(item.strip()) for item in value.split(",") if item.strip()}
    except ValueError as exc:
        raise misc.ConfigurationError(
            "MAX_STAFF_IDS должен содержать ID сотрудников через запятую"
        ) from exc
    if not result:
        raise misc.ConfigurationError("Список MAX_STAFF_IDS пуст")
    return result


TOKEN = require_env("MAX_BOT_TOKEN")
if TOKEN.casefold().startswith("bearer "):
    raise misc.ConfigurationError(
        "MAX_BOT_TOKEN должен содержать только токен, без префикса Bearer"
    )
STAFF = parse_staff_ids(require_env("MAX_STAFF_IDS"))

start_message_path = MESSAGES_DIR / "hello.txt"
try:
    START_MESSAGE = start_message_path.read_text(encoding="utf-8")
except OSError:
    START_MESSAGE = "Здравствуйте! Введите фамилию."
    log.warning("Не удалось прочитать приветствие | path=%s", start_message_path)

ms = misc.Miscellaneous(bot_dir=BOT_DIR, trust_env=False)
DIRECT_HTTP = requests.Session()
DIRECT_HTTP.trust_env = False

user_data: dict[int, dict[str, Any]] = {}


def open_word_document(filename: str | Path):
    return Package.open(filename).main_document_part.document


PartFactory.part_type_for[ms.DOCM_MAIN_CONTENT_TYPE] = DocumentPart
docxtpl_template.Document = open_word_document


class MaxAuthenticationError(RuntimeError):
    """MAX отклонил токен чат-бота."""


class DocumentDeliveryError(RuntimeError):
    """Договор сформирован, но MAX не смог доставить файл пользователю."""


class MaxAPI:
    def __init__(self, token: str):
        if not token:
            raise misc.ConfigurationError(
                "Не задана переменная окружения MAX_BOT_TOKEN"
            )
        self.session = requests.Session()
        # MAX всегда использует прямое соединение и не наследует HTTP(S)_PROXY.
        self.session.trust_env = False
        self.session.headers["Authorization"] = token

    def request(self, method: str, path: str, **kwargs) -> dict[str, Any]:
        normalized_path = "/" + path.lstrip("/")
        response = self.session.request(
            method,
            f"{API_URL}{normalized_path}",
            timeout=kwargs.pop("timeout", 40),
            **kwargs,
        )

        if response.status_code == 401:
            log.error(
                "MAX отклонил авторизацию | method=%s | path=%s | status=401",
                method,
                normalized_path,
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
                normalized_path,
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
        return self.request(
            "POST",
            "/messages",
            params={"user_id": user_id},
            json=body,
        )

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
        return {"type": "file", "payload": payload}


bot = MaxAPI(TOKEN)


def new_session(*, stage: str = "idle") -> dict[str, Any]:
    return {
        "cost": 0,
        "lastname": None,
        "ending": None,
        "complects": 0,
        "count_print": 0,
        "count_sending": 0,
        "stage": stage,
        "scenario_id": uuid4().hex,
    }


def session(user_id: int) -> dict[str, Any]:
    return user_data.setdefault(user_id, new_session())


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
    log.warning("Попытка доступа без разрешения | user_id=%s", user_id)
    try:
        bot.send(
            user_id,
            "❌ <b>Вам запрещён доступ к боту!</b> "
            "Обратитесь к техническому специалисту.",
        )
    except requests.RequestException:
        log.warning(
            "Не удалось отправить отказ в доступе | user_id=%s",
            user_id,
        )
    return False


def inline_keyboard(rows: list[list[tuple[str, str]]]) -> dict[str, Any]:
    return {
        "type": "inline_keyboard",
        "payload": {
            "buttons": [
                [
                    {
                        "type": "callback",
                        "text": text,
                        "payload": payload,
                    }
                    for text, payload in row
                ]
                for row in rows
            ]
        },
    }


def control_keyboard() -> dict[str, Any]:
    return inline_keyboard(
        [[
            (START_BUTTON_TEXT, "start"),
            (RESTART_BUTTON_TEXT, "restart"),
        ]]
    )


def contract_type_keyboard() -> dict[str, Any]:
    return inline_keyboard(
        [
            [
                ("ООО СПЕЦКОНС", "contract:specons"),
                ("ООО СПЕЦЭНЕРГОРАЗВИТИЕ", "contract:ser"),
            ],
            [(RESTART_BUTTON_TEXT, "restart")],
        ]
    )


def source_keyboard() -> dict[str, Any]:
    return inline_keyboard(
        [
            [("🧾 Документ", "doc"), ("🖼️ Картинка", "pic")],
            [("📄 PDF-скан", "pdf"), ("📢 Сообщением", "mes")],
            [(RESTART_BUTTON_TEXT, "restart")],
        ]
    )


def options_keyboard(prefix: str, options: list[str]) -> dict[str, Any]:
    rows = [[(option, f"{prefix}:{index}")] for index, option in enumerate(options)]
    rows.append([(RESTART_BUTTON_TEXT, "restart")])
    return inline_keyboard(rows)


def show_start_menu(user_id: int, text: str | None = None) -> None:
    if not allowed(user_id):
        return
    user_data[user_id] = new_session(stage="idle")
    bot.send(
        user_id,
        text or "Нажмите кнопку, чтобы сформировать новый проект договора.",
        [control_keyboard()],
    )


def landing(user_id: int) -> None:
    if not allowed(user_id):
        return
    user_data[user_id] = new_session(stage="contract_type")
    log.info("Сценарий запущен | user_id=%s", user_id)
    bot.send(
        user_id,
        "❔ Выберите тип договора",
        [contract_type_keyboard()],
    )


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
                if isinstance(candidate, str) and candidate.startswith(
                    ("http://", "https://")
                ):
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
        unquote(disposition_match.group(1).strip())
        if disposition_match
        else None
    )

    document_suffixes = set(ms.ALLOWED_EXTENSIONS)
    image_suffixes = {
        ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"
    }
    accepted_suffixes = image_suffixes if expected == "image" else document_suffixes

    for candidate in (
        original_name,
        disposition_name,
        unquote(urlparse(url).path),
    ):
        suffix = Path(candidate or "").suffix.lower()
        if suffix in accepted_suffixes:
            return suffix

    content_type = response.headers.get("Content-Type", "").split(";", 1)[0].lower()
    known_mime_types = {
        "application/pdf": ".pdf",
        "application/msword": ".doc",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
        "application/vnd.ms-word.document.macroenabled.12": ".docm",
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
        "Отправьте его повторно как DOC, DOCX, DOCM или PDF."
    )


def download_attachment(message: dict[str, Any], expected: str) -> Path:
    attachments = (message.get("body") or {}).get("attachments") or []
    selected = next((item for item in attachments if item.get("type") == expected), None)
    if selected is None:
        expected_name = "изображение" if expected == "image" else "файл"
        raise misc.DocumentExtractionError(
            f"Сообщение не содержит ожидаемое {expected_name}."
        )

    url = attachment_url(selected)
    if not url:
        raise misc.DocumentExtractionError("MAX не прислал URL вложения")

    response = DIRECT_HTTP.get(url, timeout=180)
    response.raise_for_status()
    if len(response.content) > ms.MAX_FILE_BYTES:
        raise misc.DocumentExtractionError(
            f"Файл больше {ms.MAX_FILE_BYTES // (1024 * 1024)} МБ"
        )

    suffix = detect_attachment_suffix(selected, url, response, expected)
    if expected == "file" and suffix not in ms.ALLOWED_EXTENSIONS:
        raise misc.DocumentExtractionError(
            f"Формат {suffix} не поддерживается. Отправьте DOC, DOCX, DOCM или PDF."
        )

    target_dir = IMAGES_DIR if expected == "image" else ms.USERS_DOCS_DIR
    target = target_dir / f"max_{uuid4().hex}{suffix}"
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


def report_processing_error(user_id: int, error: Exception) -> None:
    if isinstance(error, misc.DocumentExtractionError):
        log.warning(
            "Входные данные отклонены | user_id=%s | reason=%s",
            user_id,
            error,
        )
    elif isinstance(error, DocumentDeliveryError):
        log.error(
            "Готовый договор не доставлен | user_id=%s | reason=%s",
            user_id,
            error,
        )
    else:
        log.exception(
            "Ошибка обработки запроса | user_id=%s | error_type=%s",
            user_id,
            type(error).__name__,
        )

    if isinstance(
        error,
        (misc.AIServiceError, misc.DocumentExtractionError, DocumentDeliveryError),
    ):
        text = f"❌ {error}"
    else:
        text = (
            "❌ Не удалось обработать документ. "
            "Проверьте формат и содержимое либо повторите попытку позже."
        )
    bot.send(user_id, text, [control_keyboard()])


def report_missing_company_details(
    user_id: int,
    error: misc.MissingCompanyDetailsError,
) -> None:
    log.warning(
        "Формирование договора остановлено: не хватает реквизитов | "
        "user_id=%s | missing=%s",
        user_id,
        ", ".join(error.missing_fields),
    )
    fields = "\n".join(f"• <b>{field}</b>" for field in error.missing_fields)
    bot.send(
        user_id,
        "⚠️ Договор пока нельзя сформировать. Не хватает обязательных "
        f"реквизитов:\n\n{fields}\n\n"
        "Уточните их у заказчика и отправьте исправленные реквизиты повторно.",
        [control_keyboard()],
    )


def send_contract_document(user_id: int, document_path: Path) -> None:
    """Загружает готовый договор в MAX и ждёт готовности вложения."""
    try:
        attachment = bot.upload_file(document_path)
    except requests.RequestException as exc:
        raise DocumentDeliveryError(
            "Договор сформирован, но MAX не смог загрузить файл. "
            "Проверьте интернет и повторите попытку позже."
        ) from exc

    delays = (1, 2, 4, 8, 12)
    last_error: Exception | None = None

    for attempt, delay in enumerate(delays, start=1):
        try:
            bot.send(user_id, "Проект договора", [attachment])
            log.info(
                "Готовый договор передан в MAX | user_id=%s | attempt=%d/%d | bytes=%d",
                user_id,
                attempt,
                len(delays),
                document_path.stat().st_size,
            )
            return
        except requests.HTTPError as exc:
            last_error = exc
            response = exc.response
            try:
                error = response.json() if response is not None else {}
            except ValueError:
                error = {}

            if error.get("code") != "attachment.not.ready":
                raise DocumentDeliveryError(
                    "Договор сформирован, но MAX отклонил отправку файла. "
                    "Повторите попытку позже."
                ) from exc

            log.info(
                "Файл MAX ещё обрабатывается | user_id=%s | attempt=%d/%d | delay=%ss",
                user_id,
                attempt,
                len(delays),
                delay,
            )
            if attempt < len(delays):
                time.sleep(delay)
        except requests.RequestException as exc:
            last_error = exc
            log.warning(
                "Сбой отправки договора в MAX | user_id=%s | attempt=%d/%d | error=%s",
                user_id,
                attempt,
                len(delays),
                type(exc).__name__,
            )
            if attempt < len(delays):
                time.sleep(delay)

    raise DocumentDeliveryError(
        "Договор сформирован, но MAX не подготовил или не принял файл. "
        "Повторите попытку позже."
    ) from last_error


def finish_document(
    user_id: int,
    company_data: list[str | None],
    source_path: Path | None,
    scenario_id: str | None = None,
) -> bool:
    local_doc: Path | None = None
    try:
        if (
            scenario_id
            and user_data.get(user_id, {}).get("scenario_id") != scenario_id
        ):
            log.info("Отменена устаревшая обработка | user_id=%s", user_id)
            return False

        data = session(user_id)
        ms.validate_company_data(company_data)
        bot.send(
            user_id,
            "🧾 Реквизиты извлечены. Рассчитываю суммы и заполняю шаблон…",
            [control_keyboard()],
        )
        number_contract = ms.get_bot_doc_num(user_data, user_id)

        if data.get("contract_type") == "ser":
            ser_fields = data.get("ser_fields", {})
            months_count = ser_fields.get("months_count", "")
            cost_month = ser_fields.get("cost_month", "")
            try:
                cost_total_value = int(months_count) * int(cost_month)
            except (ValueError, TypeError):
                cost_total_value = 0
            texted_total = ms.integer_texted(cost_total_value)
            local_doc = Path(
                ms.bot_insert_req_ser(
                    user_data,
                    user_id,
                    company_data,
                    number_contract,
                    texted_total,
                    source_path,
                )
            )
        else:
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
                    source_path,
                )
            )

        if (
            scenario_id
            and user_data.get(user_id, {}).get("scenario_id") != scenario_id
        ):
            log.info(
                "Отправка устаревшего договора отменена | user_id=%s",
                user_id,
            )
            return False

        try:
            bot.send(user_id, "📤 Договор готов. Отправляю файл…")
        except requests.RequestException:
            log.warning(
                "Не удалось отправить статус перед загрузкой договора | user_id=%s",
                user_id,
            )

        send_contract_document(user_id, local_doc)
        log.info("Договор отправлен | user_id=%s", user_id)
        return True
    finally:
        ms.delete_local_doc(local_doc)


def ask_source(user_id: int) -> None:
    session(user_id)["stage"] = "choice"
    bot.send(
        user_id,
        "❔ Выберите источник реквизитов",
        [source_keyboard()],
    )


def begin_ser_options(user_id: int) -> None:
    session(user_id)["stage"] = "ser_months"
    bot.send(
        user_id,
        "❔ Выберите количество месяцев обслуживания",
        [options_keyboard("months", MONTHS_COUNT_OPTIONS)],
    )


def handle_text_stage(user_id: int, text: str) -> bool:
    data = session(user_id)
    stage = data["stage"]

    if stage == "name":
        if len(text) < 2 or not any(char.isalpha() for char in text):
            bot.send(
                user_id,
                "❌ Фамилия должна содержать не менее двух букв.",
                [control_keyboard()],
            )
            return True

        data["lastname"] = text[:2].upper()
        ending = datetime.now() + relativedelta(years=1)
        data["ending"] = (
            f"{ending.day} {ms.GENITIVUS[ending.month]} {ending.year} года"
        )
        log.info("Фамилия принята | user_id=%s", user_id)
        bot.send(
            user_id,
            f"✅ Фамилия принята. Код документов: <b>{data['lastname']}</b>",
            [control_keyboard()],
        )

        if data.get("contract_type") == "ser":
            data["ser_step"] = 0
            data["stage"] = "ser_text"
            bot.send(user_id, SER_QUESTIONS[0][1], [control_keyboard()])
        else:
            data["stage"] = "cost"
            bot.send(user_id, "❔ Введите стоимость разовой услуги", [control_keyboard()])
        return True

    if stage == "cost":
        if not text.isdigit():
            bot.send(
                user_id,
                "❌ Введите целое неотрицательное число.",
                [control_keyboard()],
            )
            return True
        value = int(text)
        data["cost"] = value
        data["stage"] = "complects"
        bot.send(
            user_id,
            f"✅ Стоимость принята: <b>{value:,} ₽</b>".replace(",", " "),
            [control_keyboard()],
        )
        bot.send(user_id, "❔ Введите количество комплектов", [control_keyboard()])
        return True

    if stage == "complects":
        if not text.isdigit():
            bot.send(
                user_id,
                "❌ Введите целое неотрицательное число.",
                [control_keyboard()],
            )
            return True
        amount = int(text)
        data["complects"] = amount
        data["count_print"] = 3900 * amount
        data["count_sending"] = data["count_print"] + 1000 if amount else 0
        bot.send(
            user_id,
            f"✅ Количество комплектов принято: <b>{amount}</b>",
            [control_keyboard()],
        )
        ask_source(user_id)
        return True

    if stage == "ser_text":
        step_index = int(data.get("ser_step", 0))
        if step_index >= len(SER_QUESTIONS):
            begin_ser_options(user_id)
            return True
        if not text:
            bot.send(
                user_id,
                "❌ Значение не может быть пустым.",
                [control_keyboard()],
            )
            return True

        key, _ = SER_QUESTIONS[step_index]
        data.setdefault("ser_fields", {})[key] = text
        step_index += 1
        data["ser_step"] = step_index
        if step_index >= len(SER_QUESTIONS):
            begin_ser_options(user_id)
        else:
            bot.send(
                user_id,
                SER_QUESTIONS[step_index][1],
                [control_keyboard()],
            )
        return True

    return False


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

    if stage == "contract_type":
        bot.send(
            user_id,
            "❔ Сначала выберите тип договора кнопкой ниже.",
            [contract_type_keyboard()],
        )
        return

    if handle_text_stage(user_id, text):
        return

    if stage in {"ser_months", "ser_advance", "ser_visits", "ser_termination", "choice"}:
        bot.send(
            user_id,
            "❔ На этом шаге используйте одну из кнопок в сообщении выше.",
            [control_keyboard()],
        )
        return

    source_path: Path | None = None
    try:
        scenario_id = data.get("scenario_id")

        if stage == "doc":
            source_path = download_attachment(message, "file")
            bot.send(
                user_id,
                "✅ Документ получен. Извлекаю текст и передаю реквизиты ИИ — "
                "это может занять до минуты.",
                [control_keyboard()],
            )
            company_data = ms.sent_doc_to_ai(source_path)

        elif stage == "pdf":
            source_path = download_attachment(message, "file")
            if source_path.suffix.lower() != ".pdf":
                raise misc.DocumentExtractionError("Отправьте PDF-файл.")
            bot.send(
                user_id,
                "✅ PDF-скан получен. Распознаю страницы и извлекаю реквизиты — "
                "пожалуйста, не отправляйте файл повторно.",
                [control_keyboard()],
            )
            company_data = ms.sent_pdf_scan_to_ai(source_path)

        elif stage == "pic":
            source_path = download_attachment(message, "image")
            bot.send(
                user_id,
                "✅ Изображение получено. Распознаю текст и извлекаю реквизиты — "
                "это может занять до минуты.",
                [control_keyboard()],
            )
            # Новый misc сам нормализует изображение перед OCR.
            company_data = ms.sent_image_to_ai(source_path)

        elif stage == "mes":
            if not text:
                raise misc.DocumentExtractionError("Отправьте текстовое сообщение.")
            bot.send(
                user_id,
                "✅ Сообщение принято. Извлекаю реквизиты…",
                [control_keyboard()],
            )
            company_data = ms.sent_message_to_ai(text)

        else:
            show_start_menu(user_id)
            return

        if finish_document(
            user_id,
            company_data,
            source_path,
            scenario_id,
        ):
            show_start_menu(
                user_id,
                "✅ Договор отправлен. Можно начать новый сценарий.",
            )

    except misc.MissingCompanyDetailsError as exc:
        report_missing_company_details(user_id, exc)
    except Exception as exc:
        report_processing_error(user_id, exc)
    finally:
        ms.delete_local_doc(source_path)


def callback_user_id(update: dict[str, Any]) -> int:
    callback = update.get("callback") or {}
    message = update.get("message") or callback.get("message") or {}
    user = callback.get("user") or update.get("user") or message.get("sender") or {}
    value = user.get("user_id")
    if value is None:
        raise ValueError("В callback MAX отсутствует user_id")
    return int(value)


def parse_option_payload(payload: str, prefix: str, options: list[str]) -> str | None:
    expected_prefix = f"{prefix}:"
    if not payload.startswith(expected_prefix):
        return None
    try:
        index = int(payload.removeprefix(expected_prefix))
    except ValueError:
        return None
    if index < 0 or index >= len(options):
        return None
    return options[index]


def handle_callback(update: dict[str, Any]) -> None:
    callback = update.get("callback") or {}
    user_id = callback_user_id(update)
    if not allowed(user_id):
        return

    payload = str(callback.get("payload") or "")
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

    data = session(user_id)

    if payload in {"contract:specons", "contract:ser"}:
        if data.get("stage") != "contract_type":
            show_start_menu(
                user_id,
                "⚠️ Этот выбор относится к завершённому сценарию. Нажмите «Начать».",
            )
            return
        is_ser = payload == "contract:ser"
        data["contract_type"] = "ser" if is_ser else "specons"
        data["stage"] = "name"
        contract_label = "ООО СПЕЦЭНЕРГОРАЗВИТИЕ" if is_ser else "ООО СПЕЦКОНС"
        bot.send(
            user_id,
            f"✅ Выбран тип договора: <b>{contract_label}</b>",
            [control_keyboard()],
        )
        bot.send(user_id, START_MESSAGE, [control_keyboard()])
        return

    months = parse_option_payload(payload, "months", MONTHS_COUNT_OPTIONS)
    if months is not None:
        if data.get("stage") != "ser_months":
            return
        data.setdefault("ser_fields", {})["months_count"] = months
        data["stage"] = "ser_advance"
        bot.send(
            user_id,
            f"✅ Количество месяцев: <b>{months}</b>",
            [control_keyboard()],
        )
        bot.send(
            user_id,
            "❔ Выберите аванс за период",
            [options_keyboard("advance", ADVANCE_PERIOD_OPTIONS)],
        )
        return

    advance = parse_option_payload(payload, "advance", ADVANCE_PERIOD_OPTIONS)
    if advance is not None:
        if data.get("stage") != "ser_advance":
            return
        data.setdefault("ser_fields", {})["advance_period"] = advance
        data["stage"] = "ser_visits"
        bot.send(
            user_id,
            f"✅ Аванс за период: <b>{advance}</b>",
            [control_keyboard()],
        )
        bot.send(
            user_id,
            "❔ Выберите периодичность обходов",
            [options_keyboard("visits", VISITS_FREQUENCY_OPTIONS)],
        )
        return

    visits = parse_option_payload(payload, "visits", VISITS_FREQUENCY_OPTIONS)
    if visits is not None:
        if data.get("stage") != "ser_visits":
            return
        data.setdefault("ser_fields", {})["visits_frequency"] = visits
        data["stage"] = "ser_termination"
        bot.send(
            user_id,
            f"✅ Периодичность обходов: <b>{visits}</b>",
            [control_keyboard()],
        )
        bot.send(
            user_id,
            "❔ Выберите срок расторжения договора",
            [options_keyboard("termination", TERMINATION_PERIOD_OPTIONS)],
        )
        return

    termination = parse_option_payload(
        payload,
        "termination",
        TERMINATION_PERIOD_OPTIONS,
    )
    if termination is not None:
        if data.get("stage") != "ser_termination":
            return
        data.setdefault("ser_fields", {})["termination_period"] = termination
        bot.send(
            user_id,
            f"✅ Срок расторжения: <b>{termination}</b>",
            [control_keyboard()],
        )
        ask_source(user_id)
        return

    prompts = {
        "pic": ("pic", "<b>▶️ Вышлите картинку</b>"),
        "doc": ("doc", "<b>▶️ Вышлите документ</b>"),
        "pdf": ("pdf", "<b>▶️ Вышлите PDF-скан документа</b>"),
        "mes": ("mes", "<b>▶️ Отправьте или перешлите сообщение</b>"),
    }
    if payload in prompts:
        if data.get("stage") != "choice":
            show_start_menu(
                user_id,
                "⚠️ Предыдущий сценарий недоступен. Нажмите «Начать».",
            )
            return
        data["stage"] = prompts[payload][0]
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
        return

    log.warning("Неизвестный callback MAX | user_id=%s | payload=%s", user_id, payload)


def handle_update(update: dict[str, Any]) -> None:
    update_type = update.get("update_type")

    if update_type == "bot_started":
        user = update.get("user") or {}
        user_id = int(user["user_id"])
        show_start_menu(user_id)
        return

    if update_type == "message_callback":
        handle_callback(update)
        # POST /answers нужен только для изменения исходного сообщения или
        # показа одноразового уведомления. Для обычного перехода сценария
        # отвечать на callback отдельным запросом не требуется.
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


def notify_staff_after_restart() -> None:
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


def main() -> None:
    marker: int | None = None
    try:
        identity = bot.me()
    except MaxAuthenticationError as exc:
        log.critical("MAX-бот остановлен: %s", exc)
        return
    except requests.RequestException:
        log.exception("Не удалось проверить авторизацию MAX")
        return

    log.info(
        "Авторизация MAX подтверждена | bot_id=%s | username=%s",
        identity.get("user_id"),
        identity.get("username"),
    )
    log.info("MAX-бот запускается")
    notify_staff_after_restart()

    while True:
        try:
            page = bot.updates(marker)
            # Указатель страницы сохраняется отдельно: ошибка одного события
            # не должна приводить к бесконечной повторной обработке всей страницы.
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
            log.exception("Необработанная ошибка polling MAX; повтор через 3 секунды")
            time.sleep(3)


if __name__ == "__main__":
    main()
