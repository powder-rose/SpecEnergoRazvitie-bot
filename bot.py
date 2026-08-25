"""Telegram-интерфейс генератора договоров."""

from __future__ import annotations

import logging
import os
import threading
import time
from datetime import datetime
from functools import wraps
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlparse
from uuid import uuid4

import docxtpl.template as docxtpl_template
import telebot
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv
from docx.opc.part import PartFactory
from docx.package import Package
from docx.parts.document import DocumentPart
from telebot import apihelper, types

from core import misc as misc

BOT_DIR = Path(__file__).resolve().parent

load_dotenv(BOT_DIR / ".env")

MESSAGES_DIR = BOT_DIR / "messages"
IMAGES_DIR = BOT_DIR / "images"
LOGS_DIR = BOT_DIR / "logs"
START_BUTTON_TEXT = "🚀 Начать"
RESTART_BUTTON_TEXT = "🔄 Перезапустить"
CONTROL_WORDS = {
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
IMAGES_DIR.mkdir(parents=True, exist_ok=True)

misc.configure_logging(LOGS_DIR, os.getenv("LOG_LEVEL", "INFO"))
LOGGER = logging.getLogger(__name__)

# Yandex API работает напрямую; прокси ниже применяется только к Telegram API.
ms = misc.Miscellaneous(bot_dir=BOT_DIR, trust_env=False)


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
            "TELEGRAM_STAFF_IDS должен содержать Telegram ID через запятую"
        ) from exc
    if not result:
        raise misc.ConfigurationError("Список TELEGRAM_STAFF_IDS пуст")
    return result


start_message = (MESSAGES_DIR / "hello.txt").read_text(encoding="utf-8")
staff = parse_staff_ids(require_env("TELEGRAM_STAFF_IDS"))
# Долгоживущая HTTP-сессия после простоя иногда обрывает первую отправку.
apihelper.SESSION_TIME_TO_LIVE = 5 * 60
telegram_proxy_url = os.getenv("TELEGRAM_PROXY_URL", "").strip()
if telegram_proxy_url:
    parsed_proxy = urlparse(telegram_proxy_url)
    if parsed_proxy.scheme not in {"http", "https", "socks5", "socks5h"}:
        raise misc.ConfigurationError(
            "TELEGRAM_PROXY_URL должен начинаться с http://, https://, "
            "socks5:// или socks5h://"
        )
    if not parsed_proxy.hostname:
        raise misc.ConfigurationError("В TELEGRAM_PROXY_URL не указан адрес прокси")
    if parsed_proxy.scheme in {"socks5", "socks5h"}:
        try:
            import socks  # noqa: F401
        except ImportError as exc:
            raise misc.ConfigurationError(
                "Для SOCKS-прокси установите зависимость: pip install requests[socks]"
            ) from exc
    apihelper.proxy = {"https": telegram_proxy_url}
    LOGGER.info(
        "Для Telegram включён отдельный прокси | scheme=%s | host=%s",
        parsed_proxy.scheme,
        parsed_proxy.hostname,
    )
bot = telebot.TeleBot(require_env("TELEGRAM_BOT_TOKEN"), parse_mode="HTML")
user_data: dict[int, dict[str, Any]] = {}
CHAT_LOCKS: dict[int, threading.RLock] = {}
CHAT_LOCKS_GUARD = threading.Lock()
RECENT_ACTIONS: dict[int, tuple[str, float]] = {}
ACTION_DEBOUNCE_SECONDS = 1.5
TELEGRAM_DOCUMENT_SEND_TIMEOUT = 300
TELEGRAM_DOCUMENT_SEND_ATTEMPTS = 3


PartFactory.part_type_for[ms.DOCM_MAIN_CONTENT_TYPE] = DocumentPart


def open_word_document(filename: str):
    package = Package.open(filename)
    return package.main_document_part.document


docxtpl_template.Document = open_word_document


class DocumentDeliveryError(RuntimeError):
    """Договор сформирован, но не доставлен пользователю."""


def employees_only(handler: Callable):
    @wraps(handler)
    def wrapper(message, *args, **kwargs):
        user_id = getattr(message.from_user, "id", None)
        if user_id not in staff:
            LOGGER.warning("Попытка доступа без разрешения | user_id=%s", user_id)
            bot.reply_to(
                message,
                "❌ <b>Вам запрещён доступ к боту!</b> "
                "Обратитесь к техническому специалисту.",
            )
            return None
        return handler(message, *args, **kwargs)

    return wrapper


def safe_handler(handler: Callable):
    """Ловит любые непредвиденные ошибки в хендлере, чтобы не уронить бота."""
    @wraps(handler)
    def wrapper(update, *args, **kwargs):
        try:
            return handler(update, *args, **kwargs)
        except Exception:
            message = getattr(update, "message", update)
            chat_id = getattr(getattr(message, "chat", None), "id", None)
            user_id = getattr(getattr(update, "from_user", None), "id", None)
            LOGGER.exception(
                "Необработанная ошибка в хендлере | handler=%s | user_id=%s",
                handler.__name__,
                user_id,
            )
            if chat_id is not None:
                try:
                    bot.send_message(
                        chat_id,
                        "❌ Произошла непредвиденная ошибка. Попробуйте ещё раз "
                        "или нажмите «Начать».",
                        reply_markup=control_keyboard(),
                    )
                except Exception:
                    LOGGER.warning(
                        "Не удалось отправить сообщение об ошибке | chat_id=%s",
                        chat_id,
                    )
            return None
    return wrapper


def chat_lock(chat_id: int) -> threading.RLock:
    """Возвращает единственную блокировку для конкретного Telegram-чата."""
    with CHAT_LOCKS_GUARD:
        return CHAT_LOCKS.setdefault(chat_id, threading.RLock())


def is_repeated_action(chat_id: int, action: str) -> bool:
    """Подавляет двойной клик, не блокируя другое действие пользователя."""
    now = time.monotonic()
    previous_action, previous_at = RECENT_ACTIONS.get(chat_id, ("", 0.0))
    RECENT_ACTIONS[chat_id] = (action, now)
    return (
        previous_action == action
        and now - previous_at < ACTION_DEBOUNCE_SECONDS
    )


def replace_next_step(message, handler: Callable) -> None:
    """Атомарно оставляет у чата ровно один next-step обработчик."""
    with chat_lock(message.chat.id):
        bot.clear_step_handler_by_chat_id(message.chat.id)
        bot.register_next_step_handler(message, handler)


def register_retry(message, handler: Callable, text: str) -> None:
    bot.send_message(message.chat.id, text)
    replace_next_step(message, handler)


def control_keyboard() -> types.ReplyKeyboardMarkup:
    markup = types.ReplyKeyboardMarkup(
        resize_keyboard=True,
        one_time_keyboard=False,
    )
    markup.row(
        types.KeyboardButton(START_BUTTON_TEXT),
        types.KeyboardButton(RESTART_BUTTON_TEXT),
    )
    return markup


def show_start_button(message, user_id: int | None = None) -> None:
    """Возвращает пользователя к безопасной точке начала сценария."""
    with chat_lock(message.chat.id):
        bot.clear_step_handler_by_chat_id(message.chat.id)
        resolved_user_id = user_id or message.from_user.id
        user_data.pop(resolved_user_id, None)
        bot.send_message(
            message.chat.id,
            "Нажмите кнопку, чтобы сформировать новый проект договора.",
            reply_markup=control_keyboard(),
        )


def is_control_message(message) -> bool:
    text = (getattr(message, "text", None) or "").strip().casefold()
    return text in CONTROL_WORDS


def recover_if_requested(message) -> bool:
    """Перехватывает Начать/Перезапустить даже внутри next_step_handler."""
    if not is_control_message(message):
        return False
    landing(message)
    return True


def ensure_active_session(message) -> bool:
    """Восстанавливает управление после перезапуска процесса."""
    if message.from_user.id in user_data:
        return True
    bot.send_message(
        message.chat.id,
        "⚠️ Сценарий был прерван перезапуском бота. "
        "Нажмите «Начать», чтобы продолжить с начала.",
        reply_markup=control_keyboard(),
    )
    return False


def report_processing_error(message, error: Exception) -> None:
    if isinstance(error, misc.DocumentExtractionError):
        LOGGER.warning(
            "Входные данные отклонены | user_id=%s | reason=%s",
            getattr(message.from_user, "id", None),
            error,
        )
    elif isinstance(error, DocumentDeliveryError):
        LOGGER.error(
            "Готовый договор не доставлен | user_id=%s | reason=%s",
            getattr(message.from_user, "id", None),
            error,
        )
    else:
        LOGGER.exception(
            "Ошибка обработки запроса | user_id=%s | error_type=%s",
            getattr(message.from_user, "id", None),
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
    bot.send_message(message.chat.id, text, reply_markup=control_keyboard())


def report_missing_company_details(
    message,
    error: misc.MissingCompanyDetailsError,
    retry_handler: Callable,
) -> None:
    LOGGER.warning(
        "Формирование договора остановлено: не хватает реквизитов | "
        "user_id=%s | missing=%s",
        getattr(message.from_user, "id", None),
        ", ".join(error.missing_fields),
    )
    fields = "\n".join(f"• <b>{field}</b>" for field in error.missing_fields)
    bot.send_message(
        message.chat.id,
        "⚠️ Договор пока нельзя сформировать. Не хватает обязательных "
        f"реквизитов:\n\n{fields}\n\n"
        "Уточните их у заказчика и отправьте исправленные реквизиты повторно.",
        reply_markup=control_keyboard(),
    )
    replace_next_step(message, retry_handler)


def send_contract_document(chat_id: int, document_path: Path) -> None:
    """Отправляет уже готовый договор с увеличенным тайм-аутом и повторами."""
    last_error: Exception | None = None

    for attempt in range(1, TELEGRAM_DOCUMENT_SEND_ATTEMPTS + 1):
        try:
            with document_path.open("rb") as document:
                bot.send_document(
                    chat_id=chat_id,
                    document=document,
                    caption="Проект договора",
                    timeout=TELEGRAM_DOCUMENT_SEND_TIMEOUT,
                )
            LOGGER.info(
                "Готовый договор передан в Telegram | chat_id=%s | "
                "attempt=%d/%d | bytes=%d",
                chat_id,
                attempt,
                TELEGRAM_DOCUMENT_SEND_ATTEMPTS,
                document_path.stat().st_size,
            )
            return
        except Exception as exc:
            last_error = exc
            error_code = getattr(exc, "error_code", None)
            retryable = (
                error_code is None
                or error_code == 429
                or (isinstance(error_code, int) and error_code >= 500)
            )
            LOGGER.warning(
                "Не удалось отправить готовый договор | chat_id=%s | "
                "attempt=%d/%d | timeout=%ds | error_type=%s | "
                "error_code=%s | retryable=%s",
                chat_id,
                attempt,
                TELEGRAM_DOCUMENT_SEND_ATTEMPTS,
                TELEGRAM_DOCUMENT_SEND_TIMEOUT,
                type(exc).__name__,
                error_code,
                retryable,
            )
            if not retryable or attempt >= TELEGRAM_DOCUMENT_SEND_ATTEMPTS:
                break

            try:
                bot.send_message(
                    chat_id,
                    "⏳ Договор уже сформирован, но соединение с Telegram "
                    f"прервалось. Повторяю отправку ({attempt + 1}/"
                    f"{TELEGRAM_DOCUMENT_SEND_ATTEMPTS})…",
                    reply_markup=control_keyboard(),
                    timeout=60,
                )
            except Exception:
                LOGGER.warning(
                    "Не удалось отправить уведомление о повторе | chat_id=%s",
                    chat_id,
                )
            time.sleep(10 * attempt)

    raise DocumentDeliveryError(
        "Договор сформирован, но Telegram не смог принять файл. "
        "Проверьте VPN/интернет и повторите попытку позже."
    ) from last_error


@bot.message_handler(commands=["start"])
@safe_handler
@employees_only
def welcome(message):
    with chat_lock(message.chat.id):
        if is_repeated_action(message.chat.id, "control:/start"):
            return
    LOGGER.info(
        "Открыто начальное меню | user_id=%s",
        message.from_user.id,
    )
    show_start_button(message)


@bot.message_handler(
    func=lambda message: bool(
        message.text
        and message.text.strip().casefold()
        in CONTROL_WORDS
    )
)
@safe_handler
@employees_only
def landing(message):
    user_id = message.from_user.id
    control_text = (message.text or START_BUTTON_TEXT).strip().casefold()
    with chat_lock(message.chat.id):
        if is_repeated_action(
            message.chat.id,
            f"control:{control_text}",
        ):
            LOGGER.debug(
                "Повторное нажатие управляющей кнопки подавлено | user_id=%s",
                user_id,
            )
            return
        bot.clear_step_handler_by_chat_id(message.chat.id)
        user_data[user_id] = {
            "cost": 0,
            "lastname": None,
            "ending": None,
            "complects": 0,
            "count_print": 0,
            "count_sending": 0,
            "scenario_id": uuid4().hex,
        }
        LOGGER.info("Сценарий запущен | user_id=%s", user_id)
        contract_type_markup = types.InlineKeyboardMarkup()
        contract_type_markup.row(
            types.InlineKeyboardButton("ООО СПЕЦКОНС", callback_data="contract_specons"),
            types.InlineKeyboardButton("ООО СПЕЦЭНЕРГОРАЗВИТИЕ", callback_data="contract_ser"),
        )
        bot.send_message(
            message.chat.id,
            "❔ Выберите тип договора",
            reply_markup=contract_type_markup,
        )


@bot.callback_query_handler(func=lambda callback: callback.data in ("contract_specons", "contract_ser"))
@safe_handler
def choose_contract_type(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    is_ser = callback.data == "contract_ser"
    user_data[user_id]["contract_type"] = "ser" if is_ser else "specons"

    contract_label = "ООО СПЕЦЭНЕРГОРАЗВИТИЕ" if is_ser else "ООО СПЕЦКОНС"
    bot.send_message(
        chat_id,
        f"✅ Выбран тип договора: <b>{contract_label}</b>",
        reply_markup=control_keyboard(),
    )
    bot.send_message(chat_id, start_message, reply_markup=control_keyboard())
    ask_surname(callback.message)


# ---------------------------------------------------------------------------
# Фамилия: готовые варианты кнопками + возможность вписать свою
# ---------------------------------------------------------------------------

SURNAME_OPTIONS = ["Яшенина", "Сиротская", "Власова", "Лосева"]


def apply_surname(message, name: str) -> None:
    """Проверяет и сохраняет фамилию (код документов — первые 2 буквы),
    затем ведёт сценарий дальше в зависимости от типа договора."""
    user_id = message.from_user.id
    name = (name or "").strip()
    if len(name) < 2 or not any(char.isalpha() for char in name):
        LOGGER.warning("Некорректная фамилия | user_id=%s", user_id)
        register_retry(
            message,
            yourname,
            "❌ Фамилия должна содержать не менее двух букв.",
        )
        return

    user_data[user_id]["lastname"] = name[:2].upper()
    now_plus = datetime.now() + relativedelta(years=1)
    user_data[user_id]["ending"] = (
        f"{now_plus.day} {ms.GENITIVUS[now_plus.month]} "
        f"{now_plus.year} года"
    )
    LOGGER.info("Фамилия принята | user_id=%s", user_id)
    bot.send_message(
        message.chat.id,
        f"✅ Фамилия принята. Код документов: "
        f"<b>{user_data[user_id]['lastname']}</b>",
        reply_markup=control_keyboard(),
    )
    if user_data[user_id].get("contract_type") == "ser":
        ask_cost_month(message)
    else:
        register_retry(message, costs, "❔ Введите стоимость разовой услуги")


@safe_handler
def yourname(message):
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    apply_surname(message, message.text or "")


def ask_surname(message) -> None:
    markup = types.InlineKeyboardMarkup()
    for option in SURNAME_OPTIONS:
        markup.add(types.InlineKeyboardButton(option, callback_data=f"surname_{option}"))
    markup.add(types.InlineKeyboardButton("✏️ Своя фамилия", callback_data="surname_custom"))
    bot.send_message(message.chat.id, "❔ Выберите фамилию", reply_markup=markup)


@bot.callback_query_handler(func=lambda callback: callback.data.startswith("surname_"))
@safe_handler
def choose_surname(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    option = callback.data.removeprefix("surname_")

    if option == "custom":
        bot.send_message(chat_id, "❔ Введите фамилию", reply_markup=control_keyboard())
        replace_next_step(callback.message, yourname)
        return

    apply_surname(callback.message, option)


# ---------------------------------------------------------------------------
# Сценарий для ООО СПЕЦЭНЕРГОРАЗВИТИЕ
# ---------------------------------------------------------------------------

SER_QUESTIONS = [
    ("object_address", "❔ Введите адрес объекта"),
    ("object_name", "❔ Введите наименование объекта"),
    (
        "service_period",
        "❔ Введите период тех. обслуживания "
        "(например: май 2024, июнь 2024, июль 2024)",
    ),
    ("email", "❔ Введите электронную почту заказчика"),
]

COST_MONTH_OPTIONS = ["4200", "5200", "8300", "9500", "19900"]

MONTHS_COUNT_OPTIONS = ["3", "2", "1"]

TERMINATION_PERIOD_OPTIONS = [
    "90 (девяносто) календарных дней",
    "60 (шестьдесят) календарных дней",
    "30 (тридцать) календарных дней",
]

VISITS_FREQUENCY_OPTIONS = [
    "Ежеквартально",
    "Ежемесячно",
    "Один раз в два месяца",
]

ADVANCE_PERIOD_OPTIONS = [
    "за один месяц",
    "за два месяца",
    "за три месяца",
]


def ask_ser_field(message, step_index: int) -> None:
    user_id = message.from_user.id
    if step_index >= len(SER_QUESTIONS):
        finish_ser_fields(message)
        return

    key, prompt = SER_QUESTIONS[step_index]

    def handler(msg, index=step_index, field_key=key):
        if recover_if_requested(msg) or not ensure_active_session(msg):
            return
        answer = (msg.text or "").strip()
        if not answer:
            register_retry(
                msg,
                lambda m, i=index: ask_ser_field(m, i),
                "❌ Значение не может быть пустым.",
            )
            return
        user_data[msg.from_user.id].setdefault("ser_fields", {})[field_key] = answer
        ask_ser_field(msg, index + 1)

    sent = bot.send_message(message.chat.id, prompt, reply_markup=control_keyboard())
    replace_next_step(sent, handler)


# ---------------------------------------------------------------------------
# Стоимость обслуживания в месяц: готовые суммы кнопками + свой вариант
# ---------------------------------------------------------------------------

def ask_cost_month(message) -> None:
    markup = types.InlineKeyboardMarkup()
    for option in COST_MONTH_OPTIONS:
        markup.add(
            types.InlineKeyboardButton(f"{option} ₽", callback_data=f"costmonth_{option}")
        )
    markup.add(types.InlineKeyboardButton("✏️ Свой вариант", callback_data="costmonth_custom"))
    bot.send_message(
        message.chat.id,
        "❔ Выберите стоимость обслуживания в месяц",
        reply_markup=markup,
    )


@bot.callback_query_handler(func=lambda callback: callback.data.startswith("costmonth_"))
@safe_handler
def choose_cost_month(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    option = callback.data.removeprefix("costmonth_")

    if option == "custom":
        bot.send_message(
            chat_id,
            "❔ Введите свою стоимость обслуживания в месяц",
            reply_markup=control_keyboard(),
        )
        replace_next_step(callback.message, receive_custom_cost_month)
        return

    user_data[user_id].setdefault("ser_fields", {})["cost_month"] = option
    bot.send_message(
        chat_id,
        f"✅ Стоимость обслуживания: <b>{option} ₽</b>",
        reply_markup=control_keyboard(),
    )
    ask_ser_field(callback.message, 0)


def receive_custom_cost_month(message) -> None:
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    value = parse_non_negative_integer(message, receive_custom_cost_month)
    if value is None:
        return
    user_data[message.from_user.id].setdefault("ser_fields", {})["cost_month"] = str(value)
    bot.send_message(
        message.chat.id,
        f"✅ Стоимость обслуживания: <b>{value} ₽</b>",
        reply_markup=control_keyboard(),
    )
    ask_ser_field(message, 0)


# ---------------------------------------------------------------------------
# Количество месяцев обслуживания: кнопки 3/2/1 + свой вариант
# ---------------------------------------------------------------------------

def finish_ser_fields(message) -> None:
    markup = types.InlineKeyboardMarkup()
    for option in MONTHS_COUNT_OPTIONS:
        markup.add(
            types.InlineKeyboardButton(option, callback_data=f"months_{option}")
        )
    markup.add(types.InlineKeyboardButton("✏️ Свой вариант", callback_data="months_custom"))
    bot.send_message(
        message.chat.id,
        "❔ Выберите количество месяцев обслуживания",
        reply_markup=markup,
    )


def proceed_after_months_count(message, chat_id, option) -> None:
    bot.send_message(
        chat_id,
        f"✅ Количество месяцев: <b>{option}</b>",
        reply_markup=control_keyboard(),
    )
    ask_advance(message)


@bot.callback_query_handler(func=lambda callback: callback.data.startswith("months_"))
@safe_handler
def choose_months_count(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    option = callback.data.removeprefix("months_")

    if option == "custom":
        bot.send_message(
            chat_id,
            "❔ Введите количество месяцев обслуживания",
            reply_markup=control_keyboard(),
        )
        replace_next_step(callback.message, receive_custom_months_count)
        return

    user_data[user_id].setdefault("ser_fields", {})["months_count"] = option
    proceed_after_months_count(callback.message, chat_id, option)


def receive_custom_months_count(message) -> None:
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    value = parse_non_negative_integer(message, receive_custom_months_count)
    if value is None:
        return
    option = str(value)
    user_data[message.from_user.id].setdefault("ser_fields", {})["months_count"] = option
    proceed_after_months_count(message, message.chat.id, option)


# ---------------------------------------------------------------------------
# Аванс: 100% кнопкой или свой вариант
# ---------------------------------------------------------------------------

def ask_advance(message) -> None:
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("100%", callback_data="advance_100%"))
    markup.add(types.InlineKeyboardButton("✏️ Свой вариант", callback_data="advance_custom"))
    bot.send_message(message.chat.id, "❔ Выберите размер аванса", reply_markup=markup)


def ask_advance_period(message) -> None:
    markup = types.InlineKeyboardMarkup()
    for advperiod_option in ADVANCE_PERIOD_OPTIONS:
        markup.add(
            types.InlineKeyboardButton(advperiod_option, callback_data=f"advperiod_{advperiod_option}")
        )
    bot.send_message(message.chat.id, "❔ Выберите аванс за период", reply_markup=markup)


@bot.callback_query_handler(func=lambda callback: callback.data.startswith("advance_"))
@safe_handler
def choose_advance(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    option = callback.data.removeprefix("advance_")

    if option == "custom":
        bot.send_message(chat_id, "❔ Введите размер аванса", reply_markup=control_keyboard())
        replace_next_step(callback.message, receive_custom_advance)
        return

    user_data[user_id].setdefault("ser_fields", {})["advance"] = option
    bot.send_message(
        chat_id,
        f"✅ Аванс: <b>{option}</b>",
        reply_markup=control_keyboard(),
    )
    ask_advance_period(callback.message)


def receive_custom_advance(message) -> None:
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    value = (message.text or "").strip()
    if not value:
        register_retry(message, receive_custom_advance, "❌ Значение не может быть пустым.")
        return
    user_data[message.from_user.id].setdefault("ser_fields", {})["advance"] = value
    bot.send_message(
        message.chat.id,
        f"✅ Аванс: <b>{value}</b>",
        reply_markup=control_keyboard(),
    )
    ask_advance_period(message)


@bot.callback_query_handler(func=lambda callback: callback.data.startswith("advperiod_"))
@safe_handler
def choose_advance_period(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    option = callback.data.removeprefix("advperiod_")
    user_data[user_id].setdefault("ser_fields", {})["advance_period"] = option

    bot.send_message(
        chat_id,
        f"✅ Аванс за период: <b>{option}</b>",
        reply_markup=control_keyboard(),
    )

    markup = types.InlineKeyboardMarkup()
    for freq_option in VISITS_FREQUENCY_OPTIONS:
        markup.add(
            types.InlineKeyboardButton(freq_option, callback_data=f"visits_{freq_option}")
        )
    bot.send_message(
        chat_id,
        "❔ Выберите периодичность обходов",
        reply_markup=markup,
    )


@bot.callback_query_handler(func=lambda callback: callback.data.startswith("visits_"))
@safe_handler
def choose_visits_frequency(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    option = callback.data.removeprefix("visits_")
    user_data[user_id].setdefault("ser_fields", {})["visits_frequency"] = option

    bot.send_message(
        chat_id,
        f"✅ Периодичность обходов: <b>{option}</b>",
        reply_markup=control_keyboard(),
    )

    markup = types.InlineKeyboardMarkup()
    for term_index, term_option in enumerate(TERMINATION_PERIOD_OPTIONS):
        markup.add(
            types.InlineKeyboardButton(term_option, callback_data=f"termperiod_{term_index}")
        )
    bot.send_message(
        chat_id,
        "❔ Выберите срок расторжения договора",
        reply_markup=markup,
    )


@bot.callback_query_handler(func=lambda callback: callback.data.startswith("termperiod_"))
@safe_handler
def choose_termination_period(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id

    if user_id not in user_data:
        show_start_button(callback.message, user_id)
        return

    option_index = int(callback.data.removeprefix("termperiod_"))
    option = TERMINATION_PERIOD_OPTIONS[option_index]
    user_data[user_id].setdefault("ser_fields", {})["termination_period"] = option

    bot.send_message(
        chat_id,
        f"✅ Срок расторжения: <b>{option}</b>",
        reply_markup=control_keyboard(),
    )

    markup = types.InlineKeyboardMarkup()
    markup.row(
        types.InlineKeyboardButton("🧾 Документ", callback_data="doc"),
        types.InlineKeyboardButton("🖼️ Картинка", callback_data="pic"),
    )
    markup.row(
        types.InlineKeyboardButton("📄 PDF-скан", callback_data="pdf"),
        types.InlineKeyboardButton("📢 Сообщение", callback_data="mes"),
    )
    bot.send_message(
        chat_id,
        "❔ Выберите источник реквизитов",
        reply_markup=markup,
    )


def parse_non_negative_integer(message, retry_handler: Callable) -> int | None:
    answer = (message.text or "").strip()
    if not answer.isdigit():
        register_retry(
            message,
            retry_handler,
            "❌ Введите целое неотрицательное число.",
        )
        return None
    return int(answer)


@safe_handler
def costs(message):
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    value = parse_non_negative_integer(message, costs)
    if value is None:
        return
    user_data[message.from_user.id]["cost"] = value
    bot.send_message(
        message.chat.id,
        f"✅ Стоимость принята: <b>{value:,} ₽</b>".replace(",", " "),
        reply_markup=control_keyboard(),
    )
    register_retry(message, complectation, "❔ Введите количество комплектов")


@safe_handler
def complectation(message):
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    amount = parse_non_negative_integer(message, complectation)
    if amount is None:
        return

    session = user_data[message.from_user.id]
    session["complects"] = amount
    session["count_print"] = 3900 * amount
    session["count_sending"] = session["count_print"] + 1000 if amount else 0
    bot.send_message(
        message.chat.id,
        f"✅ Количество комплектов принято: <b>{amount}</b>",
        reply_markup=control_keyboard(),
    )

    markup = types.InlineKeyboardMarkup()
    markup.row(
        types.InlineKeyboardButton("🧾 Документ", callback_data="doc"),
        types.InlineKeyboardButton("🖼️ Картинка", callback_data="pic"),
    )
    markup.row(
        types.InlineKeyboardButton("📄 PDF-скан", callback_data="pdf"),
        types.InlineKeyboardButton("📢 Сообщение", callback_data="mes"),
    )
    bot.send_message(
        message.chat.id,
        "❔ Выберите источник реквизитов",
        reply_markup=markup,
    )


@bot.callback_query_handler(func=lambda callback: True)
@safe_handler
def callback_message(callback):
    bot.answer_callback_query(callback.id)
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id
    with chat_lock(chat_id):
        if is_repeated_action(chat_id, f"source:{callback.data}"):
            return
        if user_id not in user_data:
            show_start_button(callback.message, user_id)
            return
        handlers = {
            "pic": (photo, "<b>▶️ Вышлите картинку</b>"),
            "doc": (sentdoc, "<b>▶️ Вышлите документ</b>"),
            "pdf": (pdf, "<b>▶️ Вышлите PDF-скан документа</b>"),
            "mes": (sentmes, "<b>▶️ Отправьте или перешлите сообщение</b>"),
        }
        selected = handlers.get(callback.data)
        if selected is None:
            LOGGER.warning("Неизвестный callback | data=%s", callback.data)
            return

        handler, prompt = selected
        source_names = {
            "pic": "изображение",
            "doc": "документ",
            "pdf": "PDF-скан",
            "mes": "текст сообщения",
        }
        bot.send_message(
            chat_id,
            f"✅ Выбран источник: <b>{source_names[callback.data]}</b>\n{prompt}",
            reply_markup=control_keyboard(),
        )
        bot.clear_step_handler_by_chat_id(chat_id)
        bot.register_next_step_handler(callback.message, handler)


def build_and_send_contract(
    message,
    company_data: list[str | None],
    source_path: Path | None = None,
    scenario_id: str | None = None,
) -> bool:
    user_id = message.from_user.id
    local_doc: Path | None = None
    try:
        if (
            scenario_id
            and user_data.get(user_id, {}).get("scenario_id") != scenario_id
        ):
            LOGGER.info("Отменена устаревшая обработка | user_id=%s", user_id)
            return False
        ms.validate_company_data(company_data)
        bot.send_message(
            message.chat.id,
            "🧾 Реквизиты извлечены. Рассчитываю суммы и заполняю шаблон…",
            reply_markup=control_keyboard(),
        )
        numer_contract = ms.get_bot_doc_num(user_data, user_id)

        if user_data[user_id].get("contract_type") == "ser":
            ser_fields = user_data[user_id].get("ser_fields", {})
            months_count = ser_fields.get("months_count", "")
            cost_month = ser_fields.get("cost_month", "")
            try:
                cost_total_value = int(months_count) * int(cost_month)
            except (ValueError, TypeError):
                cost_total_value = 0
            texted_total = ms.integer_texted(cost_total_value)
            local_doc = ms.bot_insert_req_ser(
                user_data,
                user_id,
                company_data,
                numer_contract,
                texted_total,
                source_path,
            )
        else:
            numer_count = ms.get_bot_count_num(user_data, user_id)
            texted_costs = ms.integer_texted(user_data[user_id]["cost"])
            texted_sending = ms.integer_texted(
                user_data[user_id]["count_sending"]
            )
            local_doc = ms.bot_insert_req(
                user_data,
                user_id,
                company_data,
                numer_contract,
                numer_count,
                texted_costs,
                texted_sending,
                source_path,
            )

        if (
            scenario_id
            and user_data.get(user_id, {}).get("scenario_id") != scenario_id
        ):
            LOGGER.info(
                "Отправка устаревшего договора отменена | user_id=%s",
                user_id,
            )
            return False

        try:
            bot.send_message(message.chat.id, "📤 Договор готов. Отправляю файл…")
        except Exception:
            LOGGER.warning(
                "Не удалось отправить статус перед загрузкой договора | user_id=%s",
                user_id,
            )
        send_contract_document(message.chat.id, local_doc)
        LOGGER.info("Договор отправлен | user_id=%s", user_id)
        return True
    finally:
        ms.delete_local_doc(local_doc)


def sentdoc(message):
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    source_path: Path | None = None
    try:
        scenario_id = user_data[message.from_user.id]["scenario_id"]
        source_path = ms.download_bot_doc(message, bot)
        if source_path is None:
            register_retry(
                message,
                sentdoc,
                "❌ Отправьте файл DOC, DOCX или PDF с текстовым слоем.",
            )
            return
        bot.send_message(
            message.chat.id,
            "✅ Документ получен. Извлекаю текст и передаю реквизиты ИИ — "
            "это может занять до минуты.",
            reply_markup=control_keyboard(),
        )
        company_data = ms.sent_doc_to_ai(source_path)
        if build_and_send_contract(
            message,
            company_data,
            source_path,
            scenario_id,
        ):
            show_start_button(message)
    except misc.MissingCompanyDetailsError as exc:
        report_missing_company_details(message, exc, sentdoc)
    except Exception as exc:
        report_processing_error(message, exc)
    finally:
        ms.delete_local_doc(source_path)


def pdf(message):
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    source_path: Path | None = None
    try:
        scenario_id = user_data[message.from_user.id]["scenario_id"]
        source_path = ms.download_bot_doc(message, bot)
        if source_path is None or source_path.suffix.lower() != ".pdf":
            register_retry(message, pdf, "❌ Отправьте PDF-файл.")
            return
        bot.send_message(
            message.chat.id,
            "✅ PDF-скан получен. Распознаю страницы и извлекаю реквизиты — "
            "пожалуйста, не отправляйте файл повторно.",
            reply_markup=control_keyboard(),
        )
        company_data = ms.sent_pdf_scan_to_ai(source_path)
        if build_and_send_contract(
            message,
            company_data,
            source_path,
            scenario_id,
        ):
            show_start_button(message)
    except misc.MissingCompanyDetailsError as exc:
        report_missing_company_details(message, exc, pdf)
    except Exception as exc:
        report_processing_error(message, exc)
    finally:
        ms.delete_local_doc(source_path)


def sentmes(message):
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    try:
        scenario_id = user_data[message.from_user.id]["scenario_id"]
        bot.send_message(
            message.chat.id,
            "✅ Сообщение принято. Извлекаю реквизиты…",
            reply_markup=control_keyboard(),
        )
        company_data = ms.sent_message_to_ai(message.text or "")
        if build_and_send_contract(
            message,
            company_data,
            scenario_id=scenario_id,
        ):
            show_start_button(message)
    except misc.MissingCompanyDetailsError as exc:
        report_missing_company_details(message, exc, sentmes)
    except Exception as exc:
        report_processing_error(message, exc)


def photo(message):
    if recover_if_requested(message) or not ensure_active_session(message):
        return
    source_path: Path | None = None
    try:
        scenario_id = user_data[message.from_user.id]["scenario_id"]
        if not message.photo:
            register_retry(message, photo, "❌ Отправьте изображение.")
            return
        image_id = message.photo[-1].file_id
        file_info = bot.get_file(image_id)
        content = bot.download_file(file_info.file_path)
        source_path = IMAGES_DIR / f"{uuid4().hex}.jpg"
        source_path.write_bytes(content)
        bot.send_message(
            message.chat.id,
            "✅ Изображение получено. Распознаю текст и извлекаю реквизиты — "
            "это может занять до минуты.",
            reply_markup=control_keyboard(),
        )
        company_data = ms.sent_image_to_ai(source_path)
        if build_and_send_contract(
            message,
            company_data,
            source_path,
            scenario_id,
        ):
            show_start_button(message)
    except misc.MissingCompanyDetailsError as exc:
        report_missing_company_details(message, exc, photo)
    except misc.DocumentExtractionError as exc:
        report_processing_error(message, exc)
        replace_next_step(message, photo)
    except Exception as exc:
        report_processing_error(message, exc)
    finally:
        ms.delete_local_doc(source_path)


@bot.message_handler(content_types=["text", "document", "photo"])
@employees_only
def recover_orphaned_message(message):
    """Обрабатывает сообщение, next-step которого потерялся после рестарта."""
    if is_control_message(message):
        landing(message)
        return
    LOGGER.info(
        "Получено сообщение без активного сценария | user_id=%s",
        message.from_user.id,
    )
    bot.send_message(
        message.chat.id,
        "⚠️ Активный шаг не найден. Возможно, бот был перезапущен. "
        "Нажмите «Начать».",
        reply_markup=control_keyboard(),
    )


def notify_staff_after_restart() -> None:
    """Показывает управляющую клавиатуру сотрудникам после запуска процесса."""
    for user_id in staff:
        try:
            bot.send_message(
                user_id,
                "✅ Бот доступен. Можно начать новый сценарий.",
                reply_markup=control_keyboard(),
            )
        except Exception:
            # Telegram запрещает писать пользователю, который ещё не запускал
            # бота или заблокировал его. Остальные сотрудники не должны страдать.
            LOGGER.warning(
                "Не удалось показать стартовую кнопку | user_id=%s",
                user_id,
            )


if __name__ == "__main__":
    LOGGER.info("Telegram-бот запускается")
    notify_staff_after_restart()

    while True:
        try:
            bot.infinity_polling(
                skip_pending=True,
                timeout=30,
                long_polling_timeout=30,
                allowed_updates=["message", "callback_query"],
            )
        except Exception:
            LOGGER.exception("Критическая ошибка polling — перезапускаю через 5 секунд")
            time.sleep(5)