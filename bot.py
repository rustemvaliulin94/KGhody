"""
Telegram-бот для управления музыкальной программой клуба.
Хранит ходы (идеи/события) с лайнапом артистов.
"""

import os
import json
import logging
import pathlib
import calendar
from datetime import datetime, timedelta
from collections import defaultdict
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, ConversationHandler, ContextTypes, filters
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")
_SHARED = pathlib.Path(os.getenv("SHARED_DIR", "/app/shared"))
_SHARED.mkdir(parents=True, exist_ok=True)
DATA_FILE = str(_SHARED / "data.json")

# ─── Главная клавиатура ───────────────────────────────────────────────────────

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [
        [KeyboardButton("➕ Новый ход"), KeyboardButton("📋 Список ходов")],
        [KeyboardButton("📅 Расписание"), KeyboardButton("🎤 Управление артистами")],
    ],
    resize_keyboard=True,
    is_persistent=True,
)

ARTISTS_KEYBOARD = ReplyKeyboardMarkup(
    [
        [KeyboardButton("➕ Добавить артиста"), KeyboardButton("✏️ Редактировать артиста")],
        [KeyboardButton("🗑 Удалить артиста"), KeyboardButton("◀️ Назад")],
    ],
    resize_keyboard=True,
    is_persistent=True,
)

# ─── Google Sheets ────────────────────────────────────────────────────────────

_sheets_client = None

def _get_sheets_client():
    global _sheets_client
    if _sheets_client is not None:
        return _sheets_client
    creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        return None
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        creds_dict = json.loads(creds_json)
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        _sheets_client = gspread.authorize(creds)
    except Exception as e:
        logging.getLogger(__name__).error(f"Google Sheets init error: {e}")
        return None
    return _sheets_client

_STATUS_LABEL = {"abstract": "Абстрактный", "concrete": "Конкретный", "confirmed": "Подтверждён"}

def _get_spreadsheet():
    """Возвращает объект таблицы или None."""
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    if not sheet_id:
        return None
    client = _get_sheets_client()
    if not client:
        return None
    try:
        return client.open_by_key(sheet_id)
    except Exception as e:
        logging.getLogger(__name__).error(f"Sheets open error: {e}")
        return None

def _ensure_worksheets(spreadsheet):
    """Создаёт листы с заголовками если их нет. Возвращает (ws_hods, ws_artists)."""
    try:
        ws_hods = spreadsheet.worksheet("Ходы")
    except Exception:
        ws_hods = spreadsheet.add_worksheet(title="Ходы", rows=1000, cols=10)
        ws_hods.update("A1", [["ID", "Дата", "Лайнап", "Статус", "Добавил", "Дата создания"]])

    try:
        ws_artists = spreadsheet.worksheet("Артисты")
    except Exception:
        ws_artists = spreadsheet.add_worksheet(title="Артисты", rows=1000, cols=10)
        ws_artists.update("A1", [["ID хода", "Дата хода", "Имя", "Время", "Контакт", "Соцсети", "Комментарий", "Добавил"]])

    return ws_hods, ws_artists

def _hod_row(h: dict) -> list:
    return [
        h.get("id", ""),
        h.get("date", ""),
        h.get("lineup", ""),
        _STATUS_LABEL.get(h.get("status", ""), h.get("status", "")),
        h.get("added_by", ""),
        h.get("created_at", "")[:10] if h.get("created_at") else "",
    ]

def sheets_upsert_hod(hod: dict):
    """Обновляет строку хода в листе «Ходы» или добавляет новую."""
    spreadsheet = _get_spreadsheet()
    if not spreadsheet:
        return
    try:
        ws_hods, _ = _ensure_worksheets(spreadsheet)
        col_ids = ws_hods.col_values(1)  # все значения колонки ID
        hod_id_str = str(hod.get("id", ""))
        if hod_id_str in col_ids:
            row_num = col_ids.index(hod_id_str) + 1
            ws_hods.update(f"A{row_num}", [_hod_row(hod)])
        else:
            ws_hods.append_row(_hod_row(hod))
    except Exception as e:
        logging.getLogger(__name__).error(f"sheets_upsert_hod error: {e}")

def sheets_delete_hod(hod_id: int):
    """Удаляет строку хода и все его артисты из таблицы."""
    spreadsheet = _get_spreadsheet()
    if not spreadsheet:
        return
    try:
        ws_hods, ws_artists = _ensure_worksheets(spreadsheet)
        hod_id_str = str(hod_id)

        col_ids = ws_hods.col_values(1)
        if hod_id_str in col_ids:
            ws_hods.delete_rows(col_ids.index(hod_id_str) + 1)

        # Удалить всех артистов этого хода (идём снизу чтобы не сбить индексы)
        artist_hod_ids = ws_artists.col_values(1)
        for i in range(len(artist_hod_ids) - 1, 0, -1):
            if artist_hod_ids[i] == hod_id_str:
                ws_artists.delete_rows(i + 1)
    except Exception as e:
        logging.getLogger(__name__).error(f"sheets_delete_hod error: {e}")

def sheets_sync_artists(hod: dict):
    """Перезаписывает артистов конкретного хода в листе «Артисты»."""
    spreadsheet = _get_spreadsheet()
    if not spreadsheet:
        return
    try:
        _, ws_artists = _ensure_worksheets(spreadsheet)
        hod_id_str = str(hod.get("id", ""))

        # Удалить старые строки этого хода снизу вверх
        artist_hod_ids = ws_artists.col_values(1)
        for i in range(len(artist_hod_ids) - 1, 0, -1):
            if artist_hod_ids[i] == hod_id_str:
                ws_artists.delete_rows(i + 1)

        # Добавить актуальные
        for a in hod.get("artists", []):
            ws_artists.append_row([
                hod.get("id", ""),
                hod.get("date", ""),
                a.get("name", ""),
                a.get("time", ""),
                a.get("contact", ""),
                a.get("social", ""),
                a.get("comment", ""),
                a.get("added_by", ""),
            ])
    except Exception as e:
        logging.getLogger(__name__).error(f"sheets_sync_artists error: {e}")

def sheets_full_sync(data: dict):
    """Полная перезапись — используется при первом запуске."""
    spreadsheet = _get_spreadsheet()
    if not spreadsheet:
        return
    try:
        ws_hods, ws_artists = _ensure_worksheets(spreadsheet)

        hods_rows = [["ID", "Дата", "Лайнап", "Статус", "Добавил", "Дата создания"]]
        artists_rows = [["ID хода", "Дата хода", "Имя", "Время", "Контакт", "Соцсети", "Комментарий", "Добавил"]]

        for h in data.get("hods", []):
            hods_rows.append(_hod_row(h))
            for a in h.get("artists", []):
                artists_rows.append([
                    h.get("id", ""), h.get("date", ""),
                    a.get("name", ""), a.get("time", ""),
                    a.get("contact", ""), a.get("social", ""),
                    a.get("comment", ""), a.get("added_by", ""),
                ])

        ws_hods.clear()
        ws_hods.update("A1", hods_rows)
        ws_artists.clear()
        ws_artists.update("A1", artists_rows)
    except Exception as e:
        logging.getLogger(__name__).error(f"sheets_full_sync error: {e}")

# Твой Telegram user_id. Узнай его у @userinfobot, затем задай:
#   export OWNER_ID=123456789
# Или замени 0 напрямую в коде.
try:
    OWNER_ID = int(os.getenv("OWNER_ID", "0").strip())
except ValueError:
    OWNER_ID = 0
    logger.error("OWNER_ID задан неверно — проверь переменную окружения")

def sheets_restore_to_json() -> bool:
    """Восстанавливает data.json из Google Sheets если файл пустой или отсутствует.
    Возвращает True если восстановление прошло успешно."""
    spreadsheet = _get_spreadsheet()
    if not spreadsheet:
        return False
    try:
        status_reverse = {"Абстрактный": "abstract", "Конкретный": "concrete", "Подтверждён": "confirmed"}

        # Читаем лист Ходы
        try:
            ws_hods = spreadsheet.worksheet("Ходы")
        except Exception:
            return False

        hods_rows = ws_hods.get_all_values()
        if len(hods_rows) <= 1:  # только заголовок или пусто
            return False

        # Читаем лист Артисты
        try:
            ws_artists = spreadsheet.worksheet("Артисты")
            artists_rows = ws_artists.get_all_values()[1:]  # без заголовка
        except Exception:
            artists_rows = []

        # Группируем артистов по id хода
        artists_by_hod = {}
        for row in artists_rows:
            if not row or not row[0]:
                continue
            hod_id = int(row[0])
            if hod_id not in artists_by_hod:
                artists_by_hod[hod_id] = []
            artists_by_hod[hod_id].append({
                "name":    row[2] if len(row) > 2 else "",
                "time":    row[3] if len(row) > 3 else "",
                "contact": row[4] if len(row) > 4 else "",
                "social":  row[5] if len(row) > 5 else "",
                "comment": row[6] if len(row) > 6 else "",
                "added_by": row[7] if len(row) > 7 else "",
                "photo":   "",
            })

        # Собираем ходы
        hods = []
        for row in hods_rows[1:]:  # без заголовка
            if not row or not row[0]:
                continue
            try:
                hod_id = int(row[0])
            except ValueError:
                continue
            hods.append({
                "id":         hod_id,
                "date":       row[1] if len(row) > 1 else "",
                "lineup":     row[2] if len(row) > 2 else "",
                "status":     status_reverse.get(row[3], "concrete") if len(row) > 3 else "concrete",
                "added_by":   row[4] if len(row) > 4 else "",
                "created_at": row[5] if len(row) > 5 else "",
                "artists":    artists_by_hod.get(hod_id, []),
            })

        if not hods:
            return False

        data = {"hods": hods}
        save_data(data)
        logger.info(f"Восстановлено из Google Sheets: {len(hods)} ходов")
        return True

    except Exception as e:
        logger.error(f"sheets_restore_to_json error: {e}")
        return False

# ─── Контроль доступа ──────────────────────────────────────────────────────────────────

def get_allowed_users(data: dict) -> set:
    allowed = set(data.get("allowed_users", []))
    if OWNER_ID:
        allowed.add(OWNER_ID)
    return allowed

def is_allowed(user_id: int, data: dict) -> bool:
    return user_id in get_allowed_users(data)

def is_owner(user_id: int) -> bool:
    return OWNER_ID != 0 and user_id == OWNER_ID

def require_access(func):
    """Декоратор: отклоняет команду если пользователь не в белом списке."""
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        user = update.effective_user
        data = load_data()
        if not is_allowed(user.id, data):
            await update.message.reply_text(
                "⛔ У тебя нет доступа к этому боту.\n"
                "Напиши администратору — он добавит тебя командой /adduser."
            )
            logger.warning(f"Отказ в доступе: {user.id} (@{user.username})")
            return
        return await func(update, context)
    wrapper.__name__ = func.__name__
    return wrapper

# ─── Состояния диалогов ───────────────────────────────────────────────────────

# Добавление хода
ADD_DATE, ADD_LINEUP = range(2)

# Редактирование хода
EDIT_CHOOSE, EDIT_DATE, EDIT_LINEUP = range(10, 13)

# Добавление артиста
ARTIST_SELECT_HOD, ARTIST_NAME, ARTIST_TIME, ARTIST_CONTACT, ARTIST_SOCIAL, ARTIST_PHOTO, ARTIST_COMMENT = range(20, 27)

# Редактирование артиста
EDIT_ARTIST_SELECT_HOD, EDIT_ARTIST_SELECT, EDIT_ARTIST_FIELD, EDIT_ARTIST_VALUE, EDIT_ARTIST_PHOTO, EDIT_ARTIST_HOD_PICK = range(30, 36)

# Расписание
SCHED_PERIOD, SCHED_FROM, SCHED_TO = range(40, 43)

# ─── Работа с данными ─────────────────────────────────────────────────────────

def load_data() -> dict:
    if not os.path.exists(DATA_FILE):
        return {"hods": []}
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data: dict):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_next_id(data: dict) -> int:
    if not data["hods"]:
        return 1
    return max(h["id"] for h in data["hods"]) + 1

def find_hod(data: dict, hod_id: int) -> dict | None:
    for h in data["hods"]:
        if h["id"] == hod_id:
            return h
    return None

def format_hod(hod: dict, short: bool = False) -> str:
    status_emoji = {"abstract": "🌀", "concrete": "📅", "confirmed": "✅"}
    status_label = {"abstract": "Абстрактный", "concrete": "Конкретный", "confirmed": "Подтверждён"}

    emoji = status_emoji.get(hod["status"], "❓")
    label = status_label.get(hod["status"], hod["status"])

    date_str = hod.get("date") or "без даты"
    lines = [
        f"{emoji} *Ход #{hod['id']}* — {date_str}",
        f"Статус: {label}",
    ]

    if hod.get("lineup"):
        lines.append(f"Лайнап: {hod['lineup']}")

    if not short and hod.get("artists"):
        lines.append("\n🎤 *Артисты:*")
        for a in hod["artists"]:
            lines.append(f"  • *{a['name']}* {a.get('time', '')}".strip())
            if a.get("contact"):
                lines.append(f"    📞 {a['contact']}")
            if a.get("social"):
                lines.append(f"    🔗 {a['social']}")
            if a.get("comment"):
                lines.append(f"    💬 {a['comment']}")

    lines.append(f"\n_Добавил: {hod['added_by']}_")
    return "\n".join(lines)

# ─── /start ───────────────────────────────────────────────────────────────────

@require_access
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = "🎵 *Бот музыкальной программы клуба*\n\nВыбери действие на клавиатуре ниже."
    if is_owner(user.id):
        text += (
            "\n\n👑 *Команды администратора:*\n"
            "  /adduser `<id>` `[имя]` — добавить участника\n"
            "  /removeuser `<id>` — убрать участника\n"
            "  /listusers — список участников\n"
        )
    text += "\n  /myid — узнать свой Telegram ID"
    await update.message.reply_text(text, parse_mode="Markdown", reply_markup=MAIN_KEYBOARD)

# ─── Добавление хода (/newhod) ────────────────────────────────────────────────

@require_access
async def newhod_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["_in_conversation"] = True
    await update.message.reply_text(
        "📅 Введи дату хода в формате ДД.ММ.ГГГГ\n"
        "Например: 28.06.2025\n\n"
        "Или напиши *пропустить* — ход попадёт в список абстрактных 🌀\n\n"
        "/cancel — отмена",
        parse_mode="Markdown"
    )
    return ADD_DATE

async def newhod_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() in ("пропустить", "skip", "-"):
        context.user_data["new_date"] = ""
    else:
        try:
            datetime.strptime(text, "%d.%m.%Y")
        except ValueError:
            await update.message.reply_text("❌ Неверный формат. Введи дату (ДД.ММ.ГГГГ) или напиши *пропустить*:", parse_mode="Markdown")
            return ADD_DATE
        context.user_data["new_date"] = text

    await update.message.reply_text(
        "✏️ Введи лайнап (краткое описание программы).\n"
        "Можно написать 'пропустить' и заполнить позже."
    )
    return ADD_LINEUP

async def newhod_lineup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lineup = update.message.text.strip()
    if lineup.lower() in ("пропустить", "skip", "-"):
        lineup = ""

    data = load_data()
    date = context.user_data.get("new_date", "")
    if not date:
        status = "abstract"
    else:
        status = "concrete"
    hod = {
        "id": get_next_id(data),
        "date": date,
        "lineup": lineup,
        "status": status,
        "artists": [],
        "added_by": update.effective_user.full_name,
        "created_at": datetime.now().isoformat(),
    }
    data["hods"].append(hod)
    save_data(data)
    sheets_upsert_hod(hod)

    await update.message.reply_text(
        f"✅ Ход #{hod['id']} добавлен!\n\n{format_hod(hod)}",
        parse_mode="Markdown"
    )
    return ConversationHandler.END

# ─── Список ходов (/list) ─────────────────────────────────────────────────────

@require_access
async def list_hods(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["hods"]:
        await update.message.reply_text("Пока нет ни одного хода. Добавь первый: /newhod")
        return

    keyboard = [
        [
            InlineKeyboardButton("Все актуальные", callback_data="filter_all"),
            InlineKeyboardButton("🌀 Абстрактные", callback_data="filter_abstract"),
        ],
        [
            InlineKeyboardButton("📅 Конкретные", callback_data="filter_concrete"),
            InlineKeyboardButton("✅ Подтверждённые", callback_data="filter_confirmed"),
        ],
    ]
    await update.message.reply_text(
        "Показать ходы:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def filter_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    data = load_data()
    filter_key = query.data.replace("filter_", "")
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    def is_relevant(h):
        """Абстрактные всегда актуальны. Ходы с датой — только сегодня или в будущем."""
        if not h.get("date"):
            return True
        try:
            return datetime.strptime(h["date"], "%d.%m.%Y") >= today
        except ValueError:
            return True

    # Сначала фильтруем прошедшие
    hods = [h for h in data["hods"] if is_relevant(h)]

    # Затем по статусу если нужно
    if filter_key != "all":
        hods = [h for h in hods if h["status"] == filter_key]

    if not hods:
        await query.edit_message_text("Актуальных ходов с таким статусом нет.")
        return

    # Сортируем: конкретные по дате, абстрактные в конец
    def sort_key(h):
        if h.get("date"):
            try:
                return (0, datetime.strptime(h["date"], "%d.%m.%Y"))
            except Exception:
                pass
        return (1, datetime.min)
    hods_sorted = sorted(hods, key=sort_key)

    # Компактный список в одно сообщение
    status_emoji = {"abstract": "🌀", "concrete": "📅", "confirmed": "✅"}
    lines = [f"Найдено: {len(hods_sorted)}\n"]
    for hod in hods_sorted:
        emoji = status_emoji.get(hod.get("status", ""), "❓")
        date_str = hod.get("date") or "без даты"
        lineup = hod.get("lineup", "")
        hod_id = hod.get("id", "")
        artist_count = len(hod.get("artists", []))
        artist_str = f" · {artist_count} арт." if artist_count else ""
        lineup_str = f" | {lineup}" if lineup else ""
        lines.append(f"{emoji} *#{hod_id}* — {date_str}{lineup_str}{artist_str} → /showhod {hod_id}")

    text = "\n".join(lines)
    # Если текст длиннее лимита — режем на части
    if len(text) <= 4096:
        await query.edit_message_text(text, parse_mode="Markdown")
    else:
        await query.edit_message_text(lines[0], parse_mode="Markdown")
        chunk = []
        chunk_len = 0
        for line in lines[1:]:
            if chunk_len + len(line) + 1 > 4000:
                await query.message.reply_text("\n".join(chunk), parse_mode="Markdown")
                chunk = []
                chunk_len = 0
            chunk.append(line)
            chunk_len += len(line) + 1
        if chunk:
            await query.message.reply_text("\n".join(chunk), parse_mode="Markdown")

# ─── Расписание (/schedule) ───────────────────────────────────────────────────

@require_access
async def schedule_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [
            InlineKeyboardButton("Эта неделя", callback_data="sched_week"),
            InlineKeyboardButton("Этот месяц", callback_data="sched_month"),
        ],
        [
            InlineKeyboardButton("Следующий месяц", callback_data="sched_nextmonth"),
            InlineKeyboardButton("Свой период", callback_data="sched_custom"),
        ],
        [
            InlineKeyboardButton("🗓 Афиша на эту неделю", callback_data="sched_poster"),
        ],
    ]
    await update.message.reply_text(
        "📅 За какой период показать расписание?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return SCHED_PERIOD

async def schedule_period_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    now = datetime.now()

    if query.data == "sched_week":
        date_from = now.replace(hour=0, minute=0, second=0, microsecond=0)
        date_to = date_from + timedelta(days=6)
        await query.edit_message_text(
            f"📅 Расписание на неделю: {date_from.strftime('%d.%m')} — {date_to.strftime('%d.%m.%Y')}"
        )
        await _send_schedule(query.message, date_from, date_to)
        return ConversationHandler.END

    elif query.data == "sched_month":
        date_from = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        last_day = calendar.monthrange(now.year, now.month)[1]
        date_to = now.replace(day=last_day, hour=23, minute=59, second=59)
        await query.edit_message_text(
            f"📅 Расписание на {now.strftime('%B %Y')}"
        )
        await _send_schedule(query.message, date_from, date_to)
        return ConversationHandler.END

    elif query.data == "sched_nextmonth":
        if now.month == 12:
            year, month = now.year + 1, 1
        else:
            year, month = now.year, now.month + 1
        date_from = now.replace(year=year, month=month, day=1, hour=0, minute=0, second=0, microsecond=0)
        last_day = calendar.monthrange(year, month)[1]
        date_to = date_from.replace(day=last_day, hour=23, minute=59, second=59)
        await query.edit_message_text(
            f"📅 Расписание на следующий месяц ({date_from.strftime('%B %Y')})"
        )
        await _send_schedule(query.message, date_from, date_to)
        return ConversationHandler.END

    elif query.data == "sched_poster":
        date_from = now.replace(hour=0, minute=0, second=0, microsecond=0)
        date_to = date_from + timedelta(days=6)
        await query.edit_message_text("🗓 Генерирую афишу на неделю...")

        data = load_data()
        confirmed = [h for h in data["hods"] if h["status"] == "confirmed"]
        filtered = []
        for hod in confirmed:
            try:
                hod_date = datetime.strptime(hod["date"], "%d.%m.%Y")
                if date_from <= hod_date <= date_to:
                    filtered.append((hod_date, hod))
            except ValueError:
                continue

        if not filtered:
            await query.message.reply_text("На эту неделю подтверждённых выступлений нет.")
            return ConversationHandler.END

        poster = _format_poster(filtered)
        await query.message.reply_text(poster, parse_mode="Markdown", disable_web_page_preview=True)
        return ConversationHandler.END

    elif query.data == "sched_custom":
        await query.edit_message_text(
            "Введи дату начала периода (ДД.ММ.ГГГГ):\n\n/cancel — отмена"
        )
        return SCHED_FROM

async def schedule_from(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        date_from = datetime.strptime(text, "%d.%m.%Y")
    except ValueError:
        await update.message.reply_text("❌ Неверный формат. Введи дату начала (ДД.ММ.ГГГГ):")
        return SCHED_FROM

    context.user_data["sched_from"] = date_from
    await update.message.reply_text("Введи дату конца периода (ДД.ММ.ГГГГ):")
    return SCHED_TO

async def schedule_to(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        date_to = datetime.strptime(text, "%d.%m.%Y")
    except ValueError:
        await update.message.reply_text("❌ Неверный формат. Введи дату конца (ДД.ММ.ГГГГ):")
        return SCHED_TO

    date_from = context.user_data["sched_from"]
    if date_to < date_from:
        await update.message.reply_text("❌ Дата конца не может быть раньше даты начала. Введи ещё раз:")
        return SCHED_TO

    await update.message.reply_text(
        f"📅 Расписание: {date_from.strftime('%d.%m.%Y')} — {date_to.strftime('%d.%m.%Y')}"
    )
    await _send_schedule(update.message, date_from, date_to)
    return ConversationHandler.END

def _format_poster(filtered: list) -> str:
    """Форматирует афишу по дням в нужном стиле."""
    days = defaultdict(list)
    for hod_date, hod in filtered:
        days[hod_date].append(hod)

    lines = []
    for day_date in sorted(days.keys()):
        # Wednesday 24 June
        day_str = day_date.strftime("%A %-d %B")
        lines.append(f"*{day_str}*")
        for hod in sorted(days[day_date], key=lambda h: h.get("artists", [{}])[0].get("time", "") if h.get("artists") else ""):
            for artist in sorted(hod.get("artists", []), key=lambda a: a.get("time", "")):
                time_ = artist.get("time", "")
                name = artist.get("name", "")
                social = artist.get("social", "").strip()
                # Формируем строку с именем и ссылкой
                if social:
                    artist_str = f"[{name}]({social})"
                else:
                    artist_str = name
                lines.append(f"{time_} {artist_str}")
        lines.append("")  # пустая строка между днями

    return "\n".join(lines).strip()

async def _send_schedule(message, date_from: datetime, date_to: datetime):
    """Вспомогательная функция: отфильтровать и отправить расписание."""
    data = load_data()
    confirmed = [h for h in data["hods"] if h["status"] == "confirmed"]

    filtered = []
    for hod in confirmed:
        try:
            hod_date = datetime.strptime(hod["date"], "%d.%m.%Y")
            if date_from <= hod_date <= date_to:
                filtered.append((hod_date, hod))
        except ValueError:
            continue

    if not filtered:
        await message.reply_text("За этот период подтверждённых выступлений нет.")
        return

    filtered.sort(key=lambda x: x[0])

    # Собираем компактный список в одно сообщение
    lines = [f"Найдено выступлений: {len(filtered)}\n"]
    for hod_date, hod in filtered:
        date_str = hod_date.strftime("%d.%m.%Y")
        weekday = hod_date.strftime("%A")
        lineup = hod.get("lineup", "")
        hod_id = hod.get("id", "")
        artists = hod.get("artists", [])

        lines.append(f"*{weekday}, {date_str}* — {lineup}")
        for a in sorted(artists, key=lambda x: x.get("time", "")):
            time_ = a.get("time", "")
            name = a.get("name", "")
            lines.append(f"  {time_} {name}".strip())
        lines.append(f"  👉 /showhod {hod_id}")
        lines.append("")

    # Telegram лимит 4096 символов — если длиннее, режем на части
    text = "\n".join(lines)
    if len(text) <= 4096:
        await message.reply_text(text, parse_mode="Markdown")
    else:
        chunk = []
        chunk_len = 0
        for line in lines:
            if chunk_len + len(line) + 1 > 4000:
                await message.reply_text("\n".join(chunk), parse_mode="Markdown")
                chunk = []
                chunk_len = 0
            chunk.append(line)
            chunk_len += len(line) + 1
        if chunk:
            await message.reply_text("\n".join(chunk), parse_mode="Markdown")

@require_access
async def showhod(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or not args[0].isdigit():
        data = load_data()
        await update.message.reply_text(
            "Выбери ход:",
            reply_markup=_hod_picker_keyboard(data, "pick_show_")
        )
        return

    hod_id = int(args[0])
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await update.message.reply_text(f"❌ Ход #{hod_id} не найден.")
        return

    keyboard = []
    if hod.get("status") == "confirmed" and hod.get("artists"):
        keyboard.append([InlineKeyboardButton("📤 Отправить пак", callback_data=f"sendpack_{hod_id}")])

    await update.message.reply_text(
        format_hod(hod),
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard) if keyboard else None
    )

# ─── Подтверждение хода (/confirm) ───────────────────────────────────────────

@require_access
async def confirm_hod(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or not args[0].isdigit():
        data = load_data()
        # Только конкретные (с датой, не подтверждённые)
        confirmable = {h["id"]: h for h in data["hods"]
                       if h.get("status") == "concrete" and h.get("date")}
        if not confirmable:
            await update.message.reply_text("Нет ходов готовых к подтверждению.")
            return
        await update.message.reply_text(
            "Выбери ход для подтверждения:",
            reply_markup=_hod_picker_keyboard(
                {"hods": list(confirmable.values())}, "pick_confirm_"
            )
        )
        return

    hod_id = int(args[0])
    data = load_data()
    hod = find_hod(data, hod_id)

    if not hod:
        await update.message.reply_text(f"❌ Ход #{hod_id} не найден.")
        return

    if not hod.get("date"):
        await update.message.reply_text(
            f"❌ Нельзя подтвердить абстрактный ход без даты.\n"
            f"Сначала добавь дату: /edit {hod_id}"
        )
        return

    hod["status"] = "confirmed"
    hod["confirmed_by"] = update.effective_user.full_name
    hod["confirmed_at"] = datetime.now().isoformat()
    save_data(data)
    sheets_upsert_hod(hod)

    await update.message.reply_text(
        f"✅ Ход #{hod_id} подтверждён!\n\n{format_hod(hod)}",
        parse_mode="Markdown"
    )

async def confirm_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    hod_id = int(query.data.replace("confirm_", ""))
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await query.edit_message_text("Ход не найден.")
        return

    if not hod.get("date"):
        await query.answer("Сначала добавь дату командой /edit", show_alert=True)
        return

    hod["status"] = "confirmed"
    hod["confirmed_by"] = query.from_user.full_name
    hod["confirmed_at"] = datetime.now().isoformat()
    save_data(data)
    sheets_upsert_hod(hod)

    await query.edit_message_text(
        f"✅ Подтверждено!\n\n{format_hod(hod, short=True)}",
        parse_mode="Markdown"
    )

# ─── Редактирование хода (/edit) ─────────────────────────────────────────────

@require_access
async def edit_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or not args[0].isdigit():
        data = load_data()
        await update.message.reply_text(
            "Выбери ход для редактирования:",
            reply_markup=_hod_picker_keyboard(data, "pick_edit_")
        )
        return ConversationHandler.END

    hod_id = int(args[0])
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await update.message.reply_text(f"❌ Ход #{hod_id} не найден.")
        return ConversationHandler.END

    context.user_data["edit_hod_id"] = hod_id
    # Показываем "Убрать дату" только если дата есть и ход не подтверждён
    buttons = [
        [InlineKeyboardButton("📅 Дата", callback_data="edit_field_date"),
         InlineKeyboardButton("✏️ Лайнап", callback_data="edit_field_lineup")],
    ]
    if hod.get("date") and hod.get("status") != "confirmed":
        buttons.append([InlineKeyboardButton("🌀 Убрать дату → сделать абстрактным", callback_data="edit_field_removedate")])
    buttons.append([InlineKeyboardButton("❌ Отмена", callback_data="edit_cancel")])

    await update.message.reply_text(
        f"Редактируем ход #{hod_id}. Что изменить?",
        reply_markup=InlineKeyboardMarkup(buttons)
    )
    return EDIT_CHOOSE

async def edit_field_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "edit_cancel":
        await query.edit_message_text("Отмена.")
        return ConversationHandler.END

    field = query.data.replace("edit_field_", "")
    context.user_data["edit_field"] = field

    if field == "date":
        await query.edit_message_text("Введи новую дату (ДД.ММ.ГГГГ):")
        return EDIT_DATE
    elif field == "lineup":
        await query.edit_message_text("Введи новый лайнап:")
        return EDIT_LINEUP
    elif field == "removedate":
        hod_id = context.user_data["edit_hod_id"]
        data = load_data()
        hod = find_hod(data, hod_id)
        hod["date"] = ""
        hod["status"] = "abstract"
        save_data(data)
        sheets_upsert_hod(hod)
        await query.edit_message_text(
            f"🌀 Дата убрана. Ход #{hod_id} теперь абстрактный."
        )
        return ConversationHandler.END

async def edit_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        datetime.strptime(text, "%d.%m.%Y")
    except ValueError:
        await update.message.reply_text("❌ Неверный формат. Введи дату (ДД.ММ.ГГГГ):")
        return EDIT_DATE

    data = load_data()
    hod = find_hod(data, context.user_data["edit_hod_id"])
    was_abstract = hod.get("status") == "abstract"
    hod["date"] = text
    if was_abstract:
        hod["status"] = "concrete"
        await update.message.reply_text(
            f"✅ Дата добавлена: {text}\nХод переведён из абстрактного в конкретный!\n\n{format_hod(hod)}",
            parse_mode="Markdown"
        )
    else:
        await update.message.reply_text(f"✅ Дата обновлена: {text}")
    save_data(data)
    sheets_upsert_hod(hod)
    return ConversationHandler.END

async def edit_lineup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    data = load_data()
    hod = find_hod(data, context.user_data["edit_hod_id"])
    hod["lineup"] = text
    save_data(data)
    sheets_upsert_hod(hod)
    await update.message.reply_text(f"✅ Лайнап обновлён.")
    return ConversationHandler.END

# ─── Добавление артиста (/addartist) ─────────────────────────────────────────

@require_access
async def addartist_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or not args[0].isdigit():
        data = load_data()
        await update.message.reply_text(
            "Выбери ход, к которому добавить артиста:",
            reply_markup=_hod_picker_keyboard(data, "pick_addartist_")
        )
        return ConversationHandler.END

    hod_id = int(args[0])
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await update.message.reply_text(f"❌ Ход #{hod_id} не найден.")
        return ConversationHandler.END

    context.user_data["artist_hod_id"] = hod_id
    context.user_data["new_artist"] = {}
    await update.message.reply_text(
        f"Добавляем артиста к ходу #{hod_id} ({hod['date']})\n\n"
        "Введи имя артиста / название:"
    )
    return ARTIST_NAME

async def artist_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_artist"]["name"] = update.message.text.strip()
    await update.message.reply_text("🕐 Время выступления (например, 21:00). Или 'пропустить':")
    return ARTIST_TIME

async def artist_time(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    context.user_data["new_artist"]["time"] = "" if val.lower() in ("пропустить", "skip", "-") else val
    await update.message.reply_text("📞 Контакт (телефон или Telegram). Или 'пропустить':")
    return ARTIST_CONTACT

async def artist_contact(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    context.user_data["new_artist"]["contact"] = "" if val.lower() in ("пропустить", "skip", "-") else val
    await update.message.reply_text("🔗 Ссылка на соцсети. Или 'пропустить':")
    return ARTIST_SOCIAL

async def artist_social(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = (update.message.text or "").strip()
    if not val:
        await update.message.reply_text(
            "Пожалуйста, отправь ссылку текстом. Или напиши пропустить."
        )
        return ARTIST_SOCIAL
    context.user_data["new_artist"]["social"] = "" if val.lower() in ("пропустить", "skip", "-") else val
    await update.message.reply_text(
        "🖼 Отправь фото артиста (или пропустить):",
    )
    return ARTIST_PHOTO

async def artist_photo_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Если пользователь написал 'пропустить' вместо фото"""
    context.user_data["new_artist"]["photo"] = ""
    await update.message.reply_text("💬 Комментарий (особые условия, райдер и т.д.). Или 'пропустить':")
    return ARTIST_COMMENT

async def artist_photo_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Если пользователь отправил фото"""
    photo = update.message.photo[-1]
    context.user_data["new_artist"]["photo"] = photo.file_id
    await update.message.reply_text("💬 Комментарий (особые условия, райдер и т.д.). Или 'пропустить':")
    return ARTIST_COMMENT

async def artist_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    context.user_data["new_artist"]["comment"] = "" if val.lower() in ("пропустить", "skip", "-") else val

    data = load_data()
    hod = find_hod(data, context.user_data["artist_hod_id"])
    artist = context.user_data["new_artist"]
    artist["added_by"] = update.effective_user.full_name
    hod["artists"].append(artist)
    save_data(data)
    sheets_sync_artists(hod)

    name = artist["name"]
    time_ = artist.get("time", "")
    await update.message.reply_text(
        f"✅ Артист *{name}* {time_} добавлен к ходу #{hod['id']}!".strip(),
        parse_mode="Markdown"
    )
    return ConversationHandler.END


def _hod_picker_keyboard(data: dict, callback_prefix: str, only_with_artists: bool = False) -> InlineKeyboardMarkup:
    """Строит клавиатуру выбора хода из актуальных ходов."""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    hods = []
    for h in data["hods"]:
        if h.get("date"):
            try:
                if datetime.strptime(h["date"], "%d.%m.%Y") < today:
                    continue
            except ValueError:
                pass
        if only_with_artists and not h.get("artists"):
            continue
        hods.append(h)

    def sort_key(h):
        if h.get("date"):
            try:
                return (0, datetime.strptime(h["date"], "%d.%m.%Y"))
            except Exception:
                pass
        return (1, datetime.min)
    hods = sorted(hods, key=sort_key)

    status_emoji = {"abstract": "🌀", "concrete": "📅", "confirmed": "✅"}
    keyboard = []
    for h in hods:
        emoji = status_emoji.get(h.get("status", ""), "❓")
        date_str = h.get("date") or "без даты"
        lineup = h.get("lineup", "")
        label = f"{emoji} #{h['id']} — {date_str}"
        if lineup:
            label += f" | {lineup[:20]}"
        keyboard.append([InlineKeyboardButton(label, callback_data=f"{callback_prefix}{h['id']}")])

    keyboard.append([InlineKeyboardButton("❌ Отмена", callback_data=f"{callback_prefix}cancel")])
    return InlineKeyboardMarkup(keyboard)

# ─── Просмотр артистов (callback) ────────────────────────────────────────────

async def artists_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    hod_id = int(query.data.replace("artists_", ""))
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await query.edit_message_text("Ход не найден.")
        return

    keyboard = None
    if hod.get("status") == "confirmed" and hod.get("artists"):
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("📤 Отправить пак", callback_data=f"sendpack_{hod['id']}")
        ]])

    await query.edit_message_text(
        format_hod(hod),
        parse_mode="Markdown",
        reply_markup=keyboard
    )

# ─── Пак сообщений (/sendpack) ───────────────────────────────────────────────

async def sendpack_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    hod_id = int(query.data.replace("sendpack_", ""))
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod or not hod.get("artists"):
        await query.message.reply_text("Нет артистов для отправки.")
        return

    # День недели на английском по дате хода
    weekday_en = ""
    if hod.get("date"):
        try:
            hod_date = datetime.strptime(hod["date"], "%d.%m.%Y")
            weekday_en = hod_date.strftime("%A")  # Monday, Tuesday...
        except ValueError:
            pass

    await query.message.reply_text(f"📤 Отправляю пак для хода #{hod_id}...")

    for artist in hod["artists"]:
        name = artist.get("name", "")
        time_ = artist.get("time", "")
        photo = artist.get("photo", "")

        # Подпись: день недели, дата, имя, время
        parts = []
        if weekday_en:
            parts.append(weekday_en)
        if hod.get("date"):
            parts.append(hod["date"])
        parts.append(name)
        if time_:
            parts.append(time_)
        caption = " | ".join(parts)

        if photo:
            try:
                await query.message.reply_photo(photo=photo, caption=caption)
                continue
            except Exception:
                pass
        # Если фото нет или не удалось отправить — только текст
        await query.message.reply_text(caption)

# ─── Редактирование артиста (/editartist) ────────────────────────────────────

@require_access
async def editartist_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or not args[0].isdigit():
        data = load_data()
        await update.message.reply_text(
            "Выбери ход для редактирования артиста:",
            reply_markup=_hod_picker_keyboard(data, "ea_hod_", only_with_artists=True)
        )
        return EDIT_ARTIST_HOD_PICK

    hod_id = int(args[0])
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod or not hod.get("artists"):
        await update.message.reply_text(f"Ход #{hod_id} не найден или у него нет артистов.")
        return ConversationHandler.END

    context.user_data["edit_artist_hod_id"] = hod_id
    keyboard = [
        [InlineKeyboardButton(f"{i+1}. {a['name']}", callback_data=f"ea_select_{i}")]
        for i, a in enumerate(hod["artists"])
    ]
    keyboard.append([InlineKeyboardButton("❌ Отмена", callback_data="ea_cancel")])
    await update.message.reply_text(
        "Выбери артиста для редактирования:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return EDIT_ARTIST_SELECT


async def editartist_hod_pick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Выбор хода внутри диалога редактирования артиста."""
    query = update.callback_query
    await query.answer()
    val = query.data.replace("ea_hod_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return ConversationHandler.END
    hod_id = int(val)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod or not hod.get("artists"):
        await query.edit_message_text("У этого хода нет артистов.")
        return ConversationHandler.END
    context.user_data["edit_artist_hod_id"] = hod_id
    keyboard = [
        [InlineKeyboardButton(f"{i+1}. {a['name']}", callback_data=f"ea_select_{i}")]
        for i, a in enumerate(hod["artists"])
    ]
    keyboard.append([InlineKeyboardButton("❌ Отмена", callback_data="ea_cancel")])
    await query.edit_message_text(
        "Выбери артиста для редактирования:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return EDIT_ARTIST_SELECT

async def editartist_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "ea_cancel":
        await query.edit_message_text("Отмена.")
        return ConversationHandler.END

    idx = int(query.data.replace("ea_select_", ""))
    context.user_data["edit_artist_idx"] = idx

    keyboard = [
        [InlineKeyboardButton("Имя", callback_data="ea_field_name"),
         InlineKeyboardButton("Время", callback_data="ea_field_time")],
        [InlineKeyboardButton("Контакт", callback_data="ea_field_contact"),
         InlineKeyboardButton("Соцсети", callback_data="ea_field_social")],
        [InlineKeyboardButton("Комментарий", callback_data="ea_field_comment"),
         InlineKeyboardButton("🖼 Фото", callback_data="ea_field_photo")],
        [InlineKeyboardButton("❌ Отмена", callback_data="ea_cancel")]
    ]
    await query.edit_message_text(
        "Что изменить?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return EDIT_ARTIST_FIELD

async def editartist_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "ea_cancel":
        await query.edit_message_text("Отмена.")
        return ConversationHandler.END

    field = query.data.replace("ea_field_", "")
    context.user_data["edit_artist_field"] = field

    if field == "photo":
        await query.edit_message_text("🖼 Отправь новое фото артиста:")
        return EDIT_ARTIST_PHOTO

    labels = {"name": "имя", "time": "время", "contact": "контакт", "social": "соцсети", "comment": "комментарий"}
    await query.edit_message_text(f"Введи новое значение для поля «{labels.get(field, field)}»:")
    return EDIT_ARTIST_VALUE

async def editartist_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    data = load_data()
    hod = find_hod(data, context.user_data["edit_artist_hod_id"])
    idx = context.user_data["edit_artist_idx"]
    field = context.user_data["edit_artist_field"]
    hod["artists"][idx][field] = val
    save_data(data)
    sheets_sync_artists(hod)
    await update.message.reply_text(f"✅ Поле «{field}» обновлено.")
    return ConversationHandler.END

async def editartist_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo = update.message.photo[-1]
    data = load_data()
    hod = find_hod(data, context.user_data["edit_artist_hod_id"])
    idx = context.user_data["edit_artist_idx"]
    hod["artists"][idx]["photo"] = photo.file_id
    save_data(data)
    sheets_sync_artists(hod)
    await update.message.reply_text("✅ Фото обновлено.")
    return ConversationHandler.END

# ─── Удаление артиста (/deleteartist) ────────────────────────────────────────

@require_access
async def deleteartist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if len(args) < 2 or not args[0].isdigit() or not args[1].isdigit():
        data = load_data()
        await update.message.reply_text(
            "Выбери ход:",
            reply_markup=_hod_picker_keyboard(data, "pick_deleteartist_", only_with_artists=True)
        )
        return

    hod_id, artist_num = int(args[0]), int(args[1])
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await update.message.reply_text(f"❌ Ход #{hod_id} не найден.")
        return

    if artist_num < 1 or artist_num > len(hod["artists"]):
        await update.message.reply_text(f"❌ Неверный номер артиста. В ходе #{hod_id} артистов: {len(hod['artists'])}")
        return

    removed = hod["artists"].pop(artist_num - 1)
    save_data(data)
    sheets_sync_artists(hod)
    await update.message.reply_text(f"✅ Артист *{removed['name']}* удалён из хода #{hod_id}.", parse_mode="Markdown")

# ─── Удаление хода (/deletehod) ────────────────────────────────────────────

@require_access
async def deletehod_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args or not args[0].isdigit():
        data = load_data()
        if not data["hods"]:
            await update.message.reply_text("Нет ходов для удаления.")
            return
        # Показываем все ходы включая прошедшие
        await update.message.reply_text(
            "Выбери ход для удаления:",
            reply_markup=_hod_picker_keyboard(data, "pick_deletehod_")
        )
        return

    hod_id = int(args[0])
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await update.message.reply_text(f"❌ Ход #{hod_id} не найден.")
        return

    date_str = hod.get("date") or "без даты"
    lineup_str = f"\nЛайнап: {hod['lineup']}" if hod.get("lineup") else ""
    artists_str = f"\nАртистов: {len(hod['artists'])}" if hod.get("artists") else ""

    keyboard = [[
        InlineKeyboardButton("🗑 Да, удалить", callback_data=f"deletehod_confirm_{hod_id}"),
        InlineKeyboardButton("❌ Отмена", callback_data="deletehod_cancel"),
    ]]
    await update.message.reply_text(
        f"Удалить ход #{hod_id}?\n\n"
        f"📅 {date_str}{lineup_str}{artists_str}\n\n"
        "Это действие необратимо.",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def deletehod_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "deletehod_cancel":
        await query.edit_message_text("Отмена. Ход не удалён.")
        return

    hod_id = int(query.data.replace("deletehod_confirm_", ""))
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await query.edit_message_text("Ход не найден — возможно, уже удалён.")
        return

    data["hods"] = [h for h in data["hods"] if h["id"] != hod_id]
    save_data(data)
    sheets_delete_hod(hod_id)
    await query.edit_message_text(f"🗑 Ход #{hod_id} удалён.")

# ─── Управление доступом ─────────────────────────────────────────────────────

async def adduser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Только владелец может добавлять пользователей."""
    if not is_owner(update.effective_user.id):
        await update.message.reply_text("⛔ Только администратор может добавлять пользователей.")
        return

    args = context.args
    if not args or not args[0].lstrip("-").isdigit():
        await update.message.reply_text(
            "Использование: /adduser <user_id> [имя]\n"
            "Узнать свой id: @userinfobot\n\n"
            "Например: /adduser 123456789 Маша"
        )
        return

    user_id = int(args[0])
    name = " ".join(args[1:]) if len(args) > 1 else str(user_id)

    data = load_data()
    if "allowed_users" not in data:
        data["allowed_users"] = []
    if "user_names" not in data:
        data["user_names"] = {}

    if user_id not in data["allowed_users"]:
        data["allowed_users"].append(user_id)
    data["user_names"][str(user_id)] = name
    save_data(data)

    await update.message.reply_text(f"✅ Пользователь *{name}* (id: `{user_id}`) добавлен.", parse_mode="Markdown")

async def removeuser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Только владелец может удалять пользователей."""
    if not is_owner(update.effective_user.id):
        await update.message.reply_text("⛔ Только администратор может удалять пользователей.")
        return

    args = context.args
    if not args or not args[0].lstrip("-").isdigit():
        await update.message.reply_text("Использование: /removeuser <user_id>")
        return

    user_id = int(args[0])
    data = load_data()
    allowed = data.get("allowed_users", [])

    if user_id not in allowed:
        await update.message.reply_text(f"Пользователь {user_id} и так не в списке.")
        return

    allowed.remove(user_id)
    data["allowed_users"] = allowed
    name = data.get("user_names", {}).get(str(user_id), str(user_id))
    save_data(data)
    await update.message.reply_text(f"✅ Пользователь *{name}* удалён.", parse_mode="Markdown")

async def listusers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Только владелец видит список участников."""
    if not is_owner(update.effective_user.id):
        await update.message.reply_text("⛔ Только администратор может просматривать список.")
        return

    data = load_data()
    allowed = data.get("allowed_users", [])
    names = data.get("user_names", {})

    if not allowed:
        await update.message.reply_text("Список пуст. Добавь участников командой /adduser.")
        return

    lines = ["👥 *Участники команды:*\n"]
    for uid in allowed:
        name = names.get(str(uid), "—")
        lines.append(f"  • {name} — `{uid}`")
    lines.append(f"\n_Итого: {len(allowed)}_")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def myid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Любой может узнать свой id — чтобы передать владельцу."""
    user = update.effective_user
    await update.message.reply_text(
        f"Твой Telegram ID: `{user.id}`\n"
        f"Имя: {user.full_name}\n\n"
        "Отправь этот id администратору — он добавит тебя командой /adduser.",
        parse_mode="Markdown"
    )

# ─── /cancel ─────────────────────────────────────────────────────────────────

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Отменено.")
    return ConversationHandler.END

# ─── Обработчик кнопок главной клавиатуры ────────────────────────────────────

@require_access
async def keyboard_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Не обрабатываем если пользователь в активном диалоге
    if context.user_data.get("_in_conversation"):
        return
    text = update.message.text

    if text == "➕ Новый ход":
        return await newhod_start(update, context)
    elif text == "📋 Список ходов":
        return await list_hods(update, context)
    elif text == "📅 Расписание":
        return await schedule_start(update, context)
    elif text == "🎤 Управление артистами":
        await update.message.reply_text(
            "Выбери действие с артистами:",
            reply_markup=ARTISTS_KEYBOARD
        )
    elif text == "➕ Добавить артиста":
        await update.message.reply_text(
            "Введи номер хода: /addartist <id>\nНапример: /addartist 3"
        )
    elif text == "✏️ Редактировать артиста":
        await update.message.reply_text(
            "Введи номер хода: /editartist <id>\nНапример: /editartist 3"
        )
    elif text == "🗑 Удалить артиста":
        await update.message.reply_text(
            "Введи номер хода и номер артиста: /deleteartist <id_хода> <номер>\nНапример: /deleteartist 3 1"
        )
    elif text == "◀️ Назад":
        await update.message.reply_text(
            "Главное меню:",
            reply_markup=MAIN_KEYBOARD
        )

# ─── Периодическая синхронизация из Google Sheets ────────────────────────────

async def _periodic_sync():
    """Фоновая задача: каждые 6 часов читает таблицу и обновляет data.json."""
    import asyncio
    while True:
        await asyncio.sleep(6 * 60 * 60)
        logger.info("Запуск плановой синхронизации из Google Sheets...")
        try:
            restored = sheets_restore_to_json()
            if restored:
                logger.info("Плановая синхронизация завершена успешно.")
            else:
                logger.info("Плановая синхронизация: таблица пуста или недоступна.")
        except Exception as e:
            logger.error(f"Ошибка плановой синхронизации: {e}")


# ─── Обработчики пикера ходов ─────────────────────────────────────────────────

async def pick_show_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    val = query.data.replace("pick_show_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return
    context.args = [val]
    context.user_data["_pick_message"] = query.message
    await query.edit_message_text(f"Ход #{val}:")
    hod_id = int(val)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await query.message.reply_text("Ход не найден.")
        return
    keyboard = []
    if hod.get("status") == "confirmed" and hod.get("artists"):
        keyboard.append([InlineKeyboardButton("📤 Отправить пак", callback_data=f"sendpack_{hod_id}")])
    await query.message.reply_text(
        format_hod(hod),
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard) if keyboard else None
    )


async def pick_edit_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    val = query.data.replace("pick_edit_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return
    hod_id = int(val)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await query.edit_message_text("Ход не найден.")
        return
    context.user_data["edit_hod_id"] = hod_id
    buttons = [
        [InlineKeyboardButton("📅 Дата", callback_data="edit_field_date"),
         InlineKeyboardButton("✏️ Лайнап", callback_data="edit_field_lineup")],
    ]
    if hod.get("date") and hod.get("status") != "confirmed":
        buttons.append([InlineKeyboardButton("🌀 Убрать дату → сделать абстрактным", callback_data="edit_field_removedate")])
    buttons.append([InlineKeyboardButton("❌ Отмена", callback_data="edit_cancel")])
    await query.edit_message_text(
        f"Редактируем ход #{hod_id}. Что изменить?",
        reply_markup=InlineKeyboardMarkup(buttons)
    )


async def pick_addartist_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    val = query.data.replace("pick_addartist_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return
    await query.edit_message_text(f"Добавляем артиста к ходу #{val}. Введи имя артиста:")
    context.user_data["artist_hod_id"] = int(val)
    context.user_data["new_artist"] = {}
    # Продолжаем диалог addartist со следующего шага
    context.user_data["_awaiting_artist_name"] = True


async def pick_editartist_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    val = query.data.replace("pick_editartist_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return
    hod_id = int(val)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod or not hod.get("artists"):
        await query.edit_message_text("У этого хода нет артистов.")
        return
    context.user_data["edit_artist_hod_id"] = hod_id
    keyboard = [
        [InlineKeyboardButton(f"{i+1}. {a['name']}", callback_data=f"ea_select_{i}")]
        for i, a in enumerate(hod["artists"])
    ]
    keyboard.append([InlineKeyboardButton("❌ Отмена", callback_data="ea_cancel")])
    await query.edit_message_text(
        "Выбери артиста для редактирования:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def pick_confirm_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    val = query.data.replace("pick_confirm_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return
    hod_id = int(val)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await query.edit_message_text("Ход не найден.")
        return
    hod["status"] = "confirmed"
    hod["confirmed_by"] = query.from_user.full_name
    hod["confirmed_at"] = datetime.now().isoformat()
    save_data(data)
    sheets_upsert_hod(hod)
    await query.edit_message_text(
        f"✅ Ход #{hod_id} подтверждён!\n\n{format_hod(hod)}",
        parse_mode="Markdown"
    )


async def pick_deletehod_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    val = query.data.replace("pick_deletehod_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return
    hod_id = int(val)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod:
        await query.edit_message_text("Ход не найден.")
        return
    date_str = hod.get("date") or "без даты"
    lineup_str = f"\nЛайнап: {hod['lineup']}" if hod.get("lineup") else ""
    artists_str = f"\nАртистов: {len(hod['artists'])}" if hod.get("artists") else ""
    keyboard = [[
        InlineKeyboardButton("🗑 Да, удалить", callback_data=f"deletehod_confirm_{hod_id}"),
        InlineKeyboardButton("❌ Отмена", callback_data="deletehod_cancel"),
    ]]
    await query.edit_message_text(
        f"Удалить ход #{hod_id}?\n\n📅 {date_str}{lineup_str}{artists_str}\n\nЭто действие необратимо.",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def pick_deleteartist_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Шаг 1: выбран ход — показываем список артистов."""
    query = update.callback_query
    await query.answer()
    val = query.data.replace("pick_deleteartist_", "")
    if val == "cancel":
        await query.edit_message_text("Отмена.")
        return
    hod_id = int(val)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod or not hod.get("artists"):
        await query.edit_message_text("У этого хода нет артистов.")
        return
    keyboard = [
        [InlineKeyboardButton(f"🗑 {a['name']}", callback_data=f"da_artist_{hod_id}_{i}")]
        for i, a in enumerate(hod["artists"])
    ]
    keyboard.append([InlineKeyboardButton("❌ Отмена", callback_data="da_cancel")])
    await query.edit_message_text(
        f"Выбери артиста для удаления из хода #{hod_id}:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def da_artist_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Шаг 2: выбран артист — удаляем с подтверждением."""
    query = update.callback_query
    await query.answer()
    if query.data == "da_cancel":
        await query.edit_message_text("Отмена.")
        return
    _, _, hod_id_str, idx_str = query.data.split("_", 3)
    hod_id, idx = int(hod_id_str), int(idx_str)
    data = load_data()
    hod = find_hod(data, hod_id)
    if not hod or idx >= len(hod["artists"]):
        await query.edit_message_text("Артист не найден.")
        return
    removed = hod["artists"].pop(idx)
    save_data(data)
    sheets_sync_artists(hod)
    await query.edit_message_text(f"✅ Артист *{removed['name']}* удалён из хода #{hod_id}.", parse_mode="Markdown")

# ─── Запуск ──────────────────────────────────────────────────────────────────

def main():
    app = Application.builder().token(BOT_TOKEN).build()

    # Диалог: добавление хода
    add_hod_conv = ConversationHandler(
        entry_points=[CommandHandler("newhod", newhod_start)],
        states={
            ADD_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, newhod_date)],
            ADD_LINEUP: [MessageHandler(filters.TEXT & ~filters.COMMAND, newhod_lineup)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    # Диалог: редактирование хода
    edit_hod_conv = ConversationHandler(
        entry_points=[CommandHandler("edit", edit_start)],
        states={
            EDIT_CHOOSE: [CallbackQueryHandler(edit_field_callback, pattern="^edit_field_|^edit_cancel$")],
            EDIT_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_date)],
            EDIT_LINEUP: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_lineup)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    # Диалог: добавление артиста
    add_artist_conv = ConversationHandler(
        entry_points=[CommandHandler("addartist", addartist_start)],
        states={
            ARTIST_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, artist_name)],
            ARTIST_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, artist_time)],
            ARTIST_CONTACT: [MessageHandler(filters.TEXT & ~filters.COMMAND, artist_contact)],
            ARTIST_SOCIAL: [MessageHandler((filters.TEXT | filters.Entity("url") | filters.Entity("text_link")) & ~filters.COMMAND, artist_social)],
            ARTIST_PHOTO: [
                MessageHandler(filters.PHOTO, artist_photo_file),
                MessageHandler(filters.TEXT & ~filters.COMMAND, artist_photo_text),
            ],
            ARTIST_COMMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, artist_comment)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    # Диалог: редактирование артиста
    edit_artist_conv = ConversationHandler(
        entry_points=[CommandHandler("editartist", editartist_start)],
        states={
            EDIT_ARTIST_HOD_PICK: [CallbackQueryHandler(editartist_hod_pick, pattern="^ea_hod_")],
            EDIT_ARTIST_SELECT: [CallbackQueryHandler(editartist_select, pattern="^ea_select_|^ea_cancel$")],
            EDIT_ARTIST_FIELD: [CallbackQueryHandler(editartist_field, pattern="^ea_field_|^ea_cancel$")],
            EDIT_ARTIST_VALUE: [MessageHandler(filters.TEXT & ~filters.COMMAND, editartist_value)],
            EDIT_ARTIST_PHOTO: [MessageHandler(filters.PHOTO, editartist_photo)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    # Диалог: расписание
    schedule_conv = ConversationHandler(
        entry_points=[CommandHandler("schedule", schedule_start)],
        states={
            SCHED_PERIOD: [CallbackQueryHandler(schedule_period_callback, pattern="^sched_")],
            SCHED_FROM: [MessageHandler(filters.TEXT & ~filters.COMMAND, schedule_from)],
            SCHED_TO: [MessageHandler(filters.TEXT & ~filters.COMMAND, schedule_to)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    # ConversationHandler-ы первыми — они перехватывают сообщения во время диалога
    app.add_handler(add_hod_conv)
    app.add_handler(edit_hod_conv)
    app.add_handler(add_artist_conv)
    app.add_handler(edit_artist_conv)
    app.add_handler(schedule_conv)
    # Команды
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("list", list_hods))
    app.add_handler(CommandHandler("showhod", showhod))
    app.add_handler(CommandHandler("confirm", confirm_hod))
    app.add_handler(CommandHandler("deleteartist", deleteartist))
    app.add_handler(CommandHandler("deletehod", deletehod_start))
    app.add_handler(CommandHandler("adduser", adduser))
    app.add_handler(CommandHandler("removeuser", removeuser))
    app.add_handler(CommandHandler("listusers", listusers))
    app.add_handler(CommandHandler("myid", myid))
    # Callback-кнопки
    app.add_handler(CallbackQueryHandler(deletehod_callback, pattern="^deletehod_"))
    app.add_handler(CallbackQueryHandler(filter_callback, pattern="^filter_"))
    app.add_handler(CallbackQueryHandler(confirm_callback, pattern="^confirm_"))
    app.add_handler(CallbackQueryHandler(artists_callback, pattern="^artists_"))
    app.add_handler(CallbackQueryHandler(pick_show_callback, pattern="^pick_show_"))
    app.add_handler(CallbackQueryHandler(pick_edit_callback, pattern="^pick_edit_"))
    app.add_handler(CallbackQueryHandler(pick_addartist_callback, pattern="^pick_addartist_"))
    app.add_handler(CallbackQueryHandler(pick_editartist_callback, pattern="^pick_editartist_"))
    app.add_handler(CallbackQueryHandler(pick_confirm_callback, pattern="^pick_confirm_"))
    app.add_handler(CallbackQueryHandler(pick_deletehod_callback, pattern="^pick_deletehod_"))
    app.add_handler(CallbackQueryHandler(pick_deleteartist_callback, pattern="^pick_deleteartist_"))
    app.add_handler(CallbackQueryHandler(da_artist_callback, pattern="^da_artist_|^da_cancel$"))
    app.add_handler(CallbackQueryHandler(sendpack_callback, pattern="^sendpack_"))
    # Кнопки клавиатуры — последними, чтобы не перехватывать текст во время диалогов
    app.add_handler(MessageHandler(
        filters.TEXT & filters.Regex("^(➕ Новый ход|📋 Список ходов|📅 Расписание|🎤 Управление артистами|➕ Добавить артиста|✏️ Редактировать артиста|🗑 Удалить артиста|◀️ Назад)$"),
        keyboard_handler
    ), group=1)

    print("Бот запущен...")
    print(f"OWNER_ID из переменной окружения: {repr(os.getenv('OWNER_ID'))}")
    print(f"OWNER_ID после обработки: {OWNER_ID}")
    # Если data.json пустой или отсутствует — восстанавливаем из Google Sheets
    current_data = load_data()
    if not current_data.get("hods"):
        print("data.json пустой — пробую восстановить из Google Sheets...")
        restored = sheets_restore_to_json()
        if restored:
            print("Данные восстановлены из Google Sheets.")
            current_data = load_data()
        else:
            print("Восстановление не удалось или таблица пуста — начинаем с нуля.")
    # Синхронизируем таблицу при старте
    sheets_full_sync(current_data)
    # Запускаем обратную синхронизацию каждые 6 часов
    import asyncio
    loop = asyncio.get_event_loop()
    loop.create_task(_periodic_sync())
    app.run_polling()

if __name__ == "__main__":
    main()
