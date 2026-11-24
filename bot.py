import os
import sqlite3
import logging
import calendar
from datetime import datetime, timedelta, timezone
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove,
    InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile, CallbackQuery,
    BotCommand, BotCommandScopeChat, BotCommandScopeDefault
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

# ===== ЛОГИРОВАНИЕ И НАСТРОЙКИ =====
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x]
LEAD_IDS = [int(x) for x in os.getenv("LEAD_IDS", "").split(",") if x]
DATA_DIR = os.getenv("DATA_DIR", ".")
DB_FILE = os.path.join(DATA_DIR, "bot.db")
EXCEL_FILE = os.path.join(DATA_DIR, "report.xlsx")
MSK = timezone(timedelta(hours=3))

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
router = Router()
dp.include_router(router)

# ===== БАЗА ДАННЫХ =====
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS users (tg_id INTEGER PRIMARY KEY, full_name TEXT NOT NULL)""")
    c.execute("""CREATE TABLE IF NOT EXISTS data (
        id INTEGER PRIMARY KEY AUTOINCREMENT, tg_id INTEGER, date TEXT,
        metric TEXT, value REAL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(tg_id) REFERENCES users(tg_id))""")
    conn.commit()
    conn.close()

init_db()

def get_user(tg_id):
    conn = sqlite3.connect(DB_FILE)
    user = conn.execute("SELECT * FROM users WHERE tg_id = ?", (tg_id,)).fetchone()
    conn.close()
    return user

def get_role(tg_id):
    if tg_id in ADMIN_IDS: return "admin"
    if tg_id in LEAD_IDS: return "lead"
    return "user"

# ===== МЕТРИКИ И СОСТОЯНИЯ =====
METRICS_ORDER = [
    ("📋 ДК заявки", "dk_zayavki"), ("📦 ДК выдача", "dk_vydacha"),
    ("📄 ИДК заявки", "idk_zayavki"), ("📤 ИДК выдача", "idk_vydacha"),
    ("🧬 Био", "bio"),
]
METRIC_NAMES = {v: k for k, v in METRICS_ORDER}

class InputState(StatesGroup):
    waiting_name = State()
    waiting_date = State()
    waiting_value = State()
    final_confirm = State()
    editing = State()
    summary_start_date = State()
    summary_end_date = State()

def main_menu():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="📝 Внести данные")],
        [KeyboardButton(text="📊 Мои данные"), KeyboardButton(text="✏️ Мои записи")]
    ], resize_keyboard=True)

# ===== РЕГИСТРАЦИЯ И ВВОД =====
@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    logger.info(f"Пользователь {message.from_user.id} ({message.from_user.username}) нажал /start")
    user = get_user(message.from_user.id)
    if user:
        await message.answer(f"Привет, {user[1]}! 👋", reply_markup=main_menu())
    else:
        await message.answer("Привет! Введи свои ФИО для регистрации:")
        await state.set_state(InputState.waiting_name)

@router.message(InputState.waiting_name)
async def save_name(message: Message, state: FSMContext):
    if message.text.strip().startswith("/"):
        await message.answer("❗ Это команда, а не имя. Введи ФИО обычным текстом:")
        return
    conn = sqlite3.connect(DB_FILE)
    conn.execute("INSERT INTO users (tg_id, full_name) VALUES (?, ?)", (message.from_user.id, message.text.strip()))
    conn.commit()
    conn.close()
    await state.clear()
    await message.answer(f"✅ Зарегистрирован: {message.text}", reply_markup=main_menu())

@router.message(F.text == "📝 Внести данные")
async def start_input(message: Message, state: FSMContext):
    if not get_user(message.from_user.id):
        await message.answer("Сначала зарегистрируйся: /start")
        return
    await message.answer("📅 Введи дату (ДД.ММ.ГГГГ) или напиши 'сегодня' / 'вчера':",
                         reply_markup=ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="сегодня")], [KeyboardButton(text="вчера")]], resize_keyboard=True))
    await state.set_state(InputState.waiting_date)

@router.message(InputState.waiting_date)
async def get_date(message: Message, state: FSMContext):
    text = message.text.strip().lower()
    now = datetime.now(MSK)
    if text == "сегодня":
        date_str = now.strftime("%d.%m.%Y")
    elif text == "вчера":
        date_str = (now - timedelta(days=1)).strftime("%d.%m.%Y")
    else:
        try:
            parsed = datetime.strptime(text, "%d.%m.%Y")
            if parsed.replace(tzinfo=MSK) > now.replace(hour=23, minute=59, second=59):
                await message.answer("❌ Дата не может быть в будущем.")
                return
            date_str = parsed.strftime("%d.%m.%Y")
        except ValueError:
            await message.answer("❌ Формат ДД.ММ.ГГГГ:")
            return

    conn = sqlite3.connect(DB_FILE)
    existing = conn.execute("SELECT metric, value FROM data WHERE tg_id = ? AND date = ?", (message.from_user.id, date_str)).fetchall()
    conn.close()
    is_replace = bool(existing)
    if is_replace:
        old = {m: int(v) for m, v in existing}
        await message.answer(f"⚠️ За {date_str} уже есть данные:\nДК {old.get('dk_zayavki',0)}/{old.get('dk_vydacha',0)} | ИДК {old.get('idk_zayavki',0)}/{old.get('idk_vydacha',0)} | Био {old.get('bio',0)}\n\nНовый ввод ЗАМЕНИТ эти значения.")

    await state.update_data(date=date_str, values={}, metric_index=0, is_replace=is_replace)
    await state.set_state(InputState.waiting_value)
    metric_label, _ = METRICS_ORDER[0]
    await message.answer(f"📊 Дата: {date_str}\n\n1️⃣ Введи: {metric_label}", reply_markup=ReplyKeyboardRemove())

@router.message(InputState.waiting_value)
async def get_value(message: Message, state: FSMContext):
    text = message.text.strip().replace(",", ".")
    try:
        value = float(text)
        if value < 0:
            await message.answer("❌ Значение не может быть отрицательным.")
            return
    except ValueError:
        await message.answer("❌ Введи число:")
        return

    data = await state.get_data()
    metric_index = data["metric_index"]
    metric_label, metric_key = METRICS_ORDER[metric_index]
    values = data.get("values", {})

    if metric_key == "dk_vydacha" and value > values.get("dk_zayavki", 0):
        await message.answer(f"⚠️ ДК выдача ({int(value)}) > ДК заявок ({int(values.get('dk_zayavki',0))}). Выдача не может быть больше заявок. Введи ≤ {int(values.get('dk_zayavki',0))}:")
        return
    if metric_key == "idk_vydacha" and value > values.get("idk_zayavki", 0):
        await message.answer(f"⚠️ ИДК выдача ({int(value)}) > ИДК заявок ({int(values.get('idk_zayavki',0))}). Выдача не может быть больше заявок. Введи ≤ {int(values.get('idk_zayavki',0))}:")
        return

    values[metric_key] = value
    await state.update_data(values=values)
    metric_index += 1

    if metric_index < len(METRICS_ORDER):
        await state.update_data(metric_index=metric_index)
        metric_label, _ = METRICS_ORDER[metric_index]
        step_names = ["2️⃣", "3️⃣", "4️⃣", "5️⃣"]
        await message.answer(f"{step_names[metric_index - 1]} Введи: {metric_label}")
    else:
        await show_final_confirm(message, state, values)

async def show_final_confirm(message: Message, state: FSMContext, values: dict):
    data = await state.get_data()
    date_str = data["date"]
    is_replace = data.get("is_replace", False)
    question = "⚠️ Старые данные за эту дату будут заменены. Продолжить?" if is_replace else "Сохранить?"
    save_label = "💾 Заменить" if is_replace else "💾 Сохранить"
    text = (f"📋 Данные за {date_str}:\n\n"
            f"🔹 ДК: {int(values.get('dk_zayavki', 0))}(заявки) / {int(values.get('dk_vydacha', 0))}(выдано)\n"
            f"🔹 ИДК: {int(values.get('idk_zayavki', 0))}(заявки) / {int(values.get('idk_vydacha', 0))}(выдано)\n"
            f"🔹 Био: {int(values.get('bio', 0))}\n\n{question}")
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=save_label, callback_data="save_all"),
         InlineKeyboardButton(text="❌ Отменить", callback_data="cancel_all")]
    ])
    await message.answer(text, reply_markup=keyboard)
    await state.set_state(InputState.final_confirm)

@router.callback_query(F.data == "save_all", InputState.final_confirm)
async def save_all(callback: CallbackQuery, state: FSMContext):
    await callback.answer()
    data = await state.get_data()
    date_str = data["date"]
    values = data.get("values", {})
    is_replace = data.get("is_replace", False)
    conn = sqlite3.connect(DB_FILE)
    for metric_key, value in values.items():
        cur = conn.execute("UPDATE data SET value = ? WHERE tg_id = ? AND date = ? AND metric = ?", (value, callback.from_user.id, date_str, metric_key))
        if cur.rowcount == 0:
            conn.execute("INSERT INTO data (tg_id, date, metric, value) VALUES (?, ?, ?, ?)", (callback.from_user.id, date_str, metric_key, value))
    conn.commit()
    conn.close()
    await state.clear()
    msg = f"✅ Данные за {date_str} заменены!" if is_replace else f"✅ Сохранено за {date_str}!"
    await callback.message.answer(msg, reply_markup=main_menu())

@router.callback_query(F.data == "cancel_all", InputState.final_confirm)
async def cancel_all(callback: CallbackQuery, state: FSMContext):
    await callback.answer("❌ Отменено")
    await state.clear()
    await callback.message.answer("Ввод отменён.", reply_markup=main_menu())

@router.message(InputState.final_confirm)
async def confirm_by_text(message: Message, state: FSMContext):
    text = message.text.strip().lower()
    if text in ("сохранить", "💾 сохранить", "заменить", "💾 заменить", "да"):
        data = await state.get_data()
        date_str = data["date"]
        values = data.get("values", {})
        is_replace = data.get("is_replace", False)
        conn = sqlite3.connect(DB_FILE)
        for metric_key, value in values.items():
            cur = conn.execute("UPDATE data SET value = ? WHERE tg_id = ? AND date = ? AND metric = ?", (value, message.from_user.id, date_str, metric_key))
            if cur.rowcount == 0:
                conn.execute("INSERT INTO data (tg_id, date, metric, value) VALUES (?, ?, ?, ?)", (message.from_user.id, date_str, metric_key, value))
        conn.commit()
        conn.close()
        await state.clear()
        msg = f"✅ Данные за {date_str} заменены!" if is_replace else f"✅ Сохранено за {date_str}!"
        await message.answer(msg, reply_markup=main_menu())
    elif text in ("отменить", "❌ отменить", "нет"):
        await state.clear()
        await message.answer("Ввод отменён.", reply_markup=main_menu())
    else:
        await message.answer("Нажми кнопку 💾 Сохранить под сводкой или напиши «Сохранить» / «Отменить».")

# ===== ПРОСМОТР И РЕДАКТИРОВАНИЕ =====
@router.message(F.text == "📊 Мои данные")
async def my_data(message: Message):
    user = get_user(message.from_user.id)
    if not user:
        await message.answer("Сначала зарегистрируйся: /start")
        return
    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("SELECT date, metric, value FROM data WHERE tg_id = ?", (message.from_user.id,)).fetchall()
    conn.close()
    now = datetime.now(MSK)
    start = (now - timedelta(days=6)).strftime("%d.%m.%Y")
    end = now.strftime("%d.%m.%Y")
    rows = filter_data_by_date(rows, start, end)
    d = {}
    for date, metric, value in rows:
        d[metric] = d.get(metric, 0) + int(value)
    text = (f"📊 {user[1]}, сводка за неделю:\n\n"
            f"🔹 ДК: {d.get('dk_zayavki', 0)}(заявки) / {d.get('dk_vydacha', 0)}(выдано)\n"
            f"🔹 ИДК: {d.get('idk_zayavki', 0)}(заявки) / {d.get('idk_vydacha', 0)}(выдано)\n"
            f"🔹 Био: {d.get('bio', 0)}")
    await message.answer(text)

@router.message(F.text == "✏️ Мои записи")
async def show_my_records(message: Message):
    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("SELECT id, date, metric, value FROM data WHERE tg_id = ? ORDER BY created_at DESC LIMIT 5", (message.from_user.id,)).fetchall()
    conn.close()
    if not rows:
        await message.answer("📭 Записей нет.")
        return
    text = "📝 Твои последние 5 записей:\n\n"
    keyboard = InlineKeyboardMarkup(inline_keyboard=[])
    for row in rows:
        rec_id, date, metric, value = row
        metric_name = METRIC_NAMES.get(metric, metric)
        v = int(value) if value == int(value) else value
        text += f"• {date} | {metric_name}: {v}\n"
        keyboard.inline_keyboard.append([
            InlineKeyboardButton(text=f"✏️ {date} {metric_name}", callback_data=f"edit_{rec_id}"),
            InlineKeyboardButton(text="🗑️", callback_data=f"delete_{rec_id}")
        ])
    await message.answer(text, reply_markup=keyboard)

@router.callback_query(F.data.startswith("edit_"))
async def start_edit(callback: CallbackQuery, state: FSMContext):
    rec_id = int(callback.data.replace("edit_", ""))
    await state.update_data(edit_id=rec_id, is_admin_edit=False)
    await callback.message.answer("Введи новое значение:")
    await state.set_state(InputState.editing)
    await callback.answer()

@router.message(InputState.editing)
async def save_edit(message: Message, state: FSMContext):
    try:
        value = float(message.text.replace(",", "."))
    except ValueError:
        await message.answer("❌ Введи число:")
        return
    data = await state.get_data()
    rec_id = data.get("edit_id")
    is_admin = data.get("is_admin_edit", False)
    conn = sqlite3.connect(DB_FILE)
    if is_admin:
        conn.execute("UPDATE data SET value = ? WHERE id = ?", (value, rec_id))
    else:
        conn.execute("UPDATE data SET value = ? WHERE id = ? AND tg_id = ?", (value, rec_id, message.from_user.id))
    conn.commit()
    conn.close()
    await state.clear()
    v = int(value) if value == int(value) else value
    await message.answer(f"✅ Обновлено: {v}", reply_markup=main_menu() if not is_admin else None)

@router.callback_query(F.data.startswith("delete_"))
async def delete_record(callback: CallbackQuery):
    rec_id = int(callback.data.replace("delete_", ""))
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.execute("DELETE FROM data WHERE id = ? AND tg_id = ?", (rec_id, callback.from_user.id))
    conn.commit()
    conn.close()
    await callback.answer("🗑️ Удалено")
    await callback.message.edit_text("✅ Запись удалена." if cursor.rowcount > 0 else "⚠️ Не найдена.")

# ===== АДМИН ФУНКЦИИ =====
@router.message(Command("admin_records"))
async def admin_records(message: Message):
    if get_role(message.from_user.id) != "admin":
        await message.answer("⛔ Только админ.")
        return
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Использование: /admin_records [Telegram ID]\nПример: /admin_records 123456789")
        return
    target_id = int(args[1])
    user = get_user(target_id)
    if not user:
        await message.answer("❌ Пользователь не найден.")
        return
    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("SELECT id, date, metric, value FROM data WHERE tg_id = ? ORDER BY created_at DESC LIMIT 10", (target_id,)).fetchall()
    conn.close()
    if not rows:
        await message.answer(f"📭 У {user[1]} нет записей.")
        return
    text = f"📝 Последние записи {user[1]}:\n\n"
    keyboard = InlineKeyboardMarkup(inline_keyboard=[])
    for row in rows:
        rec_id, date, metric, value = row
        metric_name = METRIC_NAMES.get(metric, metric)
        v = int(value) if value == int(value) else value
        text += f"• {date} | {metric_name}: {v}\n"
        keyboard.inline_keyboard.append([
            InlineKeyboardButton(text=f"✏️ {date} {metric_name}", callback_data=f"admin_edit_{rec_id}"),
            InlineKeyboardButton(text="🗑️", callback_data=f"admin_delete_{rec_id}")
        ])
    await message.answer(text, reply_markup=keyboard)

@router.message(Command("del_user"))
async def del_user(message: Message):
    if get_role(message.from_user.id) != "admin":
        await message.answer("⛔ Только админ.")
        return
    args = message.text.split()
    if len(args) < 2:
        await message.answer("Используй: /del_user [ID]")
        return
    tid = int(args[1])
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM data WHERE tg_id = ?", (tid,))
    conn.execute("DELETE FROM users WHERE tg_id = ?", (tid,))
    conn.commit()
    conn.close()
    await message.answer(f"🗑 Пользователь {tid} и все его записи удалены.")

@router.callback_query(F.data.startswith("admin_edit_"))
async def admin_start_edit(callback: CallbackQuery, state: FSMContext):
    if get_role(callback.from_user.id) != "admin":
        await callback.answer("⛔ Только админ.")
        return
    rec_id = int(callback.data.replace("admin_edit_", ""))
    await state.update_data(edit_id=rec_id, is_admin_edit=True)
    await callback.message.answer("Введи новое значение:")
    await state.set_state(InputState.editing)
    await callback.answer()

@router.callback_query(F.data.startswith("admin_delete_"))
async def admin_delete_record(callback: CallbackQuery):
    if get_role(callback.from_user.id) != "admin":
        await callback.answer("⛔ Только админ.")
        return
    rec_id = int(callback.data.replace("admin_delete_", ""))
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.execute("DELETE FROM data WHERE id = ?", (rec_id,))
    conn.commit()
    conn.close()
    await callback.answer("🗑️ Удалено")
    await callback.message.edit_text("✅ Запись удалена." if cursor.rowcount > 0 else "⚠️ Не найдена.")

# ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ДАТ =====
def filter_data_by_date(all_data, start_date, end_date):
    if not start_date or not end_date:
        return all_data

    def to_cmp(d):
        try:
            p = str(d).split('.')
            if len(p) == 3:
                return f"{p[2]}-{p[1]}-{p[0]}"
        except Exception:
            pass
        return str(d)

    s_cmp, e_cmp = to_cmp(start_date), to_cmp(end_date)
    return [row for row in all_data if s_cmp <= to_cmp(row[0]) <= e_cmp]

# ===== ОБЩАЯ ТАБЛИЦА EXCEL (КОМПАКТНАЯ, С ИТОГАМИ) =====
def generate_excel(start_date=None, end_date=None):
    conn = sqlite3.connect(DB_FILE)
    users = conn.execute("SELECT tg_id, full_name FROM users ORDER BY full_name").fetchall()
    all_data = conn.execute("SELECT tg_id, date, metric, value FROM data").fetchall()
    conn.close()

    filtered = []
    valid_dates = set()
    for tg_id, date_str, metric, value in all_data:
        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
            if start_date and end_date:
                s = datetime.strptime(start_date, "%d.%m.%Y")
                e = datetime.strptime(end_date, "%d.%m.%Y")
                if not (s <= dt <= e):
                    continue
            filtered.append((tg_id, date_str, metric, int(value)))
            valid_dates.add(date_str)
        except ValueError:
            continue

    if not filtered:
        return None

    sorted_dates = sorted(valid_dates, key=lambda x: datetime.strptime(x, "%d.%m.%Y"))
    store = {}
    for tg_id, date_str, metric, value in filtered:
        store.setdefault(tg_id, {}).setdefault(date_str, {})[metric] = value

    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет"

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    date_fill = PatternFill("solid", fgColor="2F5597")
    head_fill = PatternFill("solid", fgColor="4472C4")
    total_fill = PatternFill("solid", fgColor="FFE699")
    wb_font = Font(bold=True, color="FFFFFF", size=11)
    total_font = Font(bold=True, size=11)
    center = Alignment(horizontal="center", vertical="center")

    # Строка 1: даты
    ws.cell(row=1, column=1, value="").fill = PatternFill("solid", fgColor="000000")
    col = 2
    for d in sorted_dates:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        c = ws.cell(row=1, column=col, value=d)
        c.fill = date_fill
        c.font = wb_font
        c.alignment = center
        for i in range(3):
            ws.cell(row=1, column=col + i).border = border
        col += 3

    # Строка 2: заголовки
    c = ws.cell(row=2, column=1, value="ФИО")
    c.fill = head_fill
    c.font = wb_font
    c.border = border
    col = 2
    for _ in sorted_dates:
        for name in ("ДК", "ИДК", "Био"):
            c = ws.cell(row=2, column=col, value=name)
            c.fill = head_fill
            c.font = wb_font
            c.alignment = center
            c.border = border
            col += 1

    # Строки сотрудников
    row = 3
    for tg_id, full_name in users:
        c = ws.cell(row=row, column=1, value=full_name)
        c.font = Font(bold=True, size=11)
        c.border = border
        col = 2
        for d in sorted_dates:
            v = store.get(tg_id, {}).get(d, {})
            dk = f"{v.get('dk_zayavki', 0)}/{v.get('dk_vydacha', 0)}" if ('dk_zayavki' in v or 'dk_vydacha' in v) else ""
            idk = f"{v.get('idk_zayavki', 0)}/{v.get('idk_vydacha', 0)}" if ('idk_zayavki' in v or 'idk_vydacha' in v) else ""
            bio = v.get("bio", "")
            for i, val in enumerate((dk, idk, bio)):
                c = ws.cell(row=row, column=col + i, value=val)
                c.alignment = center
                c.border = border
            col += 3
        row += 1

    # Итого за день
    c = ws.cell(row=row, column=1, value="Итого за день")
    c.font = total_font
    c.fill = total_fill
    c.border = border
    grand = [0, 0, 0, 0, 0]
    col = 2
    for d in sorted_dates:
        s = [0, 0, 0, 0, 0]
        for tg_id, _ in users:
            v = store.get(tg_id, {}).get(d, {})
            s[0] += v.get("dk_zayavki", 0)
            s[1] += v.get("dk_vydacha", 0)
            s[2] += v.get("idk_zayavki", 0)
            s[3] += v.get("idk_vydacha", 0)
            s[4] += v.get("bio", 0)
        for i in range(5):
            grand[i] += s[i]
        for i, val in enumerate([f"{s[0]}/{s[1]}", f"{s[2]}/{s[3]}", s[4]]):
            c = ws.cell(row=row, column=col + i, value=val)
            c.font = total_font
            c.fill = total_fill
            c.alignment = center
            c.border = border
        col += 3
    row += 1

    # Итого за период / месяц
    label = "Итого за период"
    if sorted_dates:
        d0 = datetime.strptime(sorted_dates[0], "%d.%m.%Y")
        d1 = datetime.strptime(sorted_dates[-1], "%d.%m.%Y")
        month_end = (d1.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        if d0.day == 1 and d1 == month_end and d0.month == d1.month:
            label = "Итого за месяц"
    c = ws.cell(row=row, column=1, value=label)
    c.font = total_font
    c.fill = total_fill
    c.border = border
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=1 + 3 * len(sorted_dates))
    c = ws.cell(row=row, column=2, value=f"ДК: {grand[0]}/{grand[1]}   |   ИДК: {grand[2]}/{grand[3]}   |   Био: {grand[4]}")
    c.font = total_font
    c.fill = total_fill
    c.alignment = center
    c.border = border
    row += 1

    # Обрезка лишнего и чистый вид
    if ws.max_row >= row:
        ws.delete_rows(row, ws.max_row - row + 1)
    ws.sheet_view.showGridLines = False

    max_name = max([len(u[1]) for u in users] + [3])
    ws.column_dimensions["A"].width = min(max_name + 3, 45)
    for i in range(2, 2 + 3 * len(sorted_dates)):
        ws.column_dimensions[get_column_letter(i)].width = 10
    ws.freeze_panes = "B3"

    wb.save(EXCEL_FILE)
    return EXCEL_FILE

# ===== ПЕРСОНАЛЬНЫЙ EXCEL =====
def generate_excel_for_user(tg_id, start_date=None, end_date=None):
    conn = sqlite3.connect(DB_FILE)
    user = conn.execute("SELECT full_name FROM users WHERE tg_id = ?", (tg_id,)).fetchone()
    all_data = conn.execute("SELECT date, metric, value FROM data WHERE tg_id = ?", (tg_id,)).fetchall()
    conn.close()
    if not user or not all_data:
        return None, None
    all_data = filter_data_by_date(all_data, start_date, end_date)
    if not all_data:
        return None, None

    dates = sorted(set(r[0] for r in all_data), key=lambda x: datetime.strptime(x, "%d.%m.%Y"))
    sums = {}
    for date, metric, value in all_data:
        sums.setdefault(metric, {})[date] = sums.get(metric, {}).get(date, 0) + int(value)

    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет"
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="4472C4")
    head_font = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center")

    c = ws.cell(row=1, column=1, value="Показатель")
    c.font = head_font
    c.fill = head_fill
    c.border = border
    for i, d in enumerate(dates):
        c = ws.cell(row=1, column=2 + i, value=d)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = border
    r = 2
    for label, key in METRICS_ORDER:
        c = ws.cell(row=r, column=1, value=label)
        c.border = border
        for i, d in enumerate(dates):
            c = ws.cell(row=r, column=2 + i, value=sums.get(key, {}).get(d, 0))
            c.alignment = center
            c.border = border
        r += 1
    ws.column_dimensions["A"].width = 20
    for i in range(2, 2 + len(dates)):
        ws.column_dimensions[get_column_letter(i)].width = 12
    ws.sheet_view.showGridLines = False
    path = os.path.join(DATA_DIR, f"report_{tg_id}.xlsx")
    wb.save(path)
    return path, user[0]

# ===== КОМАНДА /export =====
@router.message(Command("export"))
async def export_excel(message: Message):
    if get_role(message.from_user.id) not in ("admin", "lead"):
        await message.answer("⛔ Только руководитель или админ.")
        return
    args = message.text.split()[1:]
    target = None
    if args and args[0].lower() == "user":
        if len(args) < 2:
            await message.answer("❓ Укажи ID: /export user 123456789 [week|month]")
            return
        target = int(args[1])
        args = args[2:]

    now = datetime.now(MSK)
    sd = ed = None
    if args:
        a0 = args[0].lower()
        if a0 == "week":
            sd = now - timedelta(days=6)
            ed = now
        elif a0 == "month":
            sd = now.replace(day=1)
            ed = now
        elif len(args) == 2:
            try:
                sd = datetime.strptime(args[0], "%d.%m.%Y")
                ed = datetime.strptime(args[1], "%d.%m.%Y")
            except ValueError:
                await message.answer("❓ Примеры: /export week, /export month, /export Ноябрь-2026, /export 01.11.2026 21.11.2026")
                return
        else:
            try:
                d = datetime.strptime(args[0], "%d.%m.%Y")
                sd = ed = d
            except ValueError:
                months = {"январь":1,"февраль":2,"март":3,"апрель":4,"май":5,"июнь":6,"июль":7,
                          "август":8,"сентябрь":9,"октябрь":10,"ноябрь":11,"декабрь":12}
                part = a0.replace("-", " ").split()
                if not part or part[0] not in months:
                    await message.answer("❓ Примеры: /export week, /export month, /export Ноябрь-2026, /export 01.11.2026 21.11.2026")
                    return
                m = months[part[0]]
                y = int(part[1]) if len(part) > 1 else now.year
                sd = datetime(y, m, 1)
                ed = datetime(y, m, calendar.monthrange(y, m)[1])

    sd_str = sd.strftime("%d.%m.%Y") if sd else None
    ed_str = ed.strftime("%d.%m.%Y") if ed else None

    if target:
        path, name = generate_excel_for_user(target, sd_str, ed_str)
        if not path:
            await message.answer("📭 Нет данных за выбранный период.")
            return
        await message.answer_document(FSInputFile(path), caption=f"📎 Отчёт по: {name}")
        os.remove(path)
    else:
        path = generate_excel(sd_str, ed_str)
        if not path:
            await message.answer("📭 Нет данных за выбранный период.")
            return
        await message.answer_document(FSInputFile(path), caption="📎 Отчёт готов")

# ===== ТЕКСТОВАЯ СВОДКА =====
async def send_summary_text(message_or_callback, start_date: str, end_date: str):
    if hasattr(message_or_callback, "message"):
        msg = message_or_callback.message
    else:
        msg = message_or_callback

    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("""
        SELECT u.full_name, d.date, d.metric, d.value
        FROM data d JOIN users u ON d.tg_id = u.tg_id
    """).fetchall()
    conn.close()

    def to_cmp(d):
        try:
            p = str(d).split(".")
            if len(p) == 3:
                return f"{p[2]}-{p[1]}-{p[0]}"
        except Exception:
            pass
        return str(d)

    s_cmp, e_cmp = to_cmp(start_date), to_cmp(end_date)
    filtered = [(n, m, v) for n, dt, m, v in rows if s_cmp <= to_cmp(dt) <= e_cmp]
    if not filtered:
        await msg.answer(f"📭 Нет данных за период {start_date} — {end_date}.")
        return

    user_data = {}
    for full_name, metric, value in filtered:
        d = user_data.setdefault(full_name, {"dk_z": 0, "dk_v": 0, "idk_z": 0, "idk_v": 0, "bio": 0})
        if metric == "dk_zayavki": d["dk_z"] += int(value)
        elif metric == "dk_vydacha": d["dk_v"] += int(value)
        elif metric == "idk_zayavki": d["idk_z"] += int(value)
        elif metric == "idk_vydacha": d["idk_v"] += int(value)
        elif metric == "bio": d["bio"] += int(value)

    period = start_date if start_date == end_date else f"{start_date} — {end_date}"
    text = f"📊 Сводка за {period}:\n\n"
    for full_name, m in sorted(user_data.items()):
        if m["dk_z"] or m["dk_v"] or m["idk_z"] or m["idk_v"] or m["bio"]:
            text += (f"👤 {full_name}:\n"
                     f"  🔹 ДК: {m['dk_z']}(заявки) / {m['dk_v']}(выдано)\n"
                     f"  🔹 ИДК: {m['idk_z']}(заявки) / {m['idk_v']}(выдано)\n"
                     f"  🔹 Био: {m['bio']}\n\n")
    gz = sum(d["dk_z"] for d in user_data.values())
    gv = sum(d["dk_v"] for d in user_data.values())
    iz = sum(d["idk_z"] for d in user_data.values())
    iv = sum(d["idk_v"] for d in user_data.values())
    bio = sum(d["bio"] for d in user_data.values())
    text += (f"📈 Итого по всем:\n"
             f"  🔹 ДК: {gz}(заявки) / {gv}(выдано)\n"
             f"  🔹 ИДК: {iz}(заявки) / {iv}(выдано)\n"
             f"  🔹 Био: {bio}")

    while len(text) > 4000:
        idx = text.rfind("\n\n", 0, 4000)
        if idx == -1:
            idx = 4000
        await msg.answer(text[:idx])
        text = text[idx:].lstrip()
    await msg.answer(text)

# ===== /summary С КНОПКАМИ =====
@router.message(Command("summary"))
async def summary_cmd(message: Message):
    if get_role(message.from_user.id) not in ("admin", "lead"):
        await message.answer("⛔ Только руководитель или админ.")
        return
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📅 Сегодня", callback_data="sum_today"),
         InlineKeyboardButton(text="📅 Вчера", callback_data="sum_yesterday")],
        [InlineKeyboardButton(text="📆 Неделя", callback_data="sum_week"),
         InlineKeyboardButton(text="📆 Месяц", callback_data="sum_month")],
        [InlineKeyboardButton(text="📅 Конкретная дата", callback_data="sum_single"),
         InlineKeyboardButton(text="🗓 Свой период", callback_data="sum_custom")]
    ])
    await message.answer("📊 Выбери период для сводки:", reply_markup=keyboard)

@router.callback_query(F.data == "sum_today")
@router.callback_query(F.data == "sum_yesterday")
@router.callback_query(F.data == "sum_week")
@router.callback_query(F.data == "sum_month")
async def summary_period_cb(callback: CallbackQuery):
    if get_role(callback.from_user.id) not in ("admin", "lead"):
        await callback.answer("⛔ Нет доступа.")
        return
    now = datetime.now(MSK)
    if callback.data == "sum_today":
        start = end = now.strftime("%d.%m.%Y")
    elif callback.data == "sum_yesterday":
        start = end = (now - timedelta(days=1)).strftime("%d.%m.%Y")
    elif callback.data == "sum_week":
        start = (now - timedelta(days=6)).strftime("%d.%m.%Y")
        end = now.strftime("%d.%m.%Y")
    else:
        start = now.replace(day=1).strftime("%d.%m.%Y")
        end = now.strftime("%d.%m.%Y")
    await callback.answer()
    await send_summary_text(callback, start, end)

@router.callback_query(F.data == "sum_single")
async def summary_single_cb(callback: CallbackQuery, state: FSMContext):
    if get_role(callback.from_user.id) not in ("admin", "lead"):
        await callback.answer("⛔ Нет доступа.")
        return
    await state.update_data(is_single=True)
    await callback.message.edit_text("📅 Введи дату (ДД.ММ.ГГГГ):")
    await callback.answer()
    await state.set_state(InputState.summary_start_date)

@router.callback_query(F.data == "sum_custom")
async def summary_custom_cb(callback: CallbackQuery, state: FSMContext):
    if get_role(callback.from_user.id) not in ("admin", "lead"):
        await callback.answer("⛔ Нет доступа.")
        return
    await state.update_data(is_single=False)
    await callback.message.edit_text("📅 Введи начальную дату (ДД.ММ.ГГГГ):")
    await callback.answer()
    await state.set_state(InputState.summary_start_date)

@router.message(InputState.summary_start_date)
async def summary_start_handler(message: Message, state: FSMContext):
    text = message.text.strip()
    try:
        datetime.strptime(text, "%d.%m.%Y")
    except ValueError:
        await message.answer("❌ Неверный формат. Пример: 01.11.2026. Попробуй ещё раз:")
        return
    data = await state.get_data()
    if data.get("is_single"):
        await state.clear()
        await send_summary_text(message, text, text)
    else:
        await state.update_data(start_date=text)
        await message.answer("📅 Теперь введи конечную дату (ДД.ММ.ГГГГ):")
        await state.set_state(InputState.summary_end_date)

@router.message(InputState.summary_end_date)
async def summary_end_handler(message: Message, state: FSMContext):
    text = message.text.strip()
    try:
        e = datetime.strptime(text, "%d.%m.%Y")
    except ValueError:
        await message.answer("❌ Неверный формат. Пример: 21.11.2026. Попробуй ещё раз:")
        return
    data = await state.get_data()
    start = data.get("start_date", text)
    s = datetime.strptime(start, "%d.%m.%Y")
    if s > e:
        s, e = e, s
    await state.clear()
    await send_summary_text(message, s.strftime("%d.%m.%Y"), e.strftime("%d.%m.%Y"))

# ===== СПИСОК ПОЛЬЗОВАТЕЛЕЙ И ДИНАМИКА =====
@router.message(Command("users"))
async def list_users(message: Message):
    if get_role(message.from_user.id) not in ("admin", "lead"):
        await message.answer("⛔ Только руководитель или админ.")
        return
    conn = sqlite3.connect(DB_FILE)
    users = conn.execute("SELECT tg_id, full_name FROM users ORDER BY full_name").fetchall()
    conn.close()
    if not users:
        await message.answer("📭 Пользователей нет.")
        return
    text = "👥 Выберите пользователя:\n\n"
    keyboard = InlineKeyboardMarkup(inline_keyboard=[])
    for tg_id, full_name in users:
        role_icon = "👨‍💻" if tg_id in ADMIN_IDS else ("👔" if tg_id in LEAD_IDS else "👤")
        text += f"{role_icon} {full_name} (ID: {tg_id})\n"
        keyboard.inline_keyboard.append([
            InlineKeyboardButton(text=f"📊 Динамика: {full_name}", callback_data=f"user_summary_{tg_id}"),
            InlineKeyboardButton(text=f"📎 Excel: {full_name}", callback_data=f"user_export_{tg_id}")
        ])
    await message.answer(text, reply_markup=keyboard)

@router.callback_query(F.data.startswith("user_summary_"))
async def user_summary(callback: CallbackQuery):
    if get_role(callback.from_user.id) not in ("admin", "lead"):
        await callback.answer("⛔ Нет доступа.")
        return
    target_id = int(callback.data.replace("user_summary_", ""))
    user = get_user(target_id)
    if not user:
        await callback.answer("❌ Не найден.")
        return
    now = datetime.now(MSK)
    start_date = now.replace(day=1).strftime("%d.%m.%Y")
    end_date = now.strftime("%d.%m.%Y")
    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("SELECT date, metric, value FROM data WHERE tg_id = ?", (target_id,)).fetchall()
    conn.close()
    rows = filter_data_by_date(rows, start_date, end_date)
    if not rows:
        await callback.message.answer(f"📭 У {user[1]} нет данных за текущий месяц.")
        await callback.answer()
        return
    daily_data, total_data = {}, {}
    for date, metric, value in rows:
        daily_data.setdefault(date, {})[metric] = int(value)
        total_data[metric] = total_data.get(metric, 0) + int(value)
    text = (f"📊 Динамика: {user[1]}\nПериод: {start_date} – {end_date}\n\n📈 ИТОГО:\n"
            f"  🔹 ДК: {total_data.get('dk_zayavki', 0)}(заявки) / {total_data.get('dk_vydacha', 0)}(выдано)\n"
            f"  🔹 ИДК: {total_data.get('idk_zayavki', 0)}(заявки) / {total_data.get('idk_vydacha', 0)}(выдано)\n"
            f"  🔹 Био: {total_data.get('bio', 0)}\n\n📅 ПО ДНЯМ (последние 10):\n")
    sorted_dates = sorted(daily_data.keys(), key=lambda x: datetime.strptime(x, "%d.%m.%Y"), reverse=True)[:10]
    for date in sorted_dates:
        dd = daily_data[date]
        text += f"  {date}: ДК {dd.get('dk_zayavki', 0)}/{dd.get('dk_vydacha', 0)} | ИДК {dd.get('idk_zayavki', 0)}/{dd.get('idk_vydacha', 0)} | Био {dd.get('bio', 0)}\n"
    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="📎 Выгрузить Excel", callback_data=f"user_export_{target_id}")]])
    await callback.message.answer(text, reply_markup=keyboard)
    await callback.answer()

@router.callback_query(F.data.startswith("user_export_"))
async def user_export(callback: CallbackQuery):
    if get_role(callback.from_user.id) not in ("admin", "lead"):
        await callback.answer("⛔ Нет доступа.")
        return
    target_id = int(callback.data.replace("user_export_", ""))
    try:
        file_path, name = generate_excel_for_user(target_id)
        if not file_path:
            await callback.message.answer("📭 Нет данных.")
            await callback.answer()
            return
        await callback.message.answer_document(FSInputFile(file_path), caption=f"📎 Отчёт по: {name}")
        os.remove(file_path)
        await callback.answer("✅ Выгружено")
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка: {e}")
        await callback.answer()

# ===== ПОДСКАЗКИ КОМАНД =====
USER_COMMANDS = [
    BotCommand(command="start", description="Регистрация в боте"),
    BotCommand(command="help", description="Справка по командам"),
]
LEAD_COMMANDS = [
    BotCommand(command="start", description="Регистрация в боте"),
    BotCommand(command="export", description="📎 Excel-отчёт (week / month / период)"),
    BotCommand(command="summary", description="📊 Сводка по команде (кнопки периодов)"),
    BotCommand(command="users", description="👥 Сотрудники: динамика и Excel"),
    BotCommand(command="help", description="Справка по командам"),
]
ADMIN_COMMANDS = LEAD_COMMANDS + [
    BotCommand(command="admin_records", description="🛠 Записи сотрудника (по ID)"),
    BotCommand(command="del_user", description="🗑 Удалить пользователя и его данные"),
]

async def setup_commands():
    await bot.set_my_commands(USER_COMMANDS, scope=BotCommandScopeDefault())
    for admin_id in ADMIN_IDS:
        await bot.set_my_commands(ADMIN_COMMANDS, scope=BotCommandScopeChat(chat_id=admin_id))
    for lead_id in LEAD_IDS:
        await bot.set_my_commands(LEAD_COMMANDS, scope=BotCommandScopeChat(chat_id=lead_id))

@router.message(Command("help"))
async def help_cmd(message: Message):
    role = get_role(message.from_user.id)
    if role == "admin":
        text = ("👨‍💻 Команды админа:\n\n"
                "/start — регистрация\n"
                "/export — Excel по всем (week / month / Ноябрь-2026 / 01.11.2026 21.11.2026)\n"
                "/export user [ID] — отчёт одного сотрудника\n"
                "/summary — сводка по команде (кнопки: сегодня, вчера, неделя, месяц, дата, период)\n"
                "/users — сотрудники + кнопки Динамика и Excel\n"
                "/admin_records [ID] — записи сотрудника, правка и удаление\n"
                "/del_user [ID] — удалить пользователя и его данные\n"
                "/help — эта справка")
    elif role == "lead":
        text = ("👔 Команды руководителя:\n\n"
                "/start — регистрация\n"
                "/export — Excel по всем (week / month / Ноябрь-2026 / 01.11.2026 21.11.2026)\n"
                "/export user [ID] — отчёт одного сотрудника\n"
                "/summary — сводка по команде (кнопки: сегодня, вчера, неделя, месяц, дата, период)\n"
                "/users — сотрудники + кнопки Динамика и Excel\n"
                "/help — эта справка")
    else:
        text = ("👤 Твои команды:\n\n"
                "/start — регистрация\n"
                "📝 Внести данные — кнопка в меню\n"
                "📊 Мои данные — сводка за неделю\n"
                "✏️ Мои записи — правка и удаление своих записей\n"
                "/help — эта справка")
    await message.answer(text)

async def main():
    await setup_commands()
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
