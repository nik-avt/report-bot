import os
import sqlite3
import logging
from datetime import datetime, timedelta, timezone
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message, ReplyKeyboardMarkup, KeyboardButton,ReplyKeyboardRemove
    InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile, CallbackQuery
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter  # ВАЖНО для сложной шапки
from dotenv import load_dotenv

# ===== ЛОГИРОВАНИЕ И НАСТРОЙКИ =====
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x]
LEAD_IDS = [int(x) for x in os.getenv("LEAD_IDS", "").split(",") if x]
DB_FILE = "bot.db"
EXCEL_FILE = "report.xlsx"
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
    conn.commit(); conn.close()

init_db()

def get_user(tg_id):
    conn = sqlite3.connect(DB_FILE)
    user = conn.execute("SELECT * FROM users WHERE tg_id = ?", (tg_id,)).fetchone()
    conn.close(); return user

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
    waiting_name = State(); waiting_date = State(); waiting_value = State()
    final_confirm = State(); editing = State()

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
    if user: await message.answer(f"Привет, {user[1]}! 👋", reply_markup=main_menu())
    else:
        await message.answer("Привет! Введи свои ФИО для регистрации:")
        await state.set_state(InputState.waiting_name)

@router.message(InputState.waiting_name)
async def save_name(message: Message, state: FSMContext):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("INSERT INTO users (tg_id, full_name) VALUES (?, ?)", (message.from_user.id, message.text.strip()))
    conn.commit(); conn.close()
    await state.clear()
    await message.answer(f"✅ Зарегистрирован: {message.text}", reply_markup=main_menu())

@router.message(F.text == "📝 Внести данные")
async def start_input(message: Message, state: FSMContext):
    if not get_user(message.from_user.id):
        await message.answer("Сначала зарегистрируйся: /start"); return
    await message.answer("📅 Введи дату (ДД.ММ.ГГГГ) или напиши 'сегодня' / 'вчера':",
                         reply_markup=ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="сегодня")], [KeyboardButton(text="вчера")]], resize_keyboard=True))
    await state.set_state(InputState.waiting_date)

@router.message(InputState.waiting_date)
async def get_date(message: Message, state: FSMContext):
    text = message.text.strip().lower(); now = datetime.now(MSK)
    if text == "сегодня": date_str = now.strftime("%d.%m.%Y")
    elif text == "вчера": date_str = (now - timedelta(days=1)).strftime("%d.%m.%Y")
    else:
        try:
            parsed = datetime.strptime(text, "%d.%m.%Y")
            if parsed.replace(tzinfo=MSK) > now.replace(hour=23, minute=59, second=59):
                await message.answer("❌ Дата не может быть в будущем."); return
            date_str = parsed.strftime("%d.%m.%Y")
        except ValueError:
            await message.answer("❌ Формат ДД.ММ.ГГГГ:"); return

    # Проверка существующих данных для замены
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
        if value < 0: await message.answer("❌ Значение не может быть отрицательным."); return
    except ValueError:
        await message.answer("❌ Введи число:"); return
    
    data = await state.get_data()
    metric_index = data["metric_index"]
    metric_label, metric_key = METRICS_ORDER[metric_index]
    values = data.get("values", {})
    
    if metric_key == "dk_vydacha" and value < values.get("dk_zayavki", 0):
        await message.answer(f"⚠️ ДК выдача ({int(value)}) < ДК заявок ({int(values.get('dk_zayavki',0))}). Введи ≥ {int(values.get('dk_zayavki',0))}:"); return
    if metric_key == "idk_vydacha" and value < values.get("idk_zayavki", 0):
        await message.answer(f"⚠️ ИДК выдача ({int(value)}) < ИДК заявок ({int(values.get('idk_zayavki',0))}). Введи ≥ {int(values.get('idk_zayavki',0))}:"); return
    
    values[metric_key] = value; await state.update_data(values=values)
    metric_index += 1
    if metric_index < len(METRICS_ORDER):
        await state.update_data(metric_index=metric_index)
        metric_label, _ = METRICS_ORDER[metric_index]
        step_names = ["2️⃣", "3️⃣", "4️⃣", "5️⃣"]
        await message.answer(f"{step_names[metric_index - 1]} Введи: {metric_label}")
    else: await show_final_confirm(message, state, values)

async def show_final_confirm(message: Message, state: FSMContext, values: dict):
    data = await state.get_data(); date_str = data["date"]; is_replace = data.get("is_replace", False)
    question = "⚠️ Старые данные за эту дату будут заменены. Продолжить?" if is_replace else "Сохранить?"
    save_label = "💾 Заменить" if is_replace else "💾 Сохранить"
    text = (f"📋 Данные за {date_str}:\n\n"
            f"🔹 ДК: {int(values.get('dk_zayavki', 0))}(заявки) / {int(values.get('dk_vydacha', 0))}(выдано)\n"
            f"🔹 ИДК: {int(values.get('idk_zayavki', 0))}(заявки) / {int(values.get('idk_vydacha', 0))}(выдано)\n"
            f"🔹 Био: {int(values.get('bio', 0))}\n\n{question}")
    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=save_label, callback_data="save_all"), InlineKeyboardButton(text="❌ Отменить", callback_data="cancel_all")]])
    await message.answer(text, reply_markup=keyboard); await state.set_state(InputState.final_confirm)

@router.callback_query(F.data == "save_all", InputState.final_confirm)
async def save_all(callback: CallbackQuery, state: FSMContext):
    await callback.answer(); data = await state.get_data()
    date_str = data["date"]; values = data.get("values", {}); is_replace = data.get("is_replace", False)
    conn = sqlite3.connect(DB_FILE)
    for metric_key, value in values.items():
        cur = conn.execute("UPDATE data SET value = ? WHERE tg_id = ? AND date = ? AND metric = ?", (value, callback.from_user.id, date_str, metric_key))
        if cur.rowcount == 0: conn.execute("INSERT INTO data (tg_id, date, metric, value) VALUES (?, ?, ?, ?)", (callback.from_user.id, date_str, metric_key, value))
    conn.commit(); conn.close(); await state.clear()
    msg = f"✅ Данные за {date_str} заменены!" if is_replace else f"✅ Сохранено за {date_str}!"
    await callback.message.answer(msg, reply_markup=main_menu())

@router.callback_query(F.data == "cancel_all", InputState.final_confirm)
async def cancel_all(callback: CallbackQuery, state: FSMContext):
    await callback.answer("❌ Отменено"); await state.clear()
    await callback.message.answer("Ввод отменён.", reply_markup=main_menu())

# ===== ПРОСМОТР И РЕДАКТИРОВАНИЕ =====
@router.message(F.text == "📊 Мои данные")
async def my_data(message: Message):
    user = get_user(message.from_user.id)
    if not user: await message.answer("Сначала зарегистрируйся: /start"); return
    conn = sqlite3.connect(DB_FILE)
    week_ago = (datetime.now(MSK) - timedelta(days=7)).strftime("%d.%m.%Y")
    today = datetime.now(MSK).strftime("%d.%m.%Y")
    rows = conn.execute("SELECT metric, SUM(value) FROM data WHERE tg_id = ? AND date BETWEEN ? AND ? GROUP BY metric", (message.from_user.id, week_ago, today)).fetchall()
    conn.close(); d = {m: int(v) for m, v in rows}
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
    if not rows: await message.answer("📭 Записей нет."); return
    text = "📝 Твои последние 5 записей:\n\n"; keyboard = InlineKeyboardMarkup(inline_keyboard=[])
    for row in rows:
        rec_id, date, metric, value = row; metric_name = METRIC_NAMES.get(metric, metric)
        v = int(value) if value == int(value) else value; text += f"• {date} | {metric_name}: {v}\n"
        keyboard.inline_keyboard.append([InlineKeyboardButton(text=f"✏️ {date} {metric_name}", callback_data=f"edit_{rec_id}"), InlineKeyboardButton(text="🗑️", callback_data=f"delete_{rec_id}")])
    await message.answer(text, reply_markup=keyboard)

@router.callback_query(F.data.startswith("edit_"))
async def start_edit(callback: CallbackQuery, state: FSMContext):
    rec_id = int(callback.data.replace("edit_", ""))
    await state.update_data(edit_id=rec_id, is_admin_edit=False)
    await callback.message.answer("Введи новое значение:"); await state.set_state(InputState.editing); await callback.answer()

@router.message(InputState.editing)
async def save_edit(message: Message, state: FSMContext):
    try: value = float(message.text.replace(",", "."))
    except ValueError: await message.answer("❌ Введи число:"); return
    data = await state.get_data(); rec_id = data.get("edit_id"); is_admin = data.get("is_admin_edit", False)
    conn = sqlite3.connect(DB_FILE)
    if is_admin: cursor = conn.execute("UPDATE data SET value = ? WHERE id = ?", (value, rec_id))
    else: cursor = conn.execute("UPDATE data SET value = ? WHERE id = ? AND tg_id = ?", (value, rec_id, message.from_user.id))
    conn.commit(); conn.close(); await state.clear()
    v = int(value) if value == int(value) else value
    await message.answer(f"✅ Обновлено: {v}", reply_markup=main_menu() if not is_admin else None)

@router.callback_query(F.data.startswith("delete_"))
async def delete_record(callback: CallbackQuery):
    rec_id = int(callback.data.replace("delete_", ""))
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.execute("DELETE FROM data WHERE id = ? AND tg_id = ?", (rec_id, callback.from_user.id))
    conn.commit(); conn.close(); await callback.answer("🗑️ Удалено")
    await callback.message.edit_text("✅ Запись удалена." if cursor.rowcount > 0 else "⚠️ Не найдена.")

# ===== АДМИН ФУНКЦИИ =====
@router.message(Command("admin_records"))
async def admin_records(message: Message):
    if get_role(message.from_user.id) != "admin": await message.answer("⛔ Только админ."); return
    args = message.text.split(maxsplit=1)
    if len(args) < 2: await message.answer("Использование: /admin_records [Telegram ID]\nПример: /admin_records 123456789"); return
    target_id = int(args[1]); user = get_user(target_id)
    if not user: await message.answer("❌ Пользователь не найден."); return
    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("SELECT id, date, metric, value FROM data WHERE tg_id = ? ORDER BY created_at DESC LIMIT 10", (target_id,)).fetchall()
    conn.close()
    if not rows: await message.answer(f"📭 У {user[1]} нет записей."); return
    text = f"📝 Последние записи {user[1]}:\n\n"; keyboard = InlineKeyboardMarkup(inline_keyboard=[])
    for row in rows:
        rec_id, date, metric, value = row; metric_name = METRIC_NAMES.get(metric, metric)
        v = int(value) if value == int(value) else value; text += f"• {date} | {metric_name}: {v}\n"
        keyboard.inline_keyboard.append([InlineKeyboardButton(text=f"✏️ {date} {metric_name}", callback_data=f"admin_edit_{rec_id}"), InlineKeyboardButton(text="🗑️", callback_data=f"admin_delete_{rec_id}")])
    await message.answer(text, reply_markup=keyboard)

@router.callback_query(F.data.startswith("admin_edit_"))
async def admin_start_edit(callback: CallbackQuery, state: FSMContext):
    if get_role(callback.from_user.id) != "admin": await callback.answer("⛔ Только админ."); return
    rec_id = int(callback.data.replace("admin_edit_", ""))
    await state.update_data(edit_id=rec_id, is_admin_edit=True)
    await callback.message.answer("Введи новое значение:"); await state.set_state(InputState.editing); await callback.answer()

@router.callback_query(F.data.startswith("admin_delete_"))
async def admin_delete_record(callback: CallbackQuery):
    if get_role(callback.from_user.id) != "admin": await callback.answer("⛔ Только админ."); return
    rec_id = int(callback.data.replace("admin_delete_", ""))
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.execute("DELETE FROM data WHERE id = ?", (rec_id,))
    conn.commit(); conn.close(); await callback.answer("🗑️ Удалено")
    await callback.message.edit_text("✅ Запись удалена." if cursor.rowcount > 0 else "⚠️ Не найдена.")

# ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ДАТ =====
def parse_date_range(args_str):
    now = datetime.now(MSK)
    if not args_str or args_str.strip() == '': return None, None
    args_str = args_str.strip().lower()
    if args_str == 'week': return (now - timedelta(days=6)).strftime("%d.%m.%Y"), now.strftime("%d.%m.%Y")
    if args_str == 'month': return now.replace(day=1).strftime("%d.%m.%Y"), now.strftime("%d.%m.%Y")
    if '-' in args_str and '.' not in args_str:
        try: dt = datetime.strptime(args_str, "%B-%Y")
        except ValueError:
            ru_months = {'январь': 'January', 'февраль': 'February', 'март': 'March', 'апрель': 'April', 'май': 'May', 'июнь': 'June', 'июль': 'July', 'август': 'August', 'сентябрь': 'September', 'октябрь': 'October', 'ноябрь': 'November', 'декабрь': 'December'}
            try:
                month_ru, year = args_str.split('-', 1); dt = datetime.strptime(f"{ru_months[month_ru]}-{year}", "%B-%Y")
            except: return None, None
        start = dt.replace(day=1).strftime("%d.%m.%Y")
        last_day = (dt.replace(month=dt.month % 12 + 1, day=1, year=dt.year + (1 if dt.month == 12 else 0)) - timedelta(days=1))
        return start, last_day.strftime("%d.%m.%Y")
    parts = args_str.split()
    if len(parts) == 2:
        try:
            s, e = datetime.strptime(parts[0], "%d.%m.%Y"), datetime.strptime(parts[1], "%d.%m.%Y")
            if s > e: s, e = e, s; return s.strftime("%d.%m.%Y"), e.strftime("%d.%m.%Y")
        except ValueError: pass
    try: dt = datetime.strptime(args_str, "%d.%m.%Y"); return dt.strftime("%d.%m.%Y"), dt.strftime("%d.%m.%Y")
    except ValueError: return None, None

def filter_data_by_date(all_data, start_date, end_date):
    if not start_date or not end_date: return all_data
    def to_cmp(d): p = d.split('.'); return f"{p[2]}-{p[1]}-{p[0]}"
    s, e = to_cmp(start_date), to_cmp(end_date)
    return [row for row in all_data if s <= to_cmp(row[1]) <= e]

# ===== ГЕНЕРАЦИЯ EXCEL (ФИНАЛЬНАЯ ВЕРСИЯ) =====
def generate_excel(start_date=None, end_date=None):
    conn = sqlite3.connect(DB_FILE)
    users = conn.execute("SELECT tg_id, full_name FROM users ORDER BY full_name").fetchall()
    all_data = conn.execute("SELECT tg_id, date, metric, value FROM data").fetchall()
    conn.close()
    
    clean_data = []; valid_dates_set = set()
    for tg_id, date_str, metric, value in all_data:
        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
            if dt.year < 2020 or dt > datetime.now() + timedelta(days=1): continue
            if start_date and end_date:
                s = datetime.strptime(start_date, "%d.%m.%Y"); e = datetime.strptime(end_date, "%d.%m.%Y")
                if not (s <= dt <= e): continue
            clean_data.append((tg_id, date_str, metric, value)); valid_dates_set.add(date_str)
        except ValueError: continue
    if not clean_data: return None
    
    wb = Workbook(); wb.remove(wb.active)
    sorted_dates = sorted(list(valid_dates_set), key=lambda x: datetime.strptime(x, "%d.%m.%Y"))
    ws = wb.create_sheet(title="Отчет")
    
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    subheader_fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    total_font = Font(bold=True, size=11)
    user_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    current_row = 1
    col_idx = 2 # Инициализация для ширины столбцов
    
    for tg_id, full_name in users:
        user_has_data = any(d[0] == tg_id for d in clean_data)
        if not user_has_data: continue
        
        last_col_idx = len(sorted_dates) * 3 + 1; last_col_letter = get_column_letter(last_col_idx)
        ws.merge_cells(f'A{current_row}:{last_col_letter}{current_row}')
        name_cell = ws[f'A{current_row}']; name_cell.value = full_name
        name_cell.font = Font(bold=True, size=12); name_cell.fill = user_fill
        name_cell.border = thin_border; name_cell.alignment = Alignment(horizontal="left", vertical="center")
        current_row += 1
        
        ws.cell(row=current_row, column=1).value = "Показатель"
        ws.cell(row=current_row, column=1).font = header_font; ws.cell(row=current_row, column=1).fill = header_fill
        ws.cell(row=current_row, column=1).border = thin_border; ws.cell(row=current_row+1, column=1).border = thin_border 
        
        col_idx = 2
        for date_str in sorted_dates:
            start_l = get_column_letter(col_idx); end_l = get_column_letter(col_idx + 2)
            ws.merge_cells(f'{start_l}{current_row}:{end_l}{current_row}')
            d_cell = ws[f'{start_l}{current_row}']; d_cell.value = date_str
            d_cell.font = header_font; d_cell.fill = header_fill
            d_cell.alignment = Alignment(horizontal="center"); d_cell.border = thin_border
            
            subs = ["ДК", "ИДК", "Био"]
            for i, s in enumerate(subs):
                c = ws.cell(row=current_row+1, column=col_idx+i); c.value = s
                c.font = Font(bold=True, size=10, color="FFFFFF"); c.fill = subheader_fill
                c.alignment = Alignment(horizontal="center"); c.border = thin_border
            col_idx += 3
        current_row += 2
        
        u_data = {}
        for tid, d, m, v in clean_data:
            if tid == tg_id: u_data.setdefault(d, {})[m] = u_data.get(d, {}).get(m, 0) + v
        
        metrics_map = [("ДК (заяв/выд)", 'dk_zayavki', 'dk_vydacha', 0), ("ИДК (заяв/выд)", 'idk_zayavki', 'idk_vydacha', 1), ("Био", 'bio', None, 2)]
        rows_offset = 0
        for label, m1, m2, col_offset in metrics_map:
            r = current_row + rows_offset
            c = ws.cell(row=r, column=1); c.value = label; c.fill = user_fill; c.border = thin_border
            
            date_col_start = 2
            for date_str in sorted_dates:
                dd = u_data.get(date_str, {}); target_col = date_col_start + col_offset
                if m2:
                    v1 = int(dd.get(m1, 0)); v2 = int(dd.get(m2, 0))
                    val = f"{v1}/{v2}" if (v1 or v2) else "-"
                else:
                    v1 = int(dd.get(m1, 0)); val = str(v1) if v1 else "-"
                
                cell = ws.cell(row=r, column=target_col); cell.value = val
                cell.alignment = Alignment(horizontal="center"); cell.border = thin_border
                date_col_start += 3
            rows_offset += 1
        current_row += 3; ws.row_dimensions[current_row].height = 4; current_row += 1

    ws.column_dimensions['A'].width = 22
    for i in range(2, col_idx): ws.column_dimensions[get_column_letter(i)].width = 8
    wb.save(EXCEL_FILE); return EXCEL_FILE

def generate_excel_for_user(tg_id, start_date=None, end_date=None):
    conn = sqlite3.connect(DB_FILE)
    user = conn.execute("SELECT full_name FROM users WHERE tg_id = ?", (tg_id,)).fetchone()
    all_data = conn.execute("SELECT date, metric, value FROM data WHERE tg_id = ?", (tg_id,)).fetchall()
    conn.close()
    if not user or not all_data: return None, None
    all_data = filter_data_by_date(all_data, start_date, end_date)
    if not all_data: return None, None
    
    wb = Workbook(); ws = wb.active; ws.title = f"{user[0]}"[:31]
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    dates = sorted(set(row[0] for row in all_data), key=lambda x: datetime.strptime(x, "%d.%m.%Y"))
    ws.append(["Показатель"] + [datetime.strptime(d, "%d.%m.%Y").day for d in dates])
    for cell in ws[1]: cell.font = header_font; cell.fill = header_fill; cell.alignment = Alignment(horizontal="center"); cell.border = thin_border
    
    metrics_map = {}; 
    for date, metric, value in all_data: metrics_map.setdefault(metric, {})[date] = metrics_map.get(metric, {}).get(date, 0) + value
    
    row_idx = 2
    for label, key in METRICS_ORDER:
        row = [label] + [int(metrics_map.get(key, {}).get(d, 0)) for d in dates]
        ws.append(row)
        for cell in ws[row_idx]: cell.border = thin_border; cell.alignment = Alignment(horizontal="center")
        row_idx += 1
    ws.column_dimensions['A'].width = 18
    for i in range(2, 2 + len(dates)):
        col_letter = chr(64 + i) if i <= 26 else chr(64 + (i-1)//26) + chr(65 + (i-1)%26)
        ws.column_dimensions[col_letter].width = 6
    
    filename = f"report_{user[0].replace(' ', '_')}_{datetime.now().strftime('%d%m%Y')}.xlsx"
    wb.save(filename); return filename, user[0]

# ===== КОМАНДЫ ЭКСПОРТА И СВОДКИ =====
@router.message(Command("export"))
async def export_excel(message: Message):
    role = get_role(message.from_user.id)
    if role not in ("admin", "lead"): await message.answer("⛔ Только руководитель или админ."); return
    args = message.text.split()
    if len(args) >= 3 and args[1].lower() == 'user':
        try: target_id = int(args[2])
        except ValueError: await message.answer("❌ Неверный ID. Пример: /export user 123456789 week"); return
        period_args = ' '.join(args[3:]) if len(args) > 3 else ''
        start_date, end_date = parse_date_range(period_args)
        try:
            file_path, name = generate_excel_for_user(target_id, start_date, end_date)
            if not file_path: await message.answer("📭 Нет данных."); return
            caption = f"📎 Отчёт по {name}"
            if start_date: caption += f": {start_date}{' – ' + end_date if start_date != end_date else ''}"
            await message.answer_document(FSInputFile(file_path), caption=caption); os.remove(file_path)
        except Exception as e: await message.answer(f"❌ Ошибка: {e}")
        return
    
    args_str = ' '.join(args[1:]) if len(args) > 1 else ''
    start_date, end_date = parse_date_range(args_str)
    caption = "📎 Отчёт: все данные" if not start_date else (f"📎 Отчёт за {start_date}" if start_date == end_date else f"📎 Отчёт с {start_date} по {end_date}")
    try:
        file_path = generate_excel(start_date, end_date)
        if not file_path: await message.answer("📭 Нет данных за выбранный период."); return
        await message.answer_document(FSInputFile(file_path), caption=caption)
    except Exception as e: await message.answer(f"❌ Ошибка: {e}")

def split_message(text, max_len=4000):
    parts = []
    while len(text) > max_len:
        idx = text.rfind('\n\n', 0, max_len)
        if idx == -1: idx = max_len
        parts.append(text[:idx]); text = text[idx:].lstrip()
    if text: parts.append(text)
    return parts

@router.message(Command("summary"))
async def team_summary(message: Message):
    role = get_role(message.from_user.id)
    if role not in ("admin", "lead"): await message.answer("⛔ Только руководитель или админ."); return
    args = message.text.split(maxsplit=1); args_str = args[1].strip() if len(args) > 1 else 'week'
    start_date, end_date = parse_date_range(args_str)
    conn = sqlite3.connect(DB_FILE)
    all_rows = conn.execute("SELECT u.full_name, d.metric, d.value FROM data d JOIN users u ON d.tg_id = u.tg_id").fetchall()
    conn.close()
    filtered = filter_data_by_date(all_rows, start_date, end_date)
    if not filtered: await message.answer("📭 Данных нет за выбранный период."); return
    
    user_data = {}
    for full_name, metric, value in filtered: user_data.setdefault(full_name, {})[metric] = user_data[full_name].get(metric, 0) + int(value)
    period_text = "за всё время" if not start_date else (f"за {start_date}" if start_date == end_date else f"с {start_date} по {end_date}")
    text = f"📊 Сводка по команде {period_text}:\n\n"
    for full_name, m in sorted(user_data.items()):
        text += (f"👤 {full_name}:\n  🔹 ДК: {m.get('dk_zayavki', 0)}(заявки) / {m.get('dk_vydacha', 0)}(выдано)\n  🔹 ИДК: {m.get('idk_zayavki', 0)}(заявки) / {m.get('idk_vydacha', 0)}(выдано)\n  🔹 Био: {m.get('bio', 0)}\n\n")
    for part in split_message(text): await message.answer(part)

# ===== СПИСОК ПОЛЬЗОВАТЕЛЕЙ И ДИНАМИКА =====
@router.message(Command("users"))
async def list_users(message: Message):
    role = get_role(message.from_user.id)
    if role not in ("admin", "lead"): await message.answer("⛔ Только руководитель или админ."); return
    conn = sqlite3.connect(DB_FILE)
    users = conn.execute("SELECT tg_id, full_name FROM users ORDER BY full_name").fetchall()
    conn.close()
    if not users: await message.answer("📭 Пользователей нет."); return
    text = "👥 Выберите пользователя:\n\n"; keyboard = InlineKeyboardMarkup(inline_keyboard=[])
    for tg_id, full_name in users:
        role_icon = "👨‍💻" if tg_id in ADMIN_IDS else ("👔" if tg_id in LEAD_IDS else "👤")
        text += f"{role_icon} {full_name} (ID: {tg_id})\n"
        keyboard.inline_keyboard.append([InlineKeyboardButton(text=f"📊 Динамика", callback_data=f"user_summary_{tg_id}"), InlineKeyboardButton(text=f"📎 Excel", callback_data=f"user_export_{tg_id}")])
    await message.answer(text, reply_markup=keyboard)

@router.callback_query(F.data.startswith("user_summary_"))
async def user_summary(callback: CallbackQuery):
    if get_role(callback.from_user.id) not in ("admin", "lead"): await callback.answer("⛔ Нет доступа."); return
    target_id = int(callback.data.replace("user_summary_", "")); user = get_user(target_id)
    if not user: await callback.answer("❌ Не найден."); return
    now = datetime.now(MSK); start_date, end_date = now.replace(day=1).strftime("%d.%m.%Y"), now.strftime("%d.%m.%Y")
    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("SELECT metric, SUM(value), date FROM data WHERE tg_id = ? AND date BETWEEN ? AND ? GROUP BY metric, date ORDER BY date", (target_id, start_date, end_date)).fetchall()
    conn.close()
    if not rows: await callback.message.answer(f"📭 У {user[1]} нет данных за текущий месяц."); await callback.answer(); return
    
    daily_data, total_data = {}, {}
    for metric, value, date in rows: daily_data.setdefault(date, {})[metric] = int(value); total_data[metric] = total_data.get(metric, 0) + int(value)
    text = f"📊 Динамика: {user[1]}\nПериод: {start_date} – {end_date}\n\n📈 ИТОГО:\n  🔹 ДК: {total_data.get('dk_zayavki', 0)}(заявки) / {total_data.get('dk_vydacha', 0)}(выдано)\n  🔹 ИДК: {total_data.get('idk_zayavki', 0)}(заявки) / {total_data.get('idk_vydacha', 0)}(выдано)\n  🔹 Био: {total_data.get('bio', 0)}\n\n📅 ПО ДНЯМ (последние 10):\n"
    sorted_dates = sorted(daily_data.keys(), key=lambda x: datetime.strptime(x, "%d.%m.%Y"), reverse=True)[:10]
    for date in sorted_dates:
        dd = daily_data[date]; text += f"  {date}: ДК {dd.get('dk_zayavki', 0)}/{dd.get('dk_vydacha', 0)} | ИДК {dd.get('idk_zayavki', 0)}/{dd.get('idk_vydacha', 0)} | Био {dd.get('bio', 0)}\n"
    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="📎 Выгрузить Excel", callback_data=f"user_export_{target_id}")]])
    await callback.message.answer(text, reply_markup=keyboard); await callback.answer()

@router.callback_query(F.data.startswith("user_export_"))
async def user_export(callback: CallbackQuery):
    if get_role(callback.from_user.id) not in ("admin", "lead"): await callback.answer("⛔ Нет доступа."); return
    target_id = int(callback.data.replace("user_export_", ""))
    try:
        file_path, name = generate_excel_for_user(target_id)
        if not file_path: await callback.message.answer(f"📭 Нет данных у {name}."); await callback.answer(); return
        await callback.message.answer_document(FSInputFile(file_path), caption=f"📎 Отчёт по: {name}")
        os.remove(file_path); await callback.answer("✅ Выгружено")
    except Exception as e: await callback.message.answer(f"❌ Ошибка: {e}"); await callback.answer()

if __name__ == "__main__":
    import asyncio; asyncio.run(dp.start_polling(bot))
