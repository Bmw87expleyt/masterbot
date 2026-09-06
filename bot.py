import asyncio
import json
import os
import sqlite3
from datetime import datetime
from typing import Dict, List, Optional

import openpyxl
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    FSInputFile,
    BotCommand
)
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage

# ======================
# CONFIG & INIT
# ======================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICES_FILE = os.path.join(BASE_DIR, "prices.json")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
DB_FILE = os.path.join(BASE_DIR, "users.db")

API_TOKEN = os.getenv("BOT_TOKEN", "8704247957:AAH53nb6eBhUsnbB3YH-_AsWL4_5u8ah66E").strip()
ADMIN_IDS = {292170708}

bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# ======================
# DATABASE (SQLite)
# ======================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            full_name TEXT,
            phone TEXT,
            is_blocked INTEGER DEFAULT 0,
            joined_at TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def get_user(user_id: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM users WHERE user_id = ?', (user_id,))
    user = cursor.fetchone()
    conn.close()
    return user

def add_user(user_id: int, username: str, full_name: str, phone: str):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT OR REPLACE INTO users (user_id, username, full_name, phone, is_blocked, joined_at)
        VALUES (?, ?, ?, ?, COALESCE((SELECT is_blocked FROM users WHERE user_id = ?), 0), ?)
    ''', (user_id, username, full_name, phone, user_id, datetime.now()))
    conn.commit()
    conn.close()

def update_block_status(user_id: int, status: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('UPDATE users SET is_blocked = ? WHERE user_id = ?', (status, user_id))
    conn.commit()
    conn.close()

def get_all_users():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT user_id, full_name, phone, is_blocked, joined_at FROM users')
    users = cursor.fetchall()
    conn.close()
    return users

# ======================
# DATA LOADERS
# ======================
def load_json(path: str, default):
    if not os.path.exists(path): return default
    with open(path, "r", encoding="utf-8") as f: return json.load(f)

def save_json(path: str, data):
    with open(path, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)

def settings():
    return load_json(SETTINGS_FILE, {
        "contact_text": "🛠 Специлизированная лаборатория по компонентному ремонту и пайке плат.\nПринимаем устройства со всей РФ.",
        "admin_chat_id": None
    })

def prices():
    data = load_json(PRICES_FILE, {"brands": {}})
    if "brands" not in data: data["brands"] = {}
    return data

def add_price_item(brand: str, model: str, fault: str, price: str):
    data = prices()
    if brand not in data["brands"]:
        data["brands"][brand] = {}
    if model not in data["brands"][brand]:
        data["brands"][brand][model] = {}
    data["brands"][brand][model][fault] = price
    save_json(PRICES_FILE, data)

# ======================
# EXCEL LOGIC
# ======================
def generate_excel_price(data: dict, filepath: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Прайс Пайка"
    ws.append(["Бренд", "Модель", "Услуга / Поломка", "Цена (руб)"])
    
    for brand, models in data.get("brands", {}).items():
        for model, faults in models.items():
            for fault, price in faults.items():
                ws.append([brand, model, fault, str(price)])
                
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2
    wb.save(filepath)

def parse_excel_to_json(filepath: str) -> dict:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    new_data = {"brands": {}}
        
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0] or not row[1]: continue
        brand, model, fault, price = [str(x).strip() if x is not None else "" for x in row]
        
        if brand not in new_data["brands"]:
            new_data["brands"][brand] = {}
        if model not in new_data["brands"][brand]:
            new_data["brands"][brand][model] = {}
        if fault:
            new_data["brands"][brand][model][fault] = price
                
    return new_data

# ======================
# KEYBOARDS
# ======================
main_kb = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔍 Найти модель / Прайс")],
        [KeyboardButton(text="📍 Контакты и доставка"), KeyboardButton(text="💬 Задать вопрос мастеру")]
    ],
    resize_keyboard=True,
)

cancel_kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="Отмена")]], resize_keyboard=True)
phone_request_kb = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="📞 Отправить телефон", request_contact=True)]],
    resize_keyboard=True
)

admin_kb = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📥 Скачать прайс (Excel)"), KeyboardButton(text="➕ Добавить услугу (Текстом)")],
        [KeyboardButton(text="👥 База клиентов"), KeyboardButton(text="🚫 Блокировка")],
        [KeyboardButton(text="⬅️ Выйти из админки")]
    ],
    resize_keyboard=True,
)

# ======================
# FSM STATES
# ======================
class RegFlow(StatesGroup):
    name = State()
    phone = State()

class SearchFlow(StatesGroup):
    waiting_query = State()

class QuestionFlow(StatesGroup):
    waiting_text = State()

class AdminFlow(StatesGroup):
    waiting_add_item = State()
    waiting_block_id = State()

# ======================
# MIDDLEWARE / AUTH & REGISTRATION
# ======================
async def check_access(message: types.Message, state: FSMContext) -> bool:
    user = get_user(message.from_user.id)
    if user:
        if user[4] == 1:
            await message.answer("⛔️ Ваш доступ к боту ограничен.")
            return False
        return True
    else:
        await state.set_state(RegFlow.name)
        await message.answer("Здравствуйте! 👋\nДля доступа к боту и связи с мастером, пожалуйста, введите ваше **Имя**:", parse_mode="Markdown")
        return False

@dp.message(RegFlow.name)
async def reg_name(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Пожалуйста, введите ваше имя текстом.")
    await state.update_data(full_name=message.text)
    await state.set_state(RegFlow.phone)
    await message.answer(f"Приятно познакомиться, {message.text}!\nТеперь нажмите кнопку ниже, чтобы поделиться номером телефона.", reply_markup=phone_request_kb)

@dp.message(RegFlow.phone)
async def reg_phone(message: types.Message, state: FSMContext):
    if not message.contact:
        return await message.answer("Пожалуйста, используйте кнопку «📞 Отправить телефон» внизу экрана.")
    
    data = await state.get_data()
    u = message.from_user
    add_user(u.id, u.username, data['full_name'], message.contact.phone_number)
    
    await state.clear()
    await message.answer("✅ Регистрация успешно завершена! Добро пожаловать.", reply_markup=main_kb)
    
    st = settings()
    admin_chat = st.get("admin_chat_id")
    for admin in (admin_chat and [admin_chat] or ADMIN_IDS):
        try:
            await bot.send_message(
                admin, 
                f"👤 <b>Новый пользователь в боте:</b>\nИмя: {data['full_name']}\nТелефон: {message.contact.phone_number}\nID: `{u.id}`", 
                parse_mode="HTML"
            )
        except: pass

# ======================
# CLIENT HANDLERS
# ======================
@dp.message(CommandStart())
async def start(message: types.Message, state: FSMContext):
    await state.clear()
    if await check_access(message, state):
        await message.answer("Выберите нужный раздел или введите модель для поиска цен:", reply_markup=main_kb)

@dp.message(F.text == "📍 Контакты и доставка")
async def contacts(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    await message.answer(f"🏢 <b>Контакты мастерской:</b>\n\n{settings().get('contact_text', '')}", reply_markup=main_kb, parse_mode="HTML")

@dp.message(F.text == "💬 Задать вопрос мастеру")
@dp.message(Command("ask"))
async def ask_question_start(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    await state.set_state(QuestionFlow.waiting_text)
    await message.answer("✍️ Опишите вашу проблему с платой (устройство, симптомы, после чего сломалось). Мастер получит сообщение и ответит вам:", reply_markup=cancel_kb)

@dp.message(QuestionFlow.waiting_text)
async def ask_question_process(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    db_user = get_user(message.from_user.id)
    
    admin_text = (
        f"🚨 <b>Вопрос/Заявка от клиента!</b>\n\n"
        f"👤 Клиент: {db_user[2]} (@{db_user[1]})\n"
        f"📞 Телефон: {db_user[3]}\n"
        f"🆔 ID: `{db_user[0]}`\n\n"
        f"💬 Текст: {message.text}"
    )
    
    st = settings()
    admin_chat = st.get("admin_chat_id")
    for admin in (admin_chat and [admin_chat] or ADMIN_IDS):
        try: await bot.send_message(admin, admin_text, parse_mode="HTML")
        except: pass
        
    await state.clear()
    await message.answer("✅ Ваш вопрос успешно отправлен мастеру! Ожидайте ответа.", reply_markup=main_kb)

@dp.message(F.text.in_(["Отмена", "⬅️ Выйти из админки"]))
async def cancel_action(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer("Главное меню", reply_markup=main_kb)

# ======================
# SEARCH LOGIC (ПОИСК МОДЕЛИ)
# ======================
@dp.message(F.text == "🔍 Найти модель / Прайс")
async def search_start(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    await state.set_state(SearchFlow.waiting_query)
    await message.answer(
        "🔎 **Поиск по прайсу:**\n\nВведите название модели или бренда (например: `iPhone 15`, `Xiaomi`, `Poco X3`, `MacBook`):",
        reply_markup=cancel_kb,
        parse_mode="Markdown"
    )

@dp.message(SearchFlow.waiting_query)
async def process_search(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    query = message.text.strip().lower()
    data = prices()
    
    found_results = []
    for brand, models in data.get("brands", {}).items():
        for model, faults in models.items():
            match_brand = query in brand.lower()
            match_model = query in model.lower()
            match_fault = any(query in fault.lower() for fault in faults.keys())
            
            query_words = query.split()
            match_words = all(word in brand.lower() or word in model.lower() or any(word in f.lower() for f in faults.keys()) for word in query_words)

            if match_brand or match_model or match_fault or match_words:
                found_results.append((brand, model, faults))

    await state.clear()
    if not found_results:
        return await message.answer(
            f"❌ По запросу «<b>{message.text}</b>» ничего не найдено в прайсе.\n\nВоспользуйтесь кнопкой «💬 Задать вопрос мастеру», чтобы уточнить стоимость.",
            reply_markup=main_kb,
            parse_mode="HTML"
        )

    for brand, model, faults in found_results[:5]:
        text = f"📱 <b>{brand} {model}</b>\n\n<b>Прайс на услуги пайки:</b>\n"
        if faults:
            for fault, price in faults.items():
                text += f"▪️ {fault} — {price} ₽\n"
        else:
            text += "<i>Цены индивидуальны после диагностики.</i>\n"
        await message.answer(text, parse_mode="HTML")
        
    await message.answer("Главное меню:", reply_markup=main_kb)

# ======================
# ADMIN PANEL
# ======================
@dp.message(Command("admin"))
async def admin_cmd(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    await state.clear()
    await message.answer("🔐 Панель управления мастера", reply_markup=admin_kb)

@dp.message(F.text == "📥 Скачать прайс (Excel)")
async def admin_export_excel(message: types.Message):
    if message.from_user.id not in ADMIN_IDS: return
    filepath = os.path.join(BASE_DIR, "prices.xlsx")
    generate_excel_price(prices(), filepath)
    await message.answer_document(FSInputFile(filepath), caption="📄 Прайс в формате Excel. Отредактируйте его и отправьте обратно для обновления.")

@dp.message(F.text == "➕ Добавить услугу (Текстом)")
async def admin_add_text_start(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    await state.set_state(AdminFlow.waiting_add_item)
    text = (
        "✍️ **Быстрое добавление:**\n\n"
        "Отправьте данные через черточку `|` ИЛИ просто строкой:\n"
        "`Samsung s22 ultra проц 6000`\n\n"
        "Пример:\n"
        "`iPhone | iPhone 15 | Контроллер питания | 7000`"
    )
    await message.answer(text, reply_markup=cancel_kb, parse_mode="Markdown")

@dp.message(AdminFlow.waiting_add_item)
async def admin_add_text_process(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    if message.text.lower() == "отмена":
        await state.clear()
        return await message.answer("Отменено", reply_markup=admin_kb)
        
    text = message.text.strip()
    
    if "|" in text:
        parts = [p.strip() for p in text.split("|")]
        if len(parts) != 4:
            return await message.answer("❌ Ошибка формата! Нужно ровно 4 параметра через `|`:\n`Бренд | Модель | Услуга | Цена`", parse_mode="Markdown")
        brand, model, fault, price = parts
    else:
        words = text.split()
        if len(words) < 4:
            return await message.answer(
                "❌ Слишком мало данных!\n"
                "Пример правильного ввода:\n"
                "<code>Samsung s22 ultra проц 6000</code>", 
                parse_mode="HTML"
            )
        
        brand = words[0]
        price = words[-1]
        middle_words = words[1:-1]
        fault = middle_words[-1]
        model = " ".join(middle_words[:-1])

    add_price_item(brand, model, fault, price)
    await state.clear()
    await message.answer(f"✅ Успешно добавлено в прайс:\nБренд: <b>{brand}</b>\nМодель: <b>{model}</b>\nУслуга: <b>{fault}</b>\nЦена: <b>{price} ₽</b>", reply_markup=admin_kb, parse_mode="HTML")

@dp.message(F.document, F.from_user.id.in_(ADMIN_IDS))
async def admin_import_excel(message: types.Message):
    if not message.document.file_name.endswith('.xlsx'):
        return await message.answer("❌ Файл должен быть формата .xlsx")
    
    temp_path = os.path.join(BASE_DIR, "temp_prices.xlsx")
    await bot.download(message.document, destination=temp_path)
    
    try:
        new_json = parse_excel_to_json(temp_path)
        save_json(PRICES_FILE, new_json)
        await message.answer("✅ Прайс успешно обновлен из Excel!")
    except Exception as e:
        await message.answer(f"❌ Ошибка: {e}")
    finally:
        if os.path.exists(temp_path): os.remove(temp_path)

@dp.message(F.text == "👥 База клиентов")
async def admin_users_list(message: types.Message):
    if message.from_user.id not in ADMIN_IDS: return
    users = get_all_users()
    if not users: return await message.answer("База клиентов пуста.")
    
    text = "👥 <b>База клиентов (посещения):</b>\n\n"
    for u in users:
        status = "🔴 БАН" if u[4] else "🟢 Активен"
        text += f"ID: `{u[0]}` | {u[2]} | Тел: {u[3]} | {status}\n"
    for i in range(0, len(text), 4000):
        await message.answer(text[i:i+4000], parse_mode="HTML")

@dp.message(F.text == "🚫 Блокировка")
async def admin_block_menu(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Заблокировать", callback_data="adm_ban")],
        [InlineKeyboardButton(text="Разблокировать", callback_data="adm_unban")]
    ])
    await message.answer("Выберите действие с клиентом:", reply_markup=kb)

@dp.callback_query(F.data.in_(["adm_ban", "adm_unban"]))
async def admin_block_action(cb: types.CallbackQuery, state: FSMContext):
    action = 1 if cb.data == "adm_ban" else 0
    await state.update_data(block_action=action)
    await state.set_state(AdminFlow.waiting_block_id)
    await cb.message.edit_text("Введите **ID клиента** (из базы клиентов):", parse_mode="Markdown")
    await cb.answer()

@dp.message(AdminFlow.waiting_block_id)
async def process_block(message: types.Message, state: FSMContext):
    if not message.text.isdigit(): return await message.answer("ID должен состоять только из цифр.")
    data = await state.get_data()
    update_block_status(int(message.text), data['block_action'])
    await state.clear()
    status_text = "заблокирован" if data['block_action'] == 1 else "разблокирован"
    await message.answer(f"✅ Клиент с ID {message.text} успешно {status_text}.", reply_markup=admin_kb)

# CATCH-ALL
@dp.message()
async def catch_questions(message: types.Message, state: FSMContext):
    if message.text and not message.text.startswith('/') and await check_access(message, state):
        db_user = get_user(message.from_user.id)
        st = settings()
        text = f"❓ <b>Сообщение от клиента:</b>\n{db_user[2]} ({db_user[3]})\n\n{message.text}"
        admin_chat = st.get("admin_chat_id")
        for admin in (admin_chat and [admin_chat] or ADMIN_IDS):
            try: await bot.send_message(admin, text, parse_mode="HTML")
            except: pass

# ======================
# STARTUP
# ======================
async def on_startup():
    init_db()
    await bot.set_my_commands([
        BotCommand(command="start", description="Главное меню"),
        BotCommand(command="admin", description="Панель мастера")
    ])

if __name__ == "__main__":
    if not os.path.exists(SETTINGS_FILE): save_json(SETTINGS_FILE, settings())
    if not os.path.exists(PRICES_FILE): save_json(PRICES_FILE, {"brands": {}})
    
    dp.startup.register(on_startup)
    asyncio.run(dp.start_polling(bot))
