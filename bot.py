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
MODELS_PER_PAGE = 10

# Базовый список брендов, которые ВСЕГДА должны быть (включая универсальные)
DEFAULT_BRANDS = [
    "iPhone", "iPad", "Samsung", "Xiaomi", "Poco", 
    "Honor/Huawei", "Infinix", "Realme", "Tecno", 
    "iQoo", "Oppo", "Vivo", "OnePlus", "Nothing", 
    "Asus", "Motorola", "Google Pixel"
]

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
    cursor.execute('SELECT user_id, full_name, phone, is_blocked FROM users')
    users = cursor.fetchall()
    conn.close()
    return users

# ======================
# DATA LOADERS & FIXERS
# ======================
def load_json(path: str, default):
    if not os.path.exists(path): return default
    with open(path, "r", encoding="utf-8") as f: return json.load(f)

def save_json(path: str, data):
    with open(path, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)

def settings():
    return load_json(SETTINGS_FILE, {
        "contact_text": "Мы — сеть специализированных мастерских по сложному компонентному ремонту.\nОставьте заявку, и мы подберем ближайший к вам сервисный центр!",
        "admin_chat_id": None
    })

def prices():
    data = load_json(PRICES_FILE, {"brands": {}})
    # Гарантируем наличие всех брендов
    if "brands" not in data:
        data["brands"] = {}
    for b in DEFAULT_BRANDS:
        if b not in data["brands"]:
            data["brands"][b] = {}
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
    ws.title = "Прайс-лист"
    ws.append(["Бренд", "Модель", "Поломка", "Цена (руб)"])
    
    for brand, models in data.get("brands", {}).items():
        if not models:
            ws.append([brand, "УНИВЕРСАЛЬНЫЙ", "-", "-"])
            continue
        for model, faults in models.items():
            if not faults:
                ws.append([brand, model, "-", "-"])
                continue
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
    for b in DEFAULT_BRANDS:
        new_data["brands"][b] = {}
        
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]: continue
        brand, model, fault, price = [str(x).strip() if x is not None else "" for x in row]
        
        if brand not in new_data["brands"]:
            new_data["brands"][brand] = {}
            
        if model and model != "УНИВЕРСАЛЬНЫЙ" and model != "-":
            if model not in new_data["brands"][brand]:
                new_data["brands"][brand][model] = {}
            if fault and fault != "-":
                new_data["brands"][brand][model][fault] = price
                
    return new_data

# ======================
# KEYBOARDS
# ======================
main_kb = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔬 Ремонт материнских плат")],
        [KeyboardButton(text="📸 Контакты"), KeyboardButton(text="❓ Задать вопрос")]
    ],
    resize_keyboard=True,
)

cancel_kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="Отмена")]], resize_keyboard=True)
phone_request_kb = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="📞 Отправить номер", request_contact=True)]],
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

def brands_kb(brands_list: List[str]) -> InlineKeyboardMarkup:
    rows = []
    for i in range(0, len(brands_list), 2):
        row = [InlineKeyboardButton(text=brands_list[i], callback_data=f"b:{i}")]
        if i + 1 < len(brands_list):
            row.append(InlineKeyboardButton(text=brands_list[i+1], callback_data=f"b:{i+1}"))
        rows.append(row)
    return InlineKeyboardMarkup(inline_keyboard=rows)

def models_kb(brand_idx: int, models_list: List[str], page: int) -> InlineKeyboardMarkup:
    rows = []
    pages = max(1, (len(models_list) + MODELS_PER_PAGE - 1) // MODELS_PER_PAGE)
    start = page * MODELS_PER_PAGE
    
    for i, model in enumerate(models_list[start:start + MODELS_PER_PAGE]):
        rows.append([InlineKeyboardButton(text=model, callback_data=f"m:{brand_idx}:{start + i}")])
        
    nav = []
    if page > 0: nav.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"p:{brand_idx}:{page-1}"))
    if page < pages - 1: nav.append(InlineKeyboardButton(text="Вперед ➡️", callback_data=f"p:{brand_idx}:{page+1}"))
    if nav: rows.append(nav)
    
    rows.append([InlineKeyboardButton(text="🔙 К выбору бренда", callback_data="back_to_brands")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

# ======================
# FSM STATES
# ======================
class RegFlow(StatesGroup):
    name = State()
    phone = State()

class RequestFlow(StatesGroup):
    waiting_comment = State()

class AdminFlow(StatesGroup):
    waiting_add_item = State()
    waiting_block_id = State()

# ======================
# MIDDLEWARE / AUTH
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
        await message.answer("Здравствуйте! 👋\nДля доступа к сервису, пожалуйста, введите ваше **Имя**:", parse_mode="Markdown")
        return False

# ======================
# REGISTRATION
# ======================
@dp.message(RegFlow.name)
async def reg_name(message: types.Message, state: FSMContext):
    if not message.text: return await message.answer("Пожалуйста, введите ваше имя текстом.")
    await state.update_data(full_name=message.text)
    await state.set_state(RegFlow.phone)
    await message.answer(f"Приятно познакомиться, {message.text}!\nТеперь нажмите кнопку ниже, чтобы поделиться номером телефона.", reply_markup=phone_request_kb)

@dp.message(RegFlow.phone)
async def reg_phone(message: types.Message, state: FSMContext):
    if not message.contact:
        return await message.answer("Пожалуйста, используйте кнопку «📞 Отправить номер» внизу экрана.")
    
    data = await state.get_data()
    u = message.from_user
    add_user(u.id, u.username, data['full_name'], message.contact.phone_number)
    
    await state.clear()
    await message.answer("✅ Регистрация завершена! Добро пожаловать.", reply_markup=main_kb)
    
    st = settings()
    chat_id = st.get("admin_chat_id")
    for admin in (chat_id and [chat_id] or ADMIN_IDS):
        try: await bot.send_message(admin, f"👤 <b>Новый клиент:</b>\nИмя: {data['full_name']}\nТелефон: {message.contact.phone_number}\nID: {u.id}", parse_mode="HTML")
        except: pass

# ======================
# CLIENT HANDLERS
# ======================
@dp.message(CommandStart())
async def start(message: types.Message, state: FSMContext):
    await state.clear()
    if await check_access(message, state):
        await message.answer("Выберите нужный раздел:", reply_markup=main_kb)

@dp.message(F.text == "📸 Контакты")
async def contacts(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    await message.answer(f"🏢 <b>О нас:</b>\n\n{settings().get('contact_text', '')}", reply_markup=main_kb, parse_mode="HTML")

@dp.message(F.text == "❓ Задать вопрос")
@dp.message(Command("ask"))
async def ask_question(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    await state.clear()
    await message.answer("Напишите ваш вопрос сюда. Мастер получит его и свяжется с вами.", reply_markup=main_kb)

@dp.message(F.text.in_(["Отмена", "⬅️ Выйти из админки"]))
async def cancel_action(message: types.Message, state: FSMContext):
    await state.clear()
    if message.from_user.id in ADMIN_IDS:
        await message.answer("Главное меню", reply_markup=main_kb)
    else:
        await message.answer("Главное меню", reply_markup=main_kb)

@dp.message(F.text == "🔬 Ремонт материнских плат")
async def repair_start(message: types.Message, state: FSMContext):
    if not await check_access(message, state): return
    brands = list(prices().get("brands", {}).keys())
    await message.answer("Выберите бренд вашего устройства:", reply_markup=brands_kb(brands))

@dp.callback_query(F.data == "back_to_brands")
async def back_to_brands(cb: types.CallbackQuery):
    brands = list(prices().get("brands", {}).keys())
    await cb.message.edit_text("Выберите бренд вашего устройства:", reply_markup=brands_kb(brands))
    await cb.answer()

@dp.callback_query(F.data.startswith("b:") | F.data.startswith("p:"))
async def show_models(cb: types.CallbackQuery):
    parts = cb.data.split(":")
    brand_idx, page = int(parts[1]), int(parts[2]) if len(parts) > 2 else 0
    data = prices()
    brands = list(data["brands"].keys())
    
    brand_name = brands[brand_idx]
    models_list = list(data["brands"][brand_name].keys())
    
    # ЕСЛИ У БРЕНДА НЕТ МОДЕЛЕЙ (универсальный текст для китайцев и т.д.)
    if not models_list:
        text = (
            f"🔬 <b>Компонентный ремонт: {brand_name}</b>\n\n"
            f"<b>Типовые неисправности по плате:</b>\n"
            f"▪️ Отвал процессора (Реболл) — от 3000 до 5000 ₽\n"
            f"▪️ Восстановление после влаги — ⚠️ Точная стоимость ремонта определяется мастером после диагностики и замеров напряжений."
        )
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✍️ Оставить заявку на ремонт", callback_data=f"req:{brand_idx}:UNIVERSAL")],
            [InlineKeyboardButton(text="🔙 К выбору бренда", callback_data="back_to_brands")]
        ])
        await cb.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
        return await cb.answer()
        
    await cb.message.edit_text(f"Бренд: <b>{brand_name}</b>\nВыберите вашу модель:", reply_markup=models_kb(brand_idx, models_list, page), parse_mode="HTML")
    await cb.answer()

@dp.callback_query(F.data.startswith("m:"))
async def show_faults(cb: types.CallbackQuery):
    _, brand_idx, model_idx = cb.data.split(":")
    data = prices()
    brand_name = list(data["brands"].keys())[int(brand_idx)]
    model_name = list(data["brands"][brand_name].keys())[int(model_idx)]
    faults = data["brands"][brand_name][model_name]
    
    text = f"🔬 <b>Компонентный ремонт: {brand_name} {model_name}</b>\n\n<b>Типовые неисправности:</b>\n"
    if faults:
        for fault, price in faults.items(): text += f"▪️ {fault} — {price} ₽\n"
    else:
        text += "<i>Цены уточняются после глубокой диагностики платы.</i>\n"
    text += "\n⚠️ <i>Точная стоимость ремонта определяется мастером после диагностики.</i>"
    
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✍️ Оставить заявку на ремонт", callback_data=f"req:{brand_idx}:{model_idx}")],
        [InlineKeyboardButton(text="🔙 Назад к моделям", callback_data=f"b:{brand_idx}")]
    ])
    await cb.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
    await cb.answer()

# ======================
# REQUEST CREATION
# ======================
@dp.callback_query(F.data.startswith("req:"))
async def request_start(cb: types.CallbackQuery, state: FSMContext):
    _, brand_idx, model_idx = cb.data.split(":")
    data = prices()
    brand_name = list(data["brands"].keys())[int(brand_idx)]
    
    if model_idx == "UNIVERSAL":
        model_str = f"{brand_name} (модель укажет клиент)"
    else:
        model_str = f"{brand_name} {list(data['brands'][brand_name].keys())[int(model_idx)]}"
    
    await state.update_data(req_model=model_str)
    await state.set_state(RequestFlow.waiting_comment)
    await cb.message.answer(f"📝 Заявка: <b>{model_str}</b>\n\nОпишите проблему своими словами (что случилось, после чего перестал работать):", reply_markup=cancel_kb, parse_mode="HTML")
    await cb.answer()

@dp.message(RequestFlow.waiting_comment)
async def request_comment(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    db_user = get_user(message.from_user.id)
    
    admin_text = (
        f"🚨 <b>Новая заявка на ремонт платы!</b>\n\n"
        f"📱 Устройство: <b>{user_data['req_model']}</b>\n"
        f"👤 Клиент: {db_user[2]} (@{db_user[1]})\n"
        f"📞 Телефон: {db_user[3]}\n"
        f"💬 Описание: {message.text}"
    )
    
    st = settings()
    for admin in (st.get("admin_chat_id") and [st.get("admin_chat_id")] or ADMIN_IDS):
        try: await bot.send_message(admin, admin_text, parse_mode="HTML")
        except: pass
        
    await state.clear()
    await message.answer("✅ Заявка успешно отправлена мастеру! С вами скоро свяжутся.", reply_markup=main_kb)

# ======================
# ADMIN PANEL (TEXT ADD & EXCEL & CRM)
# ======================
@dp.message(Command("admin"))
async def admin_cmd(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    await state.clear()
    await message.answer("🔐 Панель управления", reply_markup=admin_kb)

@dp.message(F.text == "📥 Скачать прайс (Excel)")
async def admin_export_excel(message: types.Message):
    if message.from_user.id not in ADMIN_IDS: return
    filepath = os.path.join(BASE_DIR, "prices.xlsx")
    generate_excel_price(prices(), filepath)
    await message.answer_document(FSInputFile(filepath), caption="📄 Ваш прайс в формате Excel.\nОтредактируйте его и отправьте файл обратно мне для обновления базы.")

@dp.message(F.text == "➕ Добавить услугу (Текстом)")
async def admin_add_text_start(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    await state.set_state(AdminFlow.waiting_add_item)
    text = (
        "✍️ **Добавление услуги через чат:**\n\n"
        "Отправьте данные в одном сообщении через черточку `|` в таком формате:\n"
        "`Бренд | Модель | Услуга / Поломка | Цена`\n\n"
        "Пример:\n"
        "`iPhone | iPhone 14 | Контроллер питания | 5000`\n"
        "или для китайцев (без модели):\n"
        "`Poco | - | Замена КП | 3500`\n\n"
        "Введите данные или нажмите «Отмена»:"
    )
    await message.answer(text, reply_markup=cancel_kb, parse_mode="Markdown")

@dp.message(AdminFlow.waiting_add_item)
async def admin_add_text_process(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    if message.text.lower() == "отмена":
        await state.clear()
        return await message.answer("Отменено", reply_markup=admin_kb)
        
    parts = [p.strip() for p in message.text.split("|")]
    if len(parts) != 4:
        return await message.answer("❌ Неверный формат! Нужно строго 4 части через `|`:\n`Бренд | Модель | Услуга | Цена`\nПопробуйте еще раз или нажмите Отмена.", parse_mode="Markdown")
        
    brand, model, fault, price = parts
    if model == "-": model = "" # если модель не нужна
    
    add_price_item(brand, model, fault, price)
    await state.clear()
    await message.answer(f"✅ Успешно добавлено!\nБренд: {brand} | Модель: {model or 'Общая'} | {fault} — {price} ₽", reply_markup=admin_kb)

@dp.message(F.document, F.from_user.id.in_(ADMIN_IDS))
async def admin_import_excel(message: types.Message):
    if not message.document.file_name.endswith('.xlsx'):
        return await message.answer("❌ Файл должен иметь расширение .xlsx")
    
    file_id = message.document.file_id
    temp_path = os.path.join(BASE_DIR, "temp_prices.xlsx")
    await bot.download(message.document, destination=temp_path)
    
    try:
        new_json = parse_excel_to_json(temp_path)
        save_json(PRICES_FILE, new_json)
        await message.answer("✅ База цен успешно обновлена из Excel!")
    except Exception as e:
        await message.answer(f"❌ Ошибка при чтении Excel: {e}")
    finally:
        if os.path.exists(temp_path): os.remove(temp_path)

@dp.message(F.text == "👥 База клиентов")
async def admin_users_list(message: types.Message):
    if message.from_user.id not in ADMIN_IDS: return
    users = get_all_users()
    if not users: return await message.answer("База клиентов пуста.")
    
    text = "👥 <b>Список клиентов:</b>\n\n"
    for u in users:
        status = "🔴 БАН" if u[3] else "🟢 Активен"
        text += f"ID: `{u[0]}` | {u[1]} | {u[2]} | {status}\n"
    
    for i in range(0, len(text), 4000):
        await message.answer(text[i:i+4000], parse_mode="HTML")

@dp.message(F.text == "🚫 Блокировка")
async def admin_block_menu(message: types.Message, state: FSMContext):
    if message.from_user.id not in ADMIN_IDS: return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Заблокировать", callback_data="adm_ban")],
        [InlineKeyboardButton(text="Разблокировать", callback_data="adm_unban")]
    ])
    await message.answer("Выберите действие:", reply_markup=kb)

@dp.callback_query(F.data.in_(["adm_ban", "adm_unban"]))
async def admin_block_action(cb: types.CallbackQuery, state: FSMContext):
    action = 1 if cb.data == "adm_ban" else 0
    await state.update_data(block_action=action)
    await state.set_state(AdminFlow.waiting_block_id)
    await cb.message.edit_text("Введите **ID клиента** (цифры из базы клиентов):", parse_mode="Markdown")
    await cb.answer()

@dp.message(AdminFlow.waiting_block_id)
async def process_block(message: types.Message, state: FSMContext):
    if not message.text.isdigit(): return await message.answer("ID должен состоять только из цифр. Попробуйте еще раз или нажмите Отмена.")
    
    data = await state.get_data()
    action = data['block_action']
    update_block_status(int(message.text), action)
    
    status = "заблокирован" if action == 1 else "разблокирован"
    await state.clear()
    await message.answer(f"✅ Клиент с ID {message.text} успешно {status}.", reply_markup=admin_kb)

# CATCH-ALL FOR MESSAGES
@dp.message()
async def catch_questions(message: types.Message, state: FSMContext):
    if message.text and not message.text.startswith('/') and await check_access(message, state):
        db_user = get_user(message.from_user.id)
        st = settings()
        text = f"❓ <b>Вопрос от клиента:</b>\n{db_user[2]} ({db_user[3]})\n\n{message.text}"
        for admin in (st.get("admin_chat_id") and [st.get("admin_chat_id")] or ADMIN_IDS):
            try: await bot.send_message(admin, text, parse_mode="HTML")
            except: pass

# ======================
# STARTUP HOOK
# ======================
async def on_startup():
    init_db()
    await bot.set_my_commands([
        BotCommand(command="start", description="Главное меню / Перезапуск"),
        BotCommand(command="admin", description="Панель управления (для мастера)")
    ])

if __name__ == "__main__":
    if not os.path.exists(SETTINGS_FILE): save_json(SETTINGS_FILE, settings())
    if not os.path.exists(PRICES_FILE): save_json(PRICES_FILE, {"brands": {}})
    
    dp.startup.register(on_startup)
    asyncio.run(dp.start_polling(bot))
