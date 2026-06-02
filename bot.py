import telebot
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime as dt, timedelta

# ========== НАСТРОЙКИ ==========
TOKEN = "8811886520:AAGM-TR31SwpDZvMYn7Hl-VFoSQVRxmrKSk"
SPREADSHEET_ID = "1GqyQAUNS-8kDop4OXuTr-2Z2dlB7SRJkJ57_F0SzaNs"

# Разрешённые пользователи (Telegram ID)
ALLOWED_USERS = [494494040, 412834655]

APARTMENTS = ["Бурнаковка 77", "Горького 152 а", "Московская 167 к 3"]

user_states = {}

# ========== ПОДКЛЮЧЕНИЕ К GOOGLE SHEETS ==========
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

sheet = client.open_by_key(SPREADSHEET_ID)
bookings_sheet = sheet.worksheet("Брони")

bot = telebot.TeleBot(TOKEN)

# ========== ПРОВЕРКА ДОСТУПА ==========
def is_allowed(message):
    return message.from_user.id in ALLOWED_USERS

# ========== КЛАВИАТУРЫ ==========
def main_menu():
    keyboard = telebot.types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    keyboard.add(
        telebot.types.KeyboardButton("📅 Добавить бронь"),
        telebot.types.KeyboardButton("❌ Отменить бронь")
    )
    keyboard.add(
        telebot.types.KeyboardButton("💰 Добавить расход"),
        telebot.types.KeyboardButton("⚡ Коммунальные")
    )
    keyboard.add(
        telebot.types.KeyboardButton("📊 Мои расходы"),
        telebot.types.KeyboardButton("📈 Доход за месяц")
    )
    keyboard.add(
        telebot.types.KeyboardButton("📆 Сегодня")
    )
    return keyboard

def apartments_keyboard():
    keyboard = telebot.types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    for apt in APARTMENTS:
        keyboard.add(telebot.types.KeyboardButton(apt))
    keyboard.add(telebot.types.KeyboardButton("🔙 Главное меню"))
    return keyboard

def months_keyboard():
    keyboard = telebot.types.ReplyKeyboardMarkup(row_width=4, resize_keyboard=True)
    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
    for m in months:
        keyboard.add(telebot.types.KeyboardButton(m))
    keyboard.add(telebot.types.KeyboardButton("🔙 Главное меню"))
    return keyboard

def platforms_keyboard():
    keyboard = telebot.types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    keyboard.add(telebot.types.KeyboardButton("Авито"), telebot.types.KeyboardButton("Суточно"))
    keyboard.add(telebot.types.KeyboardButton("🔙 Главное меню"))
    return keyboard

def cancel_keyboard():
    keyboard = telebot.types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    keyboard.add(telebot.types.KeyboardButton("🔙 Главное меню"))
    return keyboard

def remove_keyboard():
    return telebot.types.ReplyKeyboardRemove()

# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========
def get_month_worksheet(year, month):
    month_name_ru = {
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
        5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }[month]
    
    try:
        return sheet.worksheet(month_name_ru)
    except:
        try:
            template = sheet.worksheet("Июнь")
            return sheet.duplicate_sheet(template.id, new_sheet_name=month_name_ru)
        except:
            ws = sheet.add_worksheet(title=month_name_ru, rows=100, cols=36)
            headers = ["Квартира"] + [str(i) for i in range(1, 32)] + ["Доход", "Расходы", "Ком. платежи", "Итого"]
            ws.update("A1:AI1", [headers])
            for i, apt in enumerate(APARTMENTS, start=2):
                ws.update(f"A{i}", apt)
            return ws

def get_apartment_row(ws, apartment_name):
    try:
        cell = ws.find(apartment_name, in_column=1)
        return cell.row
    except:
        return None

def update_totals(ws, row):
    try:
        income = int(ws.cell(row, 33).value) if ws.cell(row, 33).value and ws.cell(row, 33).value.isdigit() else 0
        expense = int(ws.cell(row, 34).value) if ws.cell(row, 34).value and ws.cell(row, 34).value.isdigit() else 0
        utility = int(ws.cell(row, 35).value) if ws.cell(row, 35).value and ws.cell(row, 35).value.isdigit() else 0
        ws.update_cell(row, 36, income - (expense + utility))
    except:
        pass

def update_income(ws, row):
    all_bookings = bookings_sheet.get_all_values()
    if len(all_bookings) <= 1:
        ws.update_cell(row, 33, 0)
        update_totals(ws, row)
        return
    
    month_num = {
        "Январь": 1, "Февраль": 2, "Март": 3, "Апрель": 4,
        "Май": 5, "Июнь": 6, "Июль": 7, "Август": 8,
        "Сентябрь": 9, "Октябрь": 10, "Ноябрь": 11, "Декабрь": 12
    }.get(ws.title)
    
    if not month_num:
        return
    
    apartment_name = ws.cell(row, 1).value
    total = 0
    for booking in all_bookings[1:]:
        if len(booking) >= 5 and booking[1] == apartment_name:
            try:
                booking_month = int(booking[3].split(".")[1])
                if booking_month == month_num:
                    total += int(booking[6]) if booking[6].isdigit() else 0
            except:
                pass
    ws.update_cell(row, 33, total)
    update_totals(ws, row)

def write_to_calendar(apartment_name, start_day, end_day, month, year, platform):
    ws = get_month_worksheet(year, month)
    row = get_apartment_row(ws, apartment_name)
    if not row:
        return False
    
    for day in range(start_day, end_day + 1):
        col = day + 1
        current = ws.cell(row, col).value
        if day == start_day:
            ws.update_cell(row, col, "Заезд" if current != "Выезд" else "Выезд/Заезд")
        elif day == end_day:
            ws.update_cell(row, col, "Выезд" if current != "Заезд" else "Выезд/Заезд")
        else:
            ws.update_cell(row, col, platform)
    
    update_income(ws, row)
    return True

def add_booking_to_sheet(booking_id, apartment, platform, start_date, end_date, total):
    now = dt.now().strftime("%d.%m.%Y %H:%M")
    price = ""
    row = [booking_id, apartment, platform, start_date, end_date, price, total, now]
    bookings_sheet.append_row(row)

def delete_booking_from_sheet(apartment, start_date, end_date):
    all_bookings = bookings_sheet.get_all_values()
    for i, row in enumerate(all_bookings, start=1):
        if len(row) >= 5 and row[1] == apartment and row[3] == start_date and row[4] == end_date:
            bookings_sheet.delete_rows(i)
            return True
    return False

def clear_calendar(apartment_name, start_day, end_day, month, year):
    ws = get_month_worksheet(year, month)
    row = get_apartment_row(ws, apartment_name)
    if not row:
        return False
    for day in range(start_day, end_day + 1):
        col = day + 1
        ws.update_cell(row, col, "")
    update_income(ws, row)
    return True

def update_expense_or_utility(apartment_name, amount, expense_type, month, year):
    ws = get_month_worksheet(year, month)
    row = get_apartment_row(ws, apartment_name)
    if not row:
        return False, f"Квартира не найдена"
    
    col = 34 if expense_type == "expense" else 35
    current = ws.cell(row, col).value
    current_val = int(current) if current and current.isdigit() else 0
    
    if expense_type == "expense":
        new_val = current_val + amount
    else:
        new_val = amount
    
    ws.update_cell(row, col, new_val)
    update_totals(ws, row)
    return True, f"{'Расходы' if expense_type == 'expense' else 'Ком. платежи'}: {new_val} ₽"

def get_expenses(apartment_name, month, year):
    ws = get_month_worksheet(year, month)
    row = get_apartment_row(ws, apartment_name)
    if not row:
        return None, None
    expense = ws.cell(row, 34).value
    utility = ws.cell(row, 35).value
    return (int(expense) if expense and expense.isdigit() else 0,
            int(utility) if utility and utility.isdigit() else 0)

def get_user_active_bookings(apartment_name):
    all_bookings = bookings_sheet.get_all_values()
    result = []
    for booking in all_bookings[1:]:
        if len(booking) >= 5 and booking[1] == apartment_name:
            result.append({
                "start": booking[3],
                "end": booking[4],
                "platform": booking[2],
                "total": booking[6]
            })
    return result

# ========== КОМАНДЫ БОТА ==========
@bot.message_handler(commands=['start'])
def start(message):
    if not is_allowed(message):
        bot.send_message(message.chat.id, "⛔ Доступ запрещён.")
        return
    bot.send_message(message.chat.id, 
                     "🏠 *Бот для учёта аренды*\n\n"
                     "Нажимайте на кнопки ниже, чтобы управлять бронями и расходами.\n\n"
                     "📅 *Добавить бронь* — новая бронь\n"
                     "❌ *Отменить бронь* — отмена существующей\n"
                     "💰 *Добавить расход* — добавить сумму в расходы\n"
                     "⚡ *Коммунальные* — установить ком. платежи\n"
                     "📊 *Мои расходы* — посмотреть расходы за месяц\n"
                     "📈 *Доход за месяц* — посчитать доход\n"
                     "📆 *Сегодня* — заезды и выезды сегодня\n\n"
                     "⚠️ *Важно:* при добавлении брони указывайте итоговую сумму (уже за вычетом комиссии площадки)",
                     reply_markup=main_menu(),
                     parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "🔙 Главное меню")
def back_to_menu(message):
    if not is_allowed(message):
        return
    if message.chat.id in user_states:
        del user_states[message.chat.id]
    bot.send_message(message.chat.id, "🏠 *Главное меню*", reply_markup=main_menu(), parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "📅 Добавить бронь")
def cmd_close(message):
    if not is_allowed(message):
        return
    user_states[message.chat.id] = {"command": "close", "step": "apartment"}
    bot.send_message(message.chat.id, "🏠 *Выберите квартиру:*", reply_markup=apartments_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "❌ Отменить бронь")
def cmd_cancel(message):
    if not is_allowed(message):
        return
    user_states[message.chat.id] = {"command": "cancel", "step": "apartment"}
    bot.send_message(message.chat.id, "🏠 *Выберите квартиру:*", reply_markup=apartments_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "💰 Добавить расход")
def cmd_expense(message):
    if not is_allowed(message):
        return
    user_states[message.chat.id] = {"command": "expense", "step": "apartment"}
    bot.send_message(message.chat.id, "🏠 *Выберите квартиру:*", reply_markup=apartments_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "⚡ Коммунальные")
def cmd_utility(message):
    if not is_allowed(message):
        return
    user_states[message.chat.id] = {"command": "utility", "step": "apartment"}
    bot.send_message(message.chat.id, "🏠 *Выберите квартиру:*", reply_markup=apartments_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "📊 Мои расходы")
def cmd_expenses(message):
    if not is_allowed(message):
        return
    user_states[message.chat.id] = {"command": "expenses", "step": "apartment"}
    bot.send_message(message.chat.id, "🏠 *Выберите квартиру:*", reply_markup=apartments_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "📈 Доход за месяц")
def cmd_month(message):
    if not is_allowed(message):
        return
    user_states[message.chat.id] = {"command": "month", "step": "apartment"}
    bot.send_message(message.chat.id, "🏠 *Выберите квартиру:*", reply_markup=apartments_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda message: message.text == "📆 Сегодня")
def cmd_today(message):
    if not is_allowed(message):
        return
    today = dt.now()
    ws = get_month_worksheet(today.year, today.month)
    col = today.day + 1
    events = {"Заезд": [], "Выезд": [], "Выезд/Заезд": []}
    
    all_data = ws.get_all_values()
    for row in all_data[1:]:
        if row and row[0] and col <= len(row):
            value = row[col - 1]
            if value == "Заезд":
                events["Заезд"].append(row[0])
            elif value == "Выезд":
                events["Выезд"].append(row[0])
            elif value == "Выезд/Заезд":
                events["Выезд/Заезд"].append(row[0])
    
    response = f"📅 *{today.strftime('%d.%m.%Y')}*\n\n"
    response += "🚪 *Заезд:*\n" + "\n".join(f"  • {a}" for a in events["Заезд"]) + "\n\n" if events["Заезд"] else "🚪 *Заездов нет*\n\n"
    response += "🚪 *Выезд:*\n" + "\n".join(f"  • {a}" for a in events["Выезд"]) + "\n\n" if events["Выезд"] else "🚪 *Выездов нет*\n\n"
    if events["Выезд/Заезд"]:
        response += "🔄 *Суета:*\n" + "\n".join(f"  • {a}" for a in events["Выезд/Заезд"]) + "\n"
    
    bot.send_message(message.chat.id, response, reply_markup=main_menu(), parse_mode="Markdown")

# ========== ОБРАБОТЧИК СООБЩЕНИЙ ДЛЯ ДИАЛОГОВ ==========
@bot.message_handler(func=lambda message: message.chat.id in user_states and is_allowed(message))
def handle_dialog(message):
    if message.text == "🔙 Главное меню":
        back_to_menu(message)
        return
    
    state = user_states[message.chat.id]
    command = state["command"]
    step = state["step"]
    
    if command == "close":
        if step == "apartment":
            if message.text not in APARTMENTS:
                bot.send_message(message.chat.id, "❌ *Выберите квартиру из списка кнопками*", reply_markup=apartments_keyboard(), parse_mode="Markdown")
                return
            state["apartment"] = message.text
            state["step"] = "dates"
            bot.send_message(message.chat.id, "📅 *Введите даты в формате ДД.ММ-ДД.ММ*\nПример: 16.06-17.06", reply_markup=cancel_keyboard(), parse_mode="Markdown")
        
        elif step == "dates":
            match = re.search(r"(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})", message.text)
            if not match:
                bot.send_message(message.chat.id, "❌ *Неверный формат. Пример: 16.06-17.06*", reply_markup=cancel_keyboard(), parse_mode="Markdown")
                return
            state["start_day"] = int(match.group(1))
            state["start_month"] = int(match.group(2))
            state["end_day"] = int(match.group(3))
            state["end_month"] = int(match.group(4))
            state["step"] = "platform"
            bot.send_message(message.chat.id, "🏷 *Выберите площадку:*", reply_markup=platforms_keyboard(), parse_mode="Markdown")
        
        elif step == "platform":
            if message.text not in ["Авито", "Суточно"]:
                bot.send_message(message.chat.id, "❌ *Выберите площадку из списка*", reply_markup=platforms_keyboard(), parse_mode="Markdown")
                return
            state["platform"] = message.text
            state["step"] = "total"
            bot.send_message(message.chat.id, "💰 *Введите итоговую сумму за бронь (цифры):*\n(уже за вычетом комиссии площадки)", reply_markup=cancel_keyboard(), parse_mode="Markdown")
        
        elif step == "total":
            if not message.text.isdigit():
                bot.send_message(message.chat.id, "❌ *Введите число*", reply_markup=cancel_keyboard(), parse_mode="Markdown")
                return
            total = int(message.text)
            year = dt.now().year
            start_date = dt(year, state["start_month"], state["start_day"])
            end_date = dt(year, state["end_month"], state["end_day"])
            nights = max((end_date - start_date).days, 1)
            booking_id = f"#{state['apartment'].replace(' ', '')[:6]}{state['start_day']}{state['end_day']}"
            
            success = write_to_calendar(state["apartment"], state["start_day"], state["end_day"], state["start_month"], year, state["platform"])
            if not success:
                bot.send_message(message.chat.id, "❌ *Ошибка записи в календарь*", reply_markup=main_menu(), parse_mode="Markdown")
                del user_states[message.chat.id]
                return
            
            add_booking_to_sheet(booking_id, state["apartment"], state["platform"], 
                                f"{state['start_day']:02d}.{state['start_month']:02d}", 
                                f"{state['end_day']:02d}.{state['end_month']:02d}", total)
            
            response = f"✅ *Бронь {booking_id}*\n🏠 {state['apartment']}\n📅 {state['start_day']:02d}.{state['start_month']:02d} - {state['end_day']:02d}.{state['end_month']:02d} ({nights} ночей)\n🏷 {state['platform']}\n💰 Сумма: *{total} ₽*"
            bot.send_message(message.chat.id, response, reply_markup=main_menu(), parse_mode="Markdown")
            del user_states[message.chat.id]
    
    elif command == "cancel":
        if step == "apartment":
            if message.text not in APARTMENTS:
                bot.send_message(message.chat.id, "❌ *Выберите квартиру из списка*", reply_markup=apartments_keyboard(), parse_mode="Markdown")
                return
            state["apartment"] = message.text
            bookings = get_user_active_bookings(state["apartment"])
            if not bookings:
                bot.send_message(message.chat.id, f"❌ *Нет активных броней для {state['apartment']}*", reply_markup=main_menu(), parse_mode="Markdown")
                del user_states[message.chat.id]
                return
            state["bookings"] = bookings
            state["step"] = "select_booking"
            keyboard = telebot.types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
            for b in bookings:
                keyboard.add(telebot.types.KeyboardButton(f"{b['start']}-{b['end']} {b['platform']} {b['total']}₽"))
            keyboard.add(telebot.types.KeyboardButton("🔙 Главное меню"))
            bot.send_message(message.chat.id, "📅 *Выберите бронь для отмены:*", reply_markup=keyboard, parse_mode="Markdown")
        
        elif step == "select_booking":
            selected = message.text
            for b in state["bookings"]:
                if selected.startswith(f"{b['start']}-{b['end']}"):
                    start_day, start_month = map(int, b['start'].split("."))
                    end_day, end_month = map(int, b['end'].split("."))
                    year = dt.now().year
                    clear_calendar(state["apartment"], start_day, end_day, start_month, year)
                    delete_booking_from_sheet(state["apartment"], b['start'], b['end'])
                    ws = get_month_worksheet(year, start_month)
                    row = get_apartment_row(ws, state["apartment"])
                    if row:
                        update_income(ws, row)
                    bot.send_message(message.chat.id, f"✅ *Бронь отменена*\n🏠 {state['apartment']}\n📅 {b['start']} - {b['end']}", reply_markup=main_menu(), parse_mode="Markdown")
                    del user_states[message.chat.id]
                    return
            bot.send_message(message.chat.id, "❌ *Выберите бронь из списка*", reply_markup=cancel_keyboard(), parse_mode="Markdown")
    
    elif command == "expense":
        if step == "apartment":
            if message.text not in APARTMENTS:
                bot.send_message(message.chat.id, "❌ *Выберите квартиру из списка*", reply_markup=apartments_keyboard(), parse_mode="Markdown")
                return
            state["apartment"] = message.text
            state["step"] = "amount"
            bot.send_message(message.chat.id, "💰 *Введите сумму расхода (цифры):*", reply_markup=cancel_keyboard(), parse_mode="Markdown")
        
        elif step == "amount":
            if not message.text.isdigit():
                bot.send_message(message.chat.id, "❌ *Введите число*", reply_markup=cancel_keyboard(), parse_mode="Markdown")
                return
            amount = int(message.text)
            year = dt.now().year
            month = dt.now().month
            success, msg = update_expense_or_utility(state["apartment"], amount, "expense", month, year)
            bot.send_message(message.chat.id, f"{'✅' if success else '❌'} *{msg}*", reply_markup=main_menu(), parse_mode="Markdown")
            del user_states[message.chat.id]
    
    elif command == "utility":
        if step == "apartment":
            if message.text not in APARTMENTS:
                bot.send_message(message.chat.id, "❌ *Выберите квартиру из списка*", reply_markup=apartments_keyboard(), parse_mode="Markdown")
                return
            state["apartment"] = message.text
            state["step"] = "amount"
            bot.send_message(message.chat.id, "⚡ *Введите сумму коммунальных платежей (цифры):*", reply_markup=cancel_keyboard(), parse_mode="Markdown")
        
        elif step == "amount":
            if not message.text.isdigit():
                bot.send_message(message.chat.id, "❌ *Введите число*", reply_markup=cancel_keyboard(), parse_mode="Markdown")
                return
            amount = int(message.text)
            year = dt.now().year
            month = dt.now().month
            success, msg = update_expense_or_utility(state["apartment"], amount, "utility", month, year)
            bot.send_message(message.chat.id, f"{'✅' if success else '❌'} *{msg}*", reply_markup=main_menu(), parse_mode="Markdown")
            del user_states[message.chat.id]
    
    elif command == "expenses":
        if step == "apartment":
            if message.text not in APARTMENTS:
                bot.send_message(message.chat.id, "❌ *Выберите квартиру из списка*", reply_markup=apartments_keyboard(), parse_mode="Markdown")
                return
            apartment = message.text
            year = dt.now().year
            month = dt.now().month
            expense, utility = get_expenses(apartment, month, year)
            month_name = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"][month-1]
            response = f"📊 *Расходы {apartment} за {month_name} {year}:*\n\n"
            response += f"💰 Расходы: *{expense} ₽*\n"
            response += f"⚡ Ком. платежи: *{utility} ₽*\n"
            response += f"💸 Итого: *{expense + utility} ₽*"
            bot.send_message(message.chat.id, response, reply_markup=main_menu(), parse_mode="Markdown")
            del user_states[message.chat.id]
    
    elif command == "month":
        if step == "apartment":
            if message.text not in APARTMENTS:
                bot.send_message(message.chat.id, "❌ *Выберите квартиру из списка*", reply_markup=apartments_keyboard(), parse_mode="Markdown")
                return
            state["apartment"] = message.text
            state["step"] = "month"
            bot.send_message(message.chat.id, "📆 *Выберите месяц:*", reply_markup=months_keyboard(), parse_mode="Markdown")
        
        elif step == "month":
            month_map = {"Янв":1, "Фев":2, "Мар":3, "Апр":4, "Май":5, "Июн":6, "Июл":7, "Авг":8, "Сен":9, "Окт":10, "Ноя":11, "Дек":12}
            if message.text not in month_map:
                bot.send_message(message.chat.id, "❌ *Выберите месяц из списка*", reply_markup=months_keyboard(), parse_mode="Markdown")
                return
            month = month_map[message.text]
            all_bookings = bookings_sheet.get_all_values()
            total = 0
            for row in all_bookings[1:]:
                if len(row) >= 5 and row[1] == state["apartment"]:
                    try:
                        booking_month = int(row[3].split(".")[1])
                        if booking_month == month:
                            total += int(row[6]) if row[6].isdigit() else 0
                    except:
                        pass
            bot.send_message(message.chat.id, f"📊 *Доход {state['apartment']} за {month:02d}.2026: {total} ₽*", reply_markup=main_menu(), parse_mode="Markdown")
            del user_states[message.chat.id]

if __name__ == "__main__":
    bot.infinity_polling()
