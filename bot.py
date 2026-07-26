from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    ConversationHandler,
    filters,
)

from datetime import datetime
import json
import os
import openpyxl
from openpyxl import Workbook, load_workbook

# =========================
# التوكن
# =========================
TOKEN = "8403208232:AAGyopVwOIsxUJ1oisZpupHCM5HZufXJj9M"

# =========================
# الإحصائيات
# =========================
statistics = {
    "cash": 0,
    "debt": 0,
}

# =========================
# الموظفين
# =========================
EMPLOYEES = [
    "المدير ياسر",
    "نائب المدير غدير",
    "الأستاذة قمر",
    "ريم",
    "حسن",
    "دانيال",
    "محمد",
    "يزن",
    "حيدر",
    "🔙 رجوع"
]

# =========================
# الملفات
# =========================
MAINTENANCE_FILE = "maintenance_requests.json"
INSTALLATION_FILE = "installations.json"
INVENTORY_FILE = "inventory.json"
STATISTICS_FILE = "statistics.json"
SALES_FILE = "sales.json"
COLLECTED_DEBTS_FILE = "collected_debts.json"
EXCEL_FILE = "operations_log.xlsx"

# =========================
# الحالات
# =========================
(
    MAINTENANCE_NAME,
    MAINTENANCE_NATIONAL,
    MAINTENANCE_PROBLEM,
    MAINTENANCE_LOCATION,
    MAINTENANCE_PHONE,
    MAINTENANCE_REQUESTER,
    MAINTENANCE_ASSIGNEE,  # new state
    MAINTENANCE_CONFIRM,
    SEARCH_MAINTENANCE,
    ADD_PRODUCT,
    ADD_PRODUCT_QUANTITY,
    SALE_TYPE,
    PRODUCT_TYPE,
    PRODUCT_QUANTITY,
    PRODUCT_PRICE,
    PRODUCT_EMPLOYEE,
    PRODUCT_PAYMENT,
    COLLECT_AMOUNT,
    COLLECT_CUSTOMER_NAME,
    COLLECT_EMPLOYEE,
    COLLECT_CONFIRM,
    INSTALLATION_NAME,
    INSTALLATION_NATIONAL,
    INSTALLATION_TYPE,
    INSTALLATION_LOCATION,
    INSTALLATION_PHONE,
    INSTALLATION_REQUESTER,
    INSTALLATION_ASSIGNEE,
    INSTALLATION_CONFIRM,
    SEARCH_INSTALLATION,
) = range(30)

# =========================
# الكيبورد الرئيسي
# =========================
keyboard = [
    ["🛠 طلب صيانة"],
    ["🔍 بحث طلبات صيانة"],
    ["🌐 طلب تركيب جديد"],
    ["🔍 بحث طلبات تركيب"],
    ["📦 سحب بضاعة"],
    ["➕ إضافة بضاعة"],
    ["📊 الإحصائيات"],
    ["💰 تحصيل دين"],
]

reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
back_markup = ReplyKeyboardMarkup([["🔙 رجوع"]], resize_keyboard=True)
confirm_markup = ReplyKeyboardMarkup([["✅ تأكيد", "❌ إلغاء"]], resize_keyboard=True)

# =========================
# JSON Functions
# =========================
def load_json(file_name, default):
    if os.path.exists(file_name):
        with open(file_name, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(file_name, data):
    with open(file_name, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def get_inventory():
    return load_json(INVENTORY_FILE, {
        "📡 لايت بيم M5": 0,
        "📶 لايت AC": 0,
        "📟 راوتر نتس": 0,
        "📠 راوتر كودي": 0
    })

def update_inventory(product, quantity_change):
    inventory_data = get_inventory()
    if product in inventory_data:
        inventory_data[product] += quantity_change
        if inventory_data[product] < 0:
            inventory_data[product] = 0
    save_json(INVENTORY_FILE, inventory_data)
    return inventory_data[product]

def check_inventory(product, requested_quantity):
    inventory_data = get_inventory()
    return inventory_data.get(product, 0) >= requested_quantity

def load_statistics():
    global statistics
    loaded = load_json(STATISTICS_FILE, {})
    statistics = {
        "cash": loaded.get("cash", 0),
        "debt": loaded.get("debt", 0),
    }

def save_statistics():
    save_json(STATISTICS_FILE, statistics)

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        ws1 = wb.create_sheet("طلبات الصيانة")
        ws1.append(["التاريخ", "Telegram ID", "اسم المستخدم", "اسم العميل", "الرقم الوطني", "العطل", "الموقع", "الهاتف", "مقدم الطلب", "المكلف", "الحالة"])
        
        ws2 = wb.create_sheet("طلبات التركيب")
        ws2.append(["التاريخ", "Telegram ID", "اسم المستخدم", "اسم العميل", "الرقم الوطني", "نوع التركيب", "الموقع", "الهاتف", "مقدم الطلب", "المكلف بالتركيب", "الحالة"])
        
        ws3 = wb.create_sheet("المخزون")
        ws3.append(["التاريخ", "Telegram ID", "اسم المستخدم", "نوع العملية", "المنتج", "الكمية", "المخزون بعد العملية", "ملاحظات"])
        
        ws4 = wb.create_sheet("البيع المباشر")
        ws4.append(["التاريخ", "Telegram ID", "اسم المستخدم", "المنتج", "الكمية", "سعر القطعة", "الإجمالي", "الموظف", "طريقة الدفع", "نوع البيع"])
        
        ws5 = wb.create_sheet("البيع دين")
        ws5.append(["التاريخ", "Telegram ID", "اسم المستخدم", "المنتج", "الكمية", "سعر القطعة", "الإجمالي", "الموظف", "الدين المتبقي", "ملاحظات"])
        
        ws6 = wb.create_sheet("تحصيل الديون")
        ws6.append(["التاريخ", "Telegram ID", "اسم المستخدم", "المبلغ المحصل", "اسم صاحب الدين", "الموظف المحصل", "الدين المتبقي", "ملاحظات"])
        
        ws7 = wb.create_sheet("الإحصائيات")
        ws7.append(["تاريخ التحديث", "Telegram ID", "اسم المستخدم", "إجمالي المقبوضات", "إجمالي الديون", "مخزون لايت بيم M5", "مخزون لايت AC", "مخزون راوتر نتس", "مخزون راوتر كودي"])
        
        ws8 = wb.create_sheet("سجل الموظفين")
        ws8.append(["التاريخ", "الموظف", "نوع العملية", "المبلغ", "العميل", "ملاحظات"])
        
        wb.save(EXCEL_FILE)
        print("✅ تم إنشاء ملف Excel")

def save_to_excel(sheet_name, row_data):
    try:
        init_excel()
        wb = load_workbook(EXCEL_FILE)
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
        ws = wb[sheet_name]
        ws.append(row_data)
        wb.save(EXCEL_FILE)
    except Exception as e:
        print(f"Excel Error: {e}")

def get_user_name(update):
    user = update.effective_user
    return user.full_name or user.username or str(user.id)

async def show_main_menu(update: Update):
    await update.message.reply_text("أهلا بك 🌹", reply_markup=reply_markup)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_main_menu(update)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_main_menu(update)
    return ConversationHandler.END

async def buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == "🛠 طلب صيانة":
        await update.message.reply_text("اسم العميل:", reply_markup=back_markup)
        return MAINTENANCE_NAME

    elif text == "🔍 بحث طلبات صيانة":
        await update.message.reply_text("أدخل الاسم أو رقم الهاتف:", reply_markup=back_markup)
        return SEARCH_MAINTENANCE
    
    elif text == "🌐 طلب تركيب جديد":
        await update.message.reply_text("اسم العميل:", reply_markup=back_markup)
        return INSTALLATION_NAME
    
    elif text == "🔍 بحث طلبات تركيب":
        await update.message.reply_text("أدخل الاسم أو رقم الهاتف:", reply_markup=back_markup)
        return SEARCH_INSTALLATION

    elif text == "➕ إضافة بضاعة":
        product_keyboard = [["📡 لايت بيم M5"], ["📶 لايت AC"], ["📟 راوتر نتس"], ["📠 راوتر كودي"], ["🔙 رجوع"]]
        await update.message.reply_text("اختر البضاعة:", reply_markup=ReplyKeyboardMarkup(product_keyboard, resize_keyboard=True))
        return ADD_PRODUCT

    elif text == "📦 سحب بضاعة":
        sale_keyboard = [["🔧 بيع تركيب"], ["🛒 بيع مفرق"], ["🏪 بيع جملة"], ["🔙 رجوع"]]
        await update.message.reply_text("اختر نوع البيع:", reply_markup=ReplyKeyboardMarkup(sale_keyboard, resize_keyboard=True))
        return SALE_TYPE

    elif text == "💰 تحصيل دين":
        await update.message.reply_text("💰 أدخل المبلغ:", reply_markup=back_markup)
        return COLLECT_AMOUNT

    elif text == "📊 الإحصائيات":
        inv = get_inventory()
        msg = f"""📊 الإحصائيات

💰 المقبوضات: {statistics.get('cash', 0)}
💳 الديون: {statistics.get('debt', 0)}

📦 المخزون
📡 لايت بيم M5: {inv.get('📡 لايت بيم M5', 0)}
📶 لايت AC: {inv.get('📶 لايت AC', 0)}
📟 راوتر نتس: {inv.get('📟 راوتر نتس', 0)}
📠 راوتر كودي: {inv.get('📠 راوتر كودي', 0)}"""
        await update.message.reply_text(msg, reply_markup=reply_markup)
        return ConversationHandler.END

    elif text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END

    return ConversationHandler.END

# دوال الصيانة
async def maintenance_name(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["name"] = update.message.text
    await update.message.reply_text("الرقم الوطني:", reply_markup=back_markup)
    return MAINTENANCE_NATIONAL

async def maintenance_national(update, context):
    context.user_data["national"] = update.message.text
    await update.message.reply_text("العطل:", reply_markup=back_markup)
    return MAINTENANCE_PROBLEM

async def maintenance_problem(update, context):
    context.user_data["problem"] = update.message.text
    await update.message.reply_text("الموقع:", reply_markup=back_markup)
    return MAINTENANCE_LOCATION

async def maintenance_location(update, context):
    context.user_data["location"] = update.message.text
    await update.message.reply_text("رقم الهاتف:", reply_markup=back_markup)
    return MAINTENANCE_PHONE

async def maintenance_phone(update, context):
    context.user_data["phone"] = update.message.text
    await update.message.reply_text("اسم مقدم الطلب:", reply_markup=back_markup)
    return MAINTENANCE_REQUESTER

async def maintenance_requester(update, context):
    context.user_data["requester"] = update.message.text
    # عرض قائمة الموظفين للمكلف
    kb = [[e] for e in EMPLOYEES]
    await update.message.reply_text("المكلف بالصيانة:", reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
    return MAINTENANCE_ASSIGNEE

async def maintenance_assignee(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["assignee"] = update.message.text
    tid = update.effective_user.id
    summary = f"""📋 ملخص طلب الصيانة:

👤 الاسم: {context.user_data['name']}
🆔 الرقم الوطني: {context.user_data['national']}
🛠 العطل: {context.user_data['problem']}
📍 الموقع: {context.user_data['location']}
📞 الهاتف: {context.user_data['phone']}
🙋 مقدم الطلب: {context.user_data['requester']}
👨‍🔧 المكلف: {context.user_data['assignee']}
🆔 Telegram ID الخاص بالمسجل: {tid}

تأكيد إرسال الطلب؟"""
    await update.message.reply_text(summary, reply_markup=confirm_markup)
    return MAINTENANCE_CONFIRM

async def maintenance_confirm(update, context):
    tid = update.effective_user.id
    if update.message.text == "✅ تأكيد":
        data = {
            "name": context.user_data["name"], 
            "national": context.user_data["national"], 
            "problem": context.user_data["problem"], 
            "location": context.user_data["location"], 
            "phone": context.user_data["phone"], 
            "requester": context.user_data["requester"],
            "assignee": context.user_data.get("assignee", "لم يعين"),
            "telegram_id": tid, 
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"), 
            "status": "pending"
        }
        req = load_json(MAINTENANCE_FILE, [])
        req.append(data)
        save_json(MAINTENANCE_FILE, req)
        save_to_excel("طلبات الصيانة", [data['date'], tid, get_user_name(update), data['name'], data['national'], data['problem'], data['location'], data['phone'], data['requester'], data['assignee'], data['status']])
        await update.message.reply_text(f"✅ تم تسجيل طلب الصيانة\n👨‍🔧 المكلف: {data['assignee']}", reply_markup=reply_markup)
    else:
        await update.message.reply_text("❌ تم الإلغاء", reply_markup=reply_markup)
    return ConversationHandler.END

async def search_maintenance(update, context):
    term = update.message.text.lower()
    reqs = load_json(MAINTENANCE_FILE, [])
    found = [r for r in reqs if term in r["name"].lower() or term in r["phone"]]
    if not found:
        await update.message.reply_text("❌ لا نتائج", reply_markup=reply_markup)
    else:
        text = ""
        for r in found:
            text += f"\n👤 {r['name']}\n🛠 {r['problem']}\n📍 {r['location']}\n📞 {r['phone']}\n👨‍🔧 المكلف: {r.get('assignee', '-')}\n📅 {r['date']}\n📌 {r['status']}\n---"
        await update.message.reply_text(text[:4000], reply_markup=reply_markup)
    return ConversationHandler.END

# دوال التركيب
async def installation_name(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["inst_name"] = update.message.text
    await update.message.reply_text("الرقم الوطني:", reply_markup=back_markup)
    return INSTALLATION_NATIONAL

async def installation_national(update, context):
    context.user_data["inst_national"] = update.message.text
    kb = [["📡 لايت بيم M5"], ["📶 لايت AC"], ["📟 راوتر نتس"], ["📠 راوتر كودي"], ["🔙 رجوع"]]
    await update.message.reply_text("نوع التركيب:", reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
    return INSTALLATION_TYPE

async def installation_type(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["inst_type"] = update.message.text
    await update.message.reply_text("الموقع:", reply_markup=back_markup)
    return INSTALLATION_LOCATION

async def installation_location(update, context):
    context.user_data["inst_location"] = update.message.text
    await update.message.reply_text("رقم الهاتف:", reply_markup=back_markup)
    return INSTALLATION_PHONE

async def installation_phone(update, context):
    context.user_data["inst_phone"] = update.message.text
    await update.message.reply_text("مقدم الطلب:", reply_markup=back_markup)
    return INSTALLATION_REQUESTER

async def installation_requester(update, context):
    context.user_data["inst_requester"] = update.message.text
    kb = [[e] for e in EMPLOYEES]
    await update.message.reply_text("المكلف بالتركيب:", reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
    return INSTALLATION_ASSIGNEE

async def installation_assignee(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["inst_assignee"] = update.message.text
    tid = update.effective_user.id
    summary = f"""📋 ملخص طلب التركيب:

👤 الاسم: {context.user_data['inst_name']}
🔧 النوع: {context.user_data['inst_type']}
📍 الموقع: {context.user_data['inst_location']}
📞 الهاتف: {context.user_data['inst_phone']}
👨‍🔧 المكلف: {context.user_data['inst_assignee']}

تأكيد؟"""
    await update.message.reply_text(summary, reply_markup=confirm_markup)
    return INSTALLATION_CONFIRM

async def installation_confirm(update, context):
    tid = update.effective_user.id
    if update.message.text == "✅ تأكيد":
        data = {"name": context.user_data["inst_name"], "national": context.user_data["inst_national"], "install_type": context.user_data["inst_type"], "location": context.user_data["inst_location"], "phone": context.user_data["inst_phone"], "requester": context.user_data["inst_requester"], "assigned_to": context.user_data.get("inst_assignee", "لم يعين"), "telegram_id": tid, "date": datetime.now().strftime("%Y-%m-%d %H:%M"), "status": "pending"}
        inst = load_json(INSTALLATION_FILE, [])
        inst.append(data)
        save_json(INSTALLATION_FILE, inst)
        save_to_excel("طلبات التركيب", [data['date'], tid, get_user_name(update), data['name'], data['national'], data['install_type'], data['location'], data['phone'], data['requester'], data['assigned_to'], data['status']])
        await update.message.reply_text("✅ تم تسجيل طلب التركيب", reply_markup=reply_markup)
    else:
        await update.message.reply_text("❌ تم الإلغاء", reply_markup=reply_markup)
    return ConversationHandler.END

async def search_installation(update, context):
    term = update.message.text.lower()
    insts = load_json(INSTALLATION_FILE, [])
    found = [i for i in insts if term in i["name"].lower() or term in i["phone"]]
    if not found:
        await update.message.reply_text("❌ لا نتائج", reply_markup=reply_markup)
    else:
        text = ""
        for i in found:
            text += f"\n👤 {i['name']}\n🔧 {i['install_type']}\n📍 {i['location']}\n👨‍🔧 {i.get('assigned_to', '-')}\n📅 {i['date']}\n---"
        await update.message.reply_text(text[:4000], reply_markup=reply_markup)
    return ConversationHandler.END

# دوال البضاعة
async def add_product(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["product"] = update.message.text
    await update.message.reply_text("الكمية:", reply_markup=back_markup)
    return ADD_PRODUCT_QUANTITY

async def add_product_quantity(update, context):
    try:
        qty = int(update.message.text)
        product = context.user_data["product"]
        new_qty = update_inventory(product, qty)
        save_to_excel("المخزون", [datetime.now().strftime("%Y-%m-%d %H:%M"), update.effective_user.id, get_user_name(update), "إضافة", product, f"+{qty}", new_qty, ""])
        await update.message.reply_text(f"✅ تمت الإضافة\n📦 {product}: {new_qty}", reply_markup=reply_markup)
        return ConversationHandler.END
    except:
        await update.message.reply_text("❌ رقم صحيح", reply_markup=back_markup)
        return ADD_PRODUCT_QUANTITY

async def sale_type(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["sale_type"] = update.message.text
    kb = [["📡 لايت بيم M5"], ["📶 لايت AC"], ["📟 راوتر نتس"], ["📠 راوتر كودي"], ["🔙 رجوع"]]
    await update.message.reply_text("المنتج:", reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
    return PRODUCT_TYPE

async def product_type(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    inv = get_inventory()
    if inv.get(update.message.text, 0) <= 0:
        await update.message.reply_text("❌ المخزون فارغ", reply_markup=reply_markup)
        return ConversationHandler.END
    context.user_data["product"] = update.message.text
    await update.message.reply_text(f"الكمية (المتوفرة: {inv.get(update.message.text, 0)}):", reply_markup=back_markup)
    return PRODUCT_QUANTITY

async def product_quantity(update, context):
    try:
        qty = int(update.message.text)
        if not check_inventory(context.user_data["product"], qty):
            await update.message.reply_text("❌ كمية غير متوفرة", reply_markup=back_markup)
            return PRODUCT_QUANTITY
        context.user_data["quantity"] = qty
        await update.message.reply_text("سعر القطعة:", reply_markup=back_markup)
        return PRODUCT_PRICE
    except:
        await update.message.reply_text("❌ رقم صحيح", reply_markup=back_markup)
        return PRODUCT_QUANTITY

async def product_price(update, context):
    try:
        context.user_data["price"] = float(update.message.text)
        kb = [[e] for e in EMPLOYEES]
        await update.message.reply_text("الموظف:", reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
        return PRODUCT_EMPLOYEE
    except:
        await update.message.reply_text("❌ سعر صحيح", reply_markup=back_markup)
        return PRODUCT_PRICE

async def product_employee(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["employee"] = update.message.text
    kb = [["💰 نقداً"], ["📝 دين"], ["🔙 رجوع"]]
    await update.message.reply_text("طريقة الدفع:", reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
    return PRODUCT_PAYMENT

async def product_payment(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    payment = update.message.text
    product = context.user_data["product"]
    qty = context.user_data["quantity"]
    price = context.user_data["price"]
    employee = context.user_data["employee"]
    total = qty * price
    new_qty = update_inventory(product, -qty)
    tid = update.effective_user.id
    
    if payment == "💰 نقداً":
        statistics["cash"] += total
        sheet = "البيع المباشر"
    else:
        statistics["debt"] += total
        sheet = "البيع دين"
    save_statistics()
    
    save_to_excel(sheet, [datetime.now().strftime("%Y-%m-%d %H:%M"), tid, get_user_name(update), product, qty, price, total, employee, payment, context.user_data.get("sale_type", "")])
    save_to_excel("المخزون", [datetime.now().strftime("%Y-%m-%d %H:%M"), tid, get_user_name(update), "سحب", product, f"-{qty}", new_qty, employee])
    
    await update.message.reply_text(f"✅ تمت العملية\n📦 {product}\n🔢 {qty}\n💰 {total}\n👨‍💼 {employee}\n📦 المتبقي: {new_qty}", reply_markup=reply_markup)
    return ConversationHandler.END

# ديون
async def collect_amount(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    try:
        context.user_data["collect_amount"] = float(update.message.text)
        await update.message.reply_text("اسم صاحب الدين:", reply_markup=back_markup)
        return COLLECT_CUSTOMER_NAME
    except:
        await update.message.reply_text("❌ رقم صحيح", reply_markup=back_markup)
        return COLLECT_AMOUNT

async def collect_customer_name(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["collect_customer"] = update.message.text
    kb = [[e] for e in EMPLOYEES]
    await update.message.reply_text("الموظف المحصل:", reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
    return COLLECT_EMPLOYEE

async def collect_employee(update, context):
    if update.message.text == "🔙 رجوع":
        await show_main_menu(update)
        return ConversationHandler.END
    context.user_data["collect_employee"] = update.message.text
    summary = f"""📋 تحصيل دين:
💰 المبلغ: {context.user_data['collect_amount']}
👤 صاحب الدين: {context.user_data['collect_customer']}
👨‍💼 الموظف: {context.user_data['collect_employee']}

تأكيد؟"""
    await update.message.reply_text(summary, reply_markup=confirm_markup)
    return COLLECT_CONFIRM

async def collect_confirm(update, context):
    if update.message.text == "✅ تأكيد":
        amount = context.user_data["collect_amount"]
        if amount > statistics.get("debt", 0):
            await update.message.reply_text(f"❌ الدين أقل من المبلغ", reply_markup=reply_markup)
            return ConversationHandler.END
        statistics["debt"] -= amount
        statistics["cash"] += amount
        save_statistics()
        tid = update.effective_user.id
        save_to_excel("تحصيل الديون", [datetime.now().strftime("%Y-%m-%d %H:%M"), tid, get_user_name(update), amount, context.user_data["collect_customer"], context.user_data["collect_employee"], statistics["debt"], ""])
        await update.message.reply_text(f"✅ تم التحصيل\n💰 المبلغ: {amount}\n💳 الدين المتبقي: {statistics['debt']}", reply_markup=reply_markup)
    else:
        await update.message.reply_text("❌ تم الإلغاء", reply_markup=reply_markup)
    return ConversationHandler.END

async def error_handler(update, context):
    print("ERROR:", context.error)

# =========================
# MAIN
# =========================
def main():
    load_statistics()
    init_excel()
    
    app = ApplicationBuilder().token(TOKEN).build()
    
    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.TEXT & ~filters.COMMAND, buttons)],
        states={
            MAINTENANCE_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_name)],
            MAINTENANCE_NATIONAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_national)],
            MAINTENANCE_PROBLEM: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_problem)],
            MAINTENANCE_LOCATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_location)],
            MAINTENANCE_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_phone)],
            MAINTENANCE_REQUESTER: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_requester)],
            MAINTENANCE_ASSIGNEE: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_assignee)],
            MAINTENANCE_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, maintenance_confirm)],
            SEARCH_MAINTENANCE: [MessageHandler(filters.TEXT & ~filters.COMMAND, search_maintenance)],
            ADD_PRODUCT: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_product)],
            ADD_PRODUCT_QUANTITY: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_product_quantity)],
            SALE_TYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, sale_type)],
            PRODUCT_TYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, product_type)],
            PRODUCT_QUANTITY: [MessageHandler(filters.TEXT & ~filters.COMMAND, product_quantity)],
            PRODUCT_PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, product_price)],
            PRODUCT_EMPLOYEE: [MessageHandler(filters.TEXT & ~filters.COMMAND, product_employee)],
            PRODUCT_PAYMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, product_payment)],
            COLLECT_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_amount)],
            COLLECT_CUSTOMER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_customer_name)],
            COLLECT_EMPLOYEE: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_employee)],
            COLLECT_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_confirm)],
            INSTALLATION_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_name)],
            INSTALLATION_NATIONAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_national)],
            INSTALLATION_TYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_type)],
            INSTALLATION_LOCATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_location)],
            INSTALLATION_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_phone)],
            INSTALLATION_REQUESTER: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_requester)],
            INSTALLATION_ASSIGNEE: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_assignee)],
            INSTALLATION_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, installation_confirm)],
            SEARCH_INSTALLATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, search_installation)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)
    app.add_error_handler(error_handler)
    
    print("✅ BOT RUNNING...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
