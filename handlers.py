from telegram import Update, InputFile, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import ContextTypes
from processor import process_report, update_master_workbook
from config import TEMP_INPUT, TEMP_OUTPUT
from users import (
    get_or_create_user_from_telegram,
    update_user_second_name,
    build_report_filename,
)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = get_or_create_user_from_telegram(update.effective_user)
    display_name = user.first_name or user.username or "друг"
    greeting = f"Привет, {display_name}!" if user.first_name or user.username else "Привет!"
    was_new = user.first_name is None and user.username is None and user.second_name is None
    text = [greeting]
    if was_new:
        text.append("Я тебя зарегистрировал.")
    text.append("Отправь файл .xlsx, и я обработаю его.")
    text.append("Команды: /setlastname Иванов — задать фамилию; /setmain filename.xlsx — задать основной файл отчёта")
    kb = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="Добавить отчёт")]],
        resize_keyboard=True
    )
    await update.message.reply_text(" ".join(text), reply_markup=kb)

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = update.message.document
    if not file.file_name.endswith('.xlsx'):
        await update.message.reply_text("Пожалуйста, отправь файл в формате .xlsx")
        return
    # Авторизуем/создадим пользователя при получении файла
    user = get_or_create_user_from_telegram(update.effective_user)
    # Если ранее пользователь вызвал /merge — объединяем в основной файл и отправляем его
    if context.user_data.get('awaiting_merge'):
        await update.message.reply_text("Получил файл, обновляю основной отчёт...")
        file_obj = await file.get_file()
        await file_obj.download_to_drive(TEMP_INPUT)
        from users import load_user_by_id
        u = load_user_by_id(update.effective_user.id)
        if not u or not u.master_filename:
            await update.message.reply_text("Сначала укажи основной файл через /setmain <имя.xlsx>")
            context.user_data['awaiting_merge'] = False
            return
        month_sheet = update_master_workbook(u.master_filename, TEMP_INPUT)
        try:
            with open(u.master_filename, 'rb') as f:
                await update.message.reply_document(InputFile(f, filename=u.master_filename))
        except FileNotFoundError:
            await update.message.reply_text("Основной файл не найден, повтори /setmain и /merge.")
        await update.message.reply_text(f"Готово. Обновлён лист: {month_sheet}")
        context.user_data['awaiting_merge'] = False
        return

    await update.message.reply_text("Получил файл, начинаю обработку...")
    file_obj = await file.get_file()
    await file_obj.download_to_drive(TEMP_INPUT)
    process_report(TEMP_INPUT, TEMP_OUTPUT)
    personalized_name = build_report_filename(user.second_name)
    with open(TEMP_OUTPUT, 'rb') as out_file:
        await update.message.reply_document(
            document=InputFile(out_file, filename=personalized_name)
        )
    await update.message.reply_text("Обработка завершена!")

async def set_lastname(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args if hasattr(context, 'args') else []
    if not args:
        await update.message.reply_text("Использование: /setlastname Иванов")
        return
    second_name = " ".join(args).strip()
    updated = update_user_second_name(update.effective_user.id, second_name)
    await update.message.reply_text(f"Фамилия сохранена: {updated.second_name}")

async def set_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from users import set_master_filename
    args = context.args if hasattr(context, 'args') else []
    if not args:
        await update.message.reply_text("Использование: /setmain finance_report_Иванов.xlsx")
        return
    filename = " ".join(args).strip()
    user = set_master_filename(update.effective_user.id, filename)
    # Создать пустой основной файл с листом 'Общая информация', если его нет
    import os
    if not os.path.exists(user.master_filename):
        import pandas as pd
        with pd.ExcelWriter(user.master_filename, engine='xlsxwriter') as writer:
            pd.DataFrame(columns=[
                'Месяц', 'Количество операций', 'Итого расходы (без "На инвестиции")'
            ]).to_excel(writer, sheet_name='Общая информация', index=False)
    await update.message.reply_text(f"Основной файл установлен: {user.master_filename}")

async def merge_to_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Accept a document and merge its content into user's master workbook by month
    if not update.message or not update.message.document:
        return
    from users import load_user_by_id
    user = load_user_by_id(update.effective_user.id)
    if not user or not user.master_filename:
        await update.message.reply_text("Сначала задай основной файл командой /setmain <имя_файла.xlsx>")
        return
    file = update.message.document
    if not file.file_name.endswith('.xlsx'):
        await update.message.reply_text("Пожалуйста, отправь файл в формате .xlsx")
        return
    await update.message.reply_text("Обновляю основной файл отчёта...")
    file_obj = await file.get_file()
    await file_obj.download_to_drive(TEMP_INPUT)
    month_sheet = update_master_workbook(user.master_filename, TEMP_INPUT)
    # Отправить актуальный основной файл пользователю
    try:
        with open(user.master_filename, 'rb') as f:
            await update.message.reply_document(InputFile(f, filename=user.master_filename))
    except FileNotFoundError:
        await update.message.reply_text("Основной файл не найден после обновления.")
    await update.message.reply_text(f"Готово. Обновлён лист: {month_sheet}")

async def merge_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Пользователь может нажать кнопку "Добавить отчёт" или вызвать команду /merge,
    # затем прислать файл без подписи
    context.user_data['awaiting_merge'] = True
    await update.message.reply_text("Пришлите .xlsx файл для объединения с основным отчётом.")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Кнопка "Добавить отчёт"
    if update.message and update.message.text and update.message.text.strip().lower() == 'добавить отчёт':
        context.user_data['awaiting_merge'] = True
        await update.message.reply_text("Пришлите .xlsx файл для объединения с основным отчётом.")

async def get_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from users import load_user_by_id
    user = load_user_by_id(update.effective_user.id)
    if not user or not user.master_filename:
        await update.message.reply_text("Основной файл не задан. Используй /setmain <имя_файла.xlsx>")
        return
    try:
        with open(user.master_filename, 'rb') as f:
            await update.message.reply_document(InputFile(f, filename=user.master_filename))
    except FileNotFoundError:
        await update.message.reply_text("Основной файл пока не создан. Добавь первый отчёт через /merge.")
