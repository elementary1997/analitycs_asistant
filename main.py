from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters
from handlers import start, handle_file, set_lastname, set_main, merge_to_main, get_main, merge_cmd, button_handler
from config import TOKEN

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("setlastname", set_lastname))
    app.add_handler(CommandHandler("setmain", set_main))
    app.add_handler(CommandHandler("getmain", get_main))
    app.add_handler(CommandHandler("merge", merge_cmd))
    # Handle "Добавить отчёт" button presses
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL & ~filters.COMMAND, handle_file))
    # For merging, user can send a document with caption '/merge' to trigger merging
    app.add_handler(MessageHandler(filters.Document.ALL & filters.CaptionRegex(r"^/merge\b"), merge_to_main))
    print("Бот запущен...")
    app.run_polling()

if __name__ == '__main__':
    main()