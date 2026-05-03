import logging
import os
import tempfile
from pathlib import Path

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
)

from process_supplier import detect_supplier, process_file

logging.basicConfig(
    format="%(asctime)s  %(levelname)s  %(name)s  %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ["BOT_TOKEN"]
ALLOWED_IDS = {
    int(x.strip())
    for x in os.environ.get("ALLOWED_USER_IDS", "").split(",")
    if x.strip().isdigit()
}


def _is_allowed(update: Update) -> bool:
    return update.effective_user.id in ALLOWED_IDS


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not _is_allowed(update):
        await update.message.reply_text("⛔ У вас нет доступа к этому боту.")
        return
    await update.message.reply_text(
        "Привет! Я обрабатываю прайс-листы AVRORA Steel.\n\n"
        "Отправьте .xlsx файл от поставщика (Руфат или Карагандинский) — "
        "получите два файла с наценкой: розница и опт."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not _is_allowed(update):
        await update.message.reply_text("⛔ У вас нет доступа к этому боту.")
        return

    doc = update.message.document
    fname = doc.file_name or ""

    if not fname.lower().endswith((".xlsx", ".xls")):
        await update.message.reply_text(
            "Пожалуйста, отправьте файл формата .xlsx"
        )
        return

    supplier = detect_supplier(fname)
    if supplier is None:
        await update.message.reply_text(
            "❓ Не удалось определить поставщика по имени файла.\n\n"
            "Откройте проект «AVRORA прайсы» в Claude, прикрепите этот файл "
            "и напишите: «новый поставщик, добавь конфиг»."
        )
        return

    await update.message.reply_text(
        f"⏳ Обрабатываю прайс поставщика «{supplier['display_name']}»..."
    )

    tg_file = await doc.get_file()
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()

    try:
        await tg_file.download_to_drive(tmp_path)
        retail_bytes, wholesale_bytes = process_file(tmp_path, supplier)

        stem = Path(fname).stem
        await update.message.reply_document(
            document=retail_bytes,
            filename=f"{stem}_розница.xlsx",
            caption="✅ Розница",
        )
        await update.message.reply_document(
            document=wholesale_bytes,
            filename=f"{stem}_опт.xlsx",
            caption="✅ Опт",
        )

    except Exception as exc:
        logger.error("Processing error for %s: %s", fname, exc, exc_info=True)
        await update.message.reply_text(
            f"❌ Ошибка при обработке файла.\n"
            f"Откройте проект «AVRORA прайсы» в Claude и сообщите:\n"
            f"`{type(exc).__name__}: {exc}`",
            parse_mode="Markdown",
        )
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    logger.info("Bot started. Allowed user IDs: %s", ALLOWED_IDS)
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
