# app.py
import os
import logging
import tempfile
import shutil
import img2pdf
from io import BytesIO
from PIL import Image
import qrcode
import requests
import asyncio

from flask import Flask, request, Response

import pytesseract
import cv2
import numpy as np

from pdf2docx import Converter
from telegram import Update, Bot, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, filters

# =================== CONFIG ===================
TOKEN = "7797976277:AAGeRUw7sqMh_PQrPNsISTHs_9cSrXyzFiQ"
# WEBHOOK_URL should be like: https://your-app.onrender.com
WEBHOOK_URL = os.environ.get("WEBHOOK_URL")  # Render-da SERVICE URL ni shu env ga qo'ying

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Flask app
app = Flask(__name__)

# Telegram Application (async)
application = Application.builder().token(TOKEN).build()
bot = application.bot  # convenience

# ================ HELPERS ====================
def safe_filename(name: str) -> str:
    return "".join(c if c.isalnum() or c in "._-" else "_" for c in name)

def pil_image_from_bytes(data: bytes) -> Image.Image:
    return Image.open(BytesIO(data))

async def set_telegram_webhook():
    """
    Avtomatik webhook o'rnatish. WEBHOOK_URL muhit o'zgaruvchisi bo'lmasa, pass.
    """
    if not WEBHOOK_URL:
        logger.warning("WEBHOOK_URL env not set — webhook not configured automatically.")
        return False

    webhook_url = WEBHOOK_URL.rstrip("/") + f"/{TOKEN}"
    resp = requests.post(f"https://api.telegram.org/bot{TOKEN}/setWebhook", json={"url": webhook_url})
    try:
        j = resp.json()
    except Exception:
        logger.error("Webhook set request failed, non-json response.")
        return False

    if j.get("ok"):
        logger.info(f"Webhook set to {webhook_url}")
        return True
    else:
        logger.error(f"Failed to set webhook: {j}")
        return False

# ================ HANDLERS ===================
async def start(update: Update, context):
    name = update.effective_user.first_name or "Foydalanuvchi"
    text = (
        f"Assalomu alaykum, {name} 👋\n\n"
        "Quyidagi funksiyalar mavjud:\n"
        "- /ocr — rasm yoki PDF dan matn chiqarish (pytesseract)\n"
        "- /jpg2pdf — rasm yuboring, PDF ga aylantiraman\n"
        "- /pdfsplit — PDF yuboring, keyin sahifa raqamlarini yuboring (mas: 1,3-5)\n"
        "- /qrgen — matn yuboring, QR kod hosil qilaman\n"
        "- /qrscan — QR kod rasm yuboring, ichidagi matnni olaman\n"
        "- /compress — rasm yuboring, siqilgan rasm qaytariladi\n"
        "- /kiril2lotin yoki /lotin2kiril — matn almashinuvi\n\n"
        "Bot webhook orqali ishlaydi."
    )
    await update.message.reply_text(text)

# OCR: rasm yoki PDF (faqat rasm sahifalari uchun)
async def ocr_handler(update: Update, context):
    msg = update.message
    if msg.document:  # PDF yoki rasm fayl
        file = await msg.document.get_file()
        fname = msg.document.file_name or "file"
        ext = os.path.splitext(fname)[1].lower()
        with tempfile.TemporaryDirectory() as td:
            local = os.path.join(td, safe_filename(fname))
            await file.download_to_drive(local)
            texts = []
            if ext == ".pdf":
                # PDF sahifalarini rasmga aylantirish uchun pdf2image not included by default.
                # Asosiy maqsad pdf sahifalarida rasm bo'lsa, pdf2docx bilan matn olish mumkin,
                # lekin eng ishonchli usul - pdf sahifalarini rasmga konvert qilish (pdf2image).
                # Agar pdf2image o'rnatilmagan bo'lsa, xatolik qaytamiz.
                try:
                    from pdf2image import convert_from_path
                except Exception:
                    await msg.reply_text("PDF sahifalarini rasmga aylantirish uchun `pdf2image` kerak. Serverda yo'q.")
                    return

                pages = convert_from_path(local, dpi=200)
                for p in pages:
                    txt = pytesseract.image_to_string(p, lang='eng+rus+ukr+uz' )
                    texts.append(txt)
            else:
                # rasm fayli
                img = Image.open(local)
                txt = pytesseract.image_to_string(img, lang='eng+rus+ukr+uz')
                texts.append(txt)

            final = "\n\n".join(t for t in texts if t.strip())
            if not final.strip():
                final = "Matn topilmadi yoki rasm aniq emas."
            # Telegram limit: 4096 char per message
            for i in range(0, len(final), 4000):
                await msg.reply_text(final[i:i+4000])
    elif msg.photo:
        photo = msg.photo[-1]
        f = await photo.get_file()
        bio = BytesIO()
        await f.download_to_memory(out=bio)
        bio.seek(0)
        img = Image.open(bio)
        txt = pytesseract.image_to_string(img, lang='eng+rus+ukr+uz')
        if not txt.strip():
            txt = "Matn topilmadi yoki rasm aniq emas."
        for i in range(0, len(txt), 4000):
            await msg.reply_text(txt[i:i+4000])
    else:
        await msg.reply_text("Iltimos rasm yoki PDF yuboring (jpg/png/pdf).")

# JPG/PNG -> PDF
async def jpg2pdf_handler(update: Update, context):
    msg = update.message
    # rasmni olamiz (photo yoki document)
    img_bytes = None
    name = "images"
    if msg.photo:
        f = await msg.photo[-1].get_file()
        bio = BytesIO()
        await f.download_to_memory(out=bio)
        img_bytes = bio.getvalue()
    elif msg.document:
        f = await msg.document.get_file()
        bio = BytesIO()
        await f.download_to_memory(out=bio)
        img_bytes = bio.getvalue()
        name = msg.document.file_name or name
    else:
        await msg.reply_text("Iltimos, JPG yoki PNG rasm yuboring.")
        return

    # Agar bir nechta rasm ketma-ket yuboriladi: bu handler bitta rasmga
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, safe_filename(name))
            with open(in_path, "wb") as f:
                f.write(img_bytes)
            # Agar rasm bo'lsa, img2pdf ga bering
            pdf_bytes = img2pdf.convert([in_path])
            await msg.reply_document(document=BytesIO(pdf_bytes), filename=f"{os.path.splitext(name)[0]}.pdf",
                                     caption="✅ Rasm PDF ga aylantirildi")
    except Exception as e:
        logger.exception(e)
        await msg.reply_text(f"Xatolik yuz berdi: {str(e)}")

# PDF split (foydalanuvchi avval PDF yuboradi, keyin sahifalar)
# For simplicity, we use a simple state in memory per user (not persistent)
USER_STATE = {}  # user_id -> {'mode': 'pdfsplit', 'path': '/tmp/...', 'total_pages': n}

async def pdfsplit_start(update: Update, context):
    msg = update.message
    if not msg.document:
        await msg.reply_text("Iltimos, PDF yuboring (document sifatida).")
        return
    doc = msg.document
    if not doc.file_name.lower().endswith(".pdf"):
        await msg.reply_text("Iltimos .pdf formatda yuboring.")
        return
    if doc.file_size and doc.file_size > 20*1024*1024:
        await msg.reply_text("Iltimos 20MB dan kichik PDF yuboring.")
        return
    f = await doc.get_file()
    with tempfile.TemporaryDirectory() as td:
        local = os.path.join(td, safe_filename(doc.file_name))
        await f.download_to_drive(local)
        # sahifalar soni:
        try:
            import fitz  # PyMuPDF
        except Exception:
            await msg.reply_text("Serverda PyMuPDF (fitz) o'rnatilmagan.")
            return
        pdf = fitz.open(local)
        total = pdf.page_count
        pdf.close()
        # saqlaymiz temp faylni USER_STATE-da (foydalanuvchi uchun)
        user_id = update.effective_user.id
        saved_path = os.path.join(tempfile.gettempdir(), f"pdfsplit_{user_id}.pdf")
        shutil.copy(local, saved_path)
        USER_STATE[user_id] = {'mode': 'pdfsplit', 'path': saved_path, 'total_pages': total}
        await msg.reply_text(f"PDF saqlandi. Jami sahifa: {total}\nEndi ajratmoqchi bo'lgan sahifalarni kiriting (mas: 1,3-5).")

async def pdfsplit_pages(update: Update, context):
    user_id = update.effective_user.id
    msg = update.message
    st = USER_STATE.get(user_id)
    if not st or st.get('mode') != 'pdfsplit':
        return  # This message not for pdfsplit
    text = msg.text.strip()
    try:
        pages = parse_page_numbers(text, st['total_pages'])
    except Exception as e:
        await msg.reply_text(f"Sahifa raqamlarini tahlil qilishda xato: {e}")
        return
    try:
        import fitz
        input_pdf = fitz.open(st['path'])
        new_pdf = fitz.open()
        for p in pages:
            new_pdf.insert_pdf(input_pdf, from_page=p-1, to_page=p-1)
        out_path = st['path'].replace(".pdf", f"_split_{user_id}.pdf")
        new_pdf.save(out_path)
        new_pdf.close()
        input_pdf.close()
        # yuborish
        with open(out_path, "rb") as f:
            await msg.reply_document(document=f, filename=f"split_{os.path.basename(st['path'])}",
                                     caption=f"✅ Ajratildi: {text}")
    except Exception as e:
        logger.exception(e)
        await msg.reply_text(f"Xatolik: {e}")
    finally:
        # tozalash
        try:
            if os.path.exists(st['path']):
                os.unlink(st['path'])
        except:
            pass
        USER_STATE.pop(user_id, None)

def parse_page_numbers(page_input: str, total_pages: int):
    pages = set()
    for part in page_input.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a = int(a); b = int(b)
            if a < 1 or b > total_pages or a > b:
                raise ValueError(f"Noto'g'ri diapazon: {part}")
            pages.update(range(a, b+1))
        else:
            p = int(part)
            if p < 1 or p > total_pages:
                raise ValueError(f"Noto'g'ri sahifa: {p}")
            pages.add(p)
    return sorted(pages)

# QR generate
async def qrgen_handler(update: Update, context):
    msg = update.message
    text = None
    if msg.text and msg.text.startswith("/qrgen"):
        # if user used command with argument: /qrgen https://...
        parts = msg.text.split(" ", 1)
        if len(parts) == 2:
            text = parts[1].strip()
    if not text:
        if msg.reply_to_message and msg.reply_to_message.text:
            text = msg.reply_to_message.text
        else:
            await msg.reply_text("Matn yuboring yoki /qrgen <matn> tarzida ishlating.")
            return
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L)
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image()
    bio = BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    await msg.reply_photo(photo=bio, caption=f"QR yaratildi. Matn: {text[:200]}")

# QR scan (OpenCV)
async def qrscan_handler(update: Update, context):
    msg = update.message
    img_bytes = None
    if msg.photo:
        f = await msg.photo[-1].get_file()
        bio = BytesIO()
        await f.download_to_memory(out=bio)
        img_bytes = bio.getvalue()
    elif msg.document:
        f = await msg.document.get_file()
        bio = BytesIO()
        await f.download_to_memory(out=bio)
        img_bytes = bio.getvalue()
    else:
        await msg.reply_text("Iltimos, QR kod rasm yuboring.")
        return

    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_GRAYSCALE)
    detector = cv2.QRCodeDetector()
    data, points, _ = detector.detectAndDecode(img)
    if data:
        await msg.reply_text(f"QR tarkibi:\n{data}")
    else:
        await msg.reply_text("QR topilmadi yoki o'qib bo'lmadi.")

# Compress image (basic quality reduction)
async def compress_handler(update: Update, context):
    msg = update.message
    img_bytes = None
    name = "image.jpg"
    if msg.photo:
        f = await msg.photo[-1].get_file()
        bio = BytesIO()
        await f.download_to_memory(out=bio)
        img_bytes = bio.getvalue()
    elif msg.document:
        f = await msg.document.get_file()
        bio = BytesIO()
        await f.download_to_memory(out=bio)
        img_bytes = bio.getvalue()
        name = msg.document.file_name or name
    else:
        await msg.reply_text("Iltimos, rasm yuboring (jpg/png).")
        return

    try:
        img = Image.open(BytesIO(img_bytes))
        # convert to RGB if needed
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        out = BytesIO()
        # save with lowered quality (70)
        img.save(out, format="JPEG", quality=70, optimize=True)
        out.seek(0)
        await msg.reply_document(document=out, filename=f"compressed_{os.path.splitext(name)[0]}.jpg",
                                 caption="✅ Rasm siqildi (sifat 70).")
    except Exception as e:
        logger.exception(e)
        await msg.reply_text(f"Xatolik: {e}")

# Kiril <-> Lotin translate (simple mapping)
CYRILLIC_TO_LATIN = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo',
    'ж': 'j', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'x', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
    'ъ': "'", 'ы': 'i', 'ь': "'", 'э': 'e', 'ю': 'yu', 'я': 'ya',
    'ў': "o'", 'қ': 'q', 'ғ': "g'", 'ҳ': 'h',
    'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'Yo',
    'Ж': 'J', 'З': 'Z', 'И': 'I', 'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M',
    'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U',
    'Ф': 'F', 'Х': 'X', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Shch',
    'Ъ': "'", 'Ы': 'I', 'Ь': "'", 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya',
    'Ў': "O'", 'Қ': 'Q', 'Ғ': "G'", 'Ҳ': 'H'
}
LATIN_TO_CYRILLIC = {v: k for k, v in CYRILLIC_TO_LATIN.items() if len(v) == 1}  # basic inverse for single-char mapping

async def kiril2lotin(update: Update, context):
    text = None
    if update.message.text:
        parts = update.message.text.split(" ", 1)
        text = parts[1] if len(parts) > 1 else None
    if not text and update.message.reply_to_message and update.message.reply_to_message.text:
        text = update.message.reply_to_message.text
    if not text:
        await update.message.reply_text("Matn yuboring: /kiril2lotin <matn> yoki reply orqali.")
        return
    out = "".join(CYRILLIC_TO_LATIN.get(ch, ch) for ch in text)
    await update.message.reply_text(out)

async def lotin2kiril(update: Update, context):
    text = None
    if update.message.text:
        parts = update.message.text.split(" ", 1)
        text = parts[1] if len(parts) > 1 else None
    if not text and update.message.reply_to_message and update.message.reply_to_message.text:
        text = update.message.reply_to_message.text
    if not text:
        await update.message.reply_text("Matn yuboring: /lotin2kiril <matn> yoki reply orqali.")
        return
    # naive conversion for single letters only
    out = "".join(LATIN_TO_CYRILLIC.get(ch, ch) for ch in text)
    await update.message.reply_text(out)

# Fallback echo
async def echo(update: Update, context):
    await update.message.reply_text("Buyruqni tanlang yoki /start yozing.")

# ================ REGISTER HANDLERS ===================
application.add_handler(CommandHandler("start", start))
application.add_handler(CommandHandler("ocr", ocr_handler))
application.add_handler(CommandHandler("jpg2pdf", jpg2pdf_handler))
application.add_handler(CommandHandler("pdfsplit", pdfsplit_start))
# pdfsplit_pages uses text messages when in state
application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), pdfsplit_pages))
application.add_handler(CommandHandler("qrgen", qrgen_handler))
application.add_handler(CommandHandler("qrscan", qrscan_handler))
application.add_handler(CommandHandler("compress", compress_handler))
application.add_handler(CommandHandler("kiril2lotin", kiril2lotin))
application.add_handler(CommandHandler("lotin2kiril", lotin2kiril))
# generic handlers for photo/document for OCR, JPG->PDF, QR scan and compress (we map by current command usage)
application.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, ocr_handler))
# fallback
application.add_handler(MessageHandler(filters.ALL, echo))

# ================ FLASK ROUTES ===================
@app.route(f"/{TOKEN}", methods=["POST"])
def telegram_webhook():
    """Telegram webhook entrypoint. Telegram will POST updates here."""
    try:
        update = Update.de_json(request.get_json(force=True), bot)
        # process update asynchronously using application
        asyncio.run(application.process_update(update))
        return Response("OK", status=200)
    except Exception as e:
        logger.exception(e)
        return Response("OK", status=200)

@app.route("/")
def index():
    return "Telegram bot running (webhook)."

# ================ STARTUP ===================
# When module imported, try to set webhook
# NOTE: In Render, the app is imported; perform webhook set once.
if __name__ == "__main__":
    # local run
    logger.info("Starting locally with Flask dev server...")
    asyncio.run(set_telegram_webhook())
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
else:
    # When imported by gunicorn: attempt to set webhook in background.
    try:
        # attempt once (non-blocking)
        loop = asyncio.get_event_loop()
        if loop.is_running():
            # running under an async loop (rare), schedule coroutine
            asyncio.ensure_future(set_telegram_webhook())
        else:
            loop.run_until_complete(set_telegram_webhook())
    except Exception as e:
        logger.warning(f"Webhook setup skipped or failed: {e}")
