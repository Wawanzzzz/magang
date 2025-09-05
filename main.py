import logging
import sys
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from werkzeug.utils import secure_filename

from config import LOG_FORMAT, LOG_LEVEL, LOGS_DIR, ALLOWED_EXTENSIONS, UPLOAD_DIR
from src.ocr_processor import OCRProcessor
from utils.validation import validate_setup

# --- Excel ---
from utils.excel_utils import save_to_excel_structured   # 🔹 gunakan fungsi baru

# ----- Setup Flask -----
app = Flask(__name__)
app.secret_key = "supersecretkey"

UPLOAD_FOLDER = UPLOAD_DIR
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Folder hasil OCR
OUTPUT_FOLDER = Path("assets/output")
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# ----- Logging -----
def setup_logging():
    log_file = LOGS_DIR / f"ocr_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL),
        format=LOG_FORMAT,
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()

# ----- Utility -----
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_ocr_result(file_path: Path):
    """Membaca file .txt hasil OCR"""
    if not file_path.exists():
        return None
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()

# ----- Route untuk serve file upload -----
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

# ----- Route download Excel -----
@app.route("/download_excel")
def download_excel():
    excel_file = OUTPUT_FOLDER / "ocr_results.xlsx"
    if excel_file.exists():
        return send_from_directory(excel_file.parent, excel_file.name, as_attachment=True)
    else:
        flash("File Excel belum tersedia.")
        return redirect(url_for("index"))

# ----- Routes utama -----
@app.route("/", methods=["GET", "POST"])
def index():
    result_text = None
    result_filename = None

    if request.method == "POST":
        if "image" not in request.files:
            flash("Tidak ada file yang dipilih!")
            return redirect(request.url)

        file = request.files["image"]
        if file.filename == "":
            flash("Nama file kosong!")
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = app.config['UPLOAD_FOLDER'] / filename
            file.save(file_path)
            logger.info(f"File diupload: {file_path}")

            try:
                if not validate_setup():
                    flash("Setup tidak valid. Cek konfigurasi!")
                    return redirect(request.url)

                # Jalankan OCR
                ocr = OCRProcessor()
                ocr.process_image(str(file_path))
                logger.info(f"OCR selesai untuk {filename}")

                # Ambil file detail terbaru
                detail_files = list(OUTPUT_FOLDER.glob("*_detail.txt"))
                if detail_files:
                    latest_detail = max(detail_files, key=lambda f: f.stat().st_mtime)
                    result_filename = latest_detail.name

                    # Baca isi file detail untuk ditampilkan
                    result_text = read_ocr_result(latest_detail)

                    # Simpan ke Excel dalam format terstruktur
                    save_to_excel_structured(latest_detail, result_filename)
                else:
                    result_text = "Hasil OCR tidak ditemukan."
                    result_filename = ""

            except Exception as e:
                logger.error(f"Error saat OCR: {str(e)}")
                flash(f"Terjadi error saat OCR: {str(e)}")
                return redirect(request.url)
        else:
            flash("Format file tidak didukung!")
            return redirect(request.url)

    return render_template("index.html", result=result_text, filename=result_filename)


if __name__ == "__main__":
    print("Memulai EasyOCR Web App...")
    app.run(debug=True)
