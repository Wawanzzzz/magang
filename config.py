"""
Konfigurasi aplikasi EasyOCR
"""
from pathlib import Path

# ----- Path dasar -----
BASE_DIR = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"
INPUT_DIR = ASSETS_DIR / "input"
OUTPUT_DIR = ASSETS_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"
UPLOAD_DIR = BASE_DIR / "uploads"  # Untuk upload file via web

# ----- Buat direktori jika belum ada -----
for d in [INPUT_DIR, OUTPUT_DIR, LOGS_DIR, UPLOAD_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ----- Path gambar default (opsional) -----
# Jika tidak ada gambar default, IMAGE_PATHS bisa dikosongkan.
# Bisa juga menambahkan contoh gambar di INPUT_DIR.
IMAGE_PATHS = [
    # str(INPUT_DIR / "contoh1.jpg"),
    # str(INPUT_DIR / "contoh2.png"),
]

# ----- Konfigurasi OCR -----
OCR_CONFIG = {
    'languages': ['id', 'en'],      # Indonesian & English
    'gpu': False,                    # True jika ada GPU CUDA
    'detail': 1,                     # 0=text only, 1=text+confidence
    'paragraph': False,
    'width_ths': 0.7,
    'height_ths': 0.7,
    'enhancement_level': 'aggressive',   # Untuk gambar buruk
    'confidence_threshold': 0.2,
    'enable_text_correction': True,
    'enable_word_dictionary': True
}

# ----- Format file -----
SUPPORTED_FORMATS = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp']

# ----- Flask upload -----
ALLOWED_EXTENSIONS = {ext.strip('.') for ext in SUPPORTED_FORMATS}  # {'jpg','png',...}

# ----- Logging -----
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
LOG_LEVEL = 'INFO'
