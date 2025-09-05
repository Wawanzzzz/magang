import logging
import sys
from pathlib import Path
from config import IMAGE_PATHS, INPUT_DIR, OUTPUT_DIR, LOGS_DIR

def validate_setup():
    logger = logging.getLogger(__name__)
    try:
        # Validasi Python
        if sys.version_info < (3, 7):
            logger.error(f"Python version terlalu lama: {sys.version}")
            return False

        # Validasi dependencies
        if not validate_dependencies():
            return False

        # Validasi direktori
        if not validate_directories():
            return False

        # Validasi image paths (tidak fatal)
        validate_image_paths()  # hanya warning, tidak menghentikan setup

        logger.info("✅ Semua validasi berhasil!")
        return True
    except Exception as e:
        logger.error(f"Error dalam validasi setup: {str(e)}")
        return False

def validate_dependencies():
    logger = logging.getLogger(__name__)
    required_packages = ['easyocr', 'cv2', 'PIL', 'numpy']
    missing_packages = []
    for package in required_packages:
        try:
            if package == 'cv2': import cv2
            elif package == 'PIL': from PIL import Image
            elif package == 'easyocr': import easyocr
            elif package == 'numpy': import numpy
            logger.info(f"✅ {package} tersedia")
        except ImportError:
            missing_packages.append(package)
            logger.error(f"❌ {package} tidak ditemukan")
    if missing_packages:
        logger.error("Install dependencies yang hilang dengan: pip install -r requirements.txt")
        return False
    return True

def validate_directories():
    logger = logging.getLogger(__name__)
    directories = [("Input", INPUT_DIR), ("Output", OUTPUT_DIR), ("Logs", LOGS_DIR)]
    for name, directory in directories:
        try:
            directory.mkdir(parents=True, exist_ok=True)
            logger.info(f"✅ Direktori {name}: {directory}")
        except Exception as e:
            logger.error(f"❌ Error direktori {name}: {str(e)}")
            return False
    return True

def validate_image_paths():
    """Validasi path gambar tapi tidak menghentikan setup"""
    logger = logging.getLogger(__name__)
    if not IMAGE_PATHS:
        logger.warning("⚠️ IMAGE_PATHS kosong. Tidak ada gambar default.")
        return True  # tetap valid

    valid_paths = 0
    for i, image_path in enumerate(IMAGE_PATHS, 1):
        path = Path(image_path)
        if path.exists():
            logger.info(f"✅ [{i}] {path.name} - ditemukan")
            valid_paths += 1
        else:
            logger.warning(f"⚠️  [{i}] {path.name} - tidak ditemukan di: {path}")
    logger.info(f"📊 {valid_paths}/{len(IMAGE_PATHS)} gambar ditemukan")
    return True
