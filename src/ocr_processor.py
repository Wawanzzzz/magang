"""
OCR Processor - Logic utama untuk processing OCR
"""
import logging
import easyocr
from datetime import datetime
from pathlib import Path

from config import OCR_CONFIG, OUTPUT_DIR
from .image_handler import ImageHandler
from .text_processor import TextProcessor

def extract_fields(text):
    """
    Ekstrak field KTP dari hasil OCR
    """
    fields = {
        "nik": "",
        "nama": "",
        "tempat_tanggal_lahir": "",
        "jenis_kelamin": "",
        "alamat": "",
        "rt_rw": "",
        "kel_desa": "",
        "kecamatan": "",
        "agama": "",
        "status_perkawinan": "",
        "pekerjaan": "",
        "kewarganegaraan": "WNI",
        "berlaku_hingga": "Seumur Hidup"   # default selalu Seumur Hidup
    }

    for line in text.split("\n"):
        line = line.strip().upper()
        if line.startswith("NIK"):
            fields["nik"] = line.replace("NIK", "").strip()
        elif line.startswith("NAMA"):
            fields["nama"] = line.replace("NAMA", "").strip()
        elif "LAHIR" in line:
            fields["tempat_tanggal_lahir"] = line
        elif "LAKI-LAKI" in line or "PEREMPUAN" in line:
            fields["jenis_kelamin"] = line
        elif line.startswith("ALAMAT"):
            fields["alamat"] = line.replace("ALAMAT", "").strip()
        elif "AGAMA" in line:
            fields["agama"] = line.replace("AGAMA", "").strip()
        elif "KAWIN" in line:
            fields["status_perkawinan"] = line
        elif "PEKERJAAN" in line:
            fields["pekerjaan"] = line.replace("PEKERJAAN", "").strip()
        elif "WARGA" in line and "NEGARA" in line:
            val = line.replace("WARGA NEGARA", "").strip()
            if "WNI" in val:
                fields["kewarganegaraan"] = "WNI"
            elif "WNA" in val:
                fields["kewarganegaraan"] = "WNA"
            elif "INDONES" in val:
                fields["kewarganegaraan"] = "WNI"
            else:
                fields["kewarganegaraan"] = val

        elif "BERLAKU" in line:
            fields["berlaku_hingga"] = "Seumur Hidup"

    return fields

class OCRProcessor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.image_handler = ImageHandler()
        self.text_processor = TextProcessor()
        
        # Inisialisasi EasyOCR
        self.logger.info("Memuat EasyOCR reader...")
        self.reader = easyocr.Reader(
            OCR_CONFIG['languages'],
            gpu=OCR_CONFIG['gpu']
        )
        self.logger.info("EasyOCR reader berhasil dimuat!")
    
    def process_image(self, image_path):
        """
        Memproses satu gambar dengan OCR
        
        Args:
            image_path (str): Path ke file gambar
            
        Returns:
            bool: True jika berhasil, False jika gagal
        """
        try:
            image_path = Path(image_path)
            
            # Validasi file
            if not self.image_handler.validate_image(image_path):
                self.logger.error(f"File tidak valid: {image_path}")
                return False
            
            # Load dan preprocess gambar
            image = self.image_handler.load_image(image_path)
            if image is None:
                return False
            
            # Preprocess image untuk OCR yang lebih baik
            processed_image = self.image_handler.preprocess_image(image)
            
            # Lakukan OCR
            self.logger.info("Melakukan OCR...")
            results = self.reader.readtext(
                processed_image,
                detail=OCR_CONFIG['detail'],
                paragraph=OCR_CONFIG['paragraph'],
                width_ths=OCR_CONFIG['width_ths'],
                height_ths=OCR_CONFIG['height_ths']
            )
            
            if not results:
                self.logger.warning("Tidak ada text terdeteksi dalam gambar")
                return True
            
            # Process hasil OCR
            processed_text = self.text_processor.process_results(results)
            
            # Simpan hasil
            self._save_results(image_path, processed_text, results)
            
            # Log hasil
            self.logger.info(f"Text terdeteksi: {len(results)} baris")
            self.logger.info("Preview text:")
            preview = processed_text[:100] + "..." if len(processed_text) > 100 else processed_text
            self.logger.info(f"'{preview}'")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error dalam process_image: {str(e)}")
            return False
    
    def _save_results(self, image_path, text, raw_results):
        """Simpan hasil OCR ke file"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_name = image_path.stem
            
            # Simpan text bersih
            text_file = OUTPUT_DIR / f"{base_name}_{timestamp}_text.txt"
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(text)
            
            # Simpan hasil detail
            detail_file = OUTPUT_DIR / f"{base_name}_{timestamp}_detail.txt"
            with open(detail_file, 'w', encoding='utf-8') as f:
                f.write(f"OCR Results for: {image_path.name}\n")
                f.write(f"Processed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 50 + "\n\n")
                
                for i, result in enumerate(raw_results, 1):
                    if OCR_CONFIG['detail'] == 1:
                        bbox, text, confidence = result
                        f.write(f"[{i}] Text: '{text}'\n")
                        f.write(f"    Confidence: {confidence:.2f}\n")
                        f.write(f"    BBox: {bbox}\n\n")
                    else:
                        f.write(f"[{i}] {result}\n")
            
            self.logger.info(f"Hasil disimpan ke: {text_file.name}")
            
        except Exception as e:
            self.logger.error(f"Error menyimpan hasil: {str(e)}")