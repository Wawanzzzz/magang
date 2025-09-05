"""
Text Processor - Memproses hasil OCR text
"""
import logging
import re
from config import OCR_CONFIG

class TextProcessor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def process_results(self, ocr_results):
        """
        Memproses hasil OCR menjadi text yang bersih
        
        Args:
            ocr_results (list): Hasil dari EasyOCR
            
        Returns:
            str: Text yang sudah diproses
        """
        try:
            if not ocr_results:
                return ""
            
            # Extract text berdasarkan format hasil OCR
            texts = []
            
            for result in ocr_results:
                if OCR_CONFIG['detail'] == 1:
                    # Format: [bbox, text, confidence]
                    bbox, text, confidence = result
                    
                    # Filter berdasarkan confidence threshold
                    if confidence >= 0.5:  # Minimal 50% confidence
                        texts.append(text.strip())
                    else:
                        self.logger.debug(f"Text diabaikan (confidence rendah): '{text}' ({confidence:.2f})")
                else:
                    # Format: text only
                    texts.append(result.strip())
            
            # Gabungkan semua text
            raw_text = ' '.join(texts)
            
            # Clean dan normalize text
            cleaned_text = self._clean_text(raw_text)
            
            self.logger.info(f"Text processing selesai: {len(texts)} segmen text")
            return cleaned_text
            
        except Exception as e:
            self.logger.error(f"Error processing text: {str(e)}")
            return ""
    
    def _clean_text(self, text):
        """
        Membersihkan dan normalize text
        
        Args:
            text (str): Raw text
            
        Returns:
            str: Cleaned text
        """
        try:
            if not text:
                return ""
            
            # Remove extra whitespaces
            text = re.sub(r'\s+', ' ', text)
            
            # Fix common OCR errors
            text = self._fix_common_ocr_errors(text)
            
            # Remove leading/trailing whitespace
            text = text.strip()
            
            return text
            
        except Exception as e:
            self.logger.error(f"Error cleaning text: {str(e)}")
            return text
    
    def _fix_common_ocr_errors(self, text):
            """
            Perbaiki kesalahan OCR yang umum terjadi (huruf/angka + kata umum di KTP)
            """
            try:
                original_text = text  # Simpan untuk logging
                
                # Lowercase untuk memudahkan pencocokan kata
                text_lower = text.lower()

                # Kamus koreksi kata/kalimat umum di KTP
                word_corrections = {
                    r"jens kelam[ıi]n": "Jenis Kelamin",
                    r"gol\s*dara[hıi]": "Golongan Darah",
                    r"seumui hidup": "SEUMUR HIDUP",
                    r"tompavtol lah[ıi]r": "Tempat/Tgl Lahir",
                    r"nik+": "NIK",
                    r"nama[ıi]": "Nama",
                    r"provins[ıi]": "Provinsi",
                    r"kabupatenn?": "Kabupaten",
                    r"kecamatann?": "Kecamatan",
                    r"desaa?": "Desa",
                }

                # Terapkan regex word corrections (case-insensitive)
                for salah, benar in word_corrections.items():
                    text_lower = re.sub(salah, benar, text_lower, flags=re.IGNORECASE)

                # Perbaikan karakter umum OCR
                char_corrections = {
                    "0": ["O", "o"],
                    "1": ["l", "I", "|"],
                    "5": ["S"],
                    "6": ["G"],
                    "8": ["B"],
                }

                for benar, salah_list in char_corrections.items():
                    for salah in salah_list:
                        text_lower = text_lower.replace(salah, benar)

                # Rapikan spasi dan tanda baca
                text_lower = re.sub(r"\s+([.,:;!?])", r"\1", text_lower)
                text_lower = re.sub(r"([.!?])\s*([A-Z])", r"\1 \2", text_lower)

                # Logging hasil
                self.logger.debug(f"Before cleaning: {original_text}")
                self.logger.debug(f"After cleaning: {text_lower}")

                return text_lower

            except Exception as e:
                self.logger.error(f"Error fixing OCR errors: {str(e)}")
                return text