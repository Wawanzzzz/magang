"""
Enhanced Image Handler - Menangani loading dan preprocessing gambar untuk OCR yang lebih akurat
Versi sederhana yang fokus pada perbaikan masalah OCR dengan penyempurnaan kecil
"""
import logging
import cv2
import numpy as np
from PIL import Image
from pathlib import Path

from config import SUPPORTED_FORMATS

class ImageHandler:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def validate_image(self, image_path):
        """
        Validasi file gambar dengan validasi tambahan
        
        Args:
            image_path (Path): Path ke file gambar
            
        Returns:
            bool: True jika valid, False jika tidak
        """
        try:
            if not image_path.exists():
                self.logger.error(f"File tidak ditemukan: {image_path}")
                return False
            
            # Validasi ekstensi dengan menggunakan imghdr atau SUPPORTED_FORMATS
            if image_path.suffix.lower() not in SUPPORTED_FORMATS:
                self.logger.error(f"Format tidak didukung: {image_path.suffix}")
                self.logger.info(f"Format yang didukung: {', '.join(SUPPORTED_FORMATS)}")
                return False

            # Cek ukuran file (max 50MB)
            file_size = image_path.stat().st_size / (1024 * 1024)
            if file_size > 50:
                self.logger.error(f"File terlalu besar: {file_size:.1f}MB (max 50MB)")
                return False
            
            # Coba buka gambar untuk memastikan tidak corrupt
            try:
                with Image.open(image_path) as img:
                    img.verify()
            except Exception as ex:
                self.logger.error(f"File gambar corrupt atau tidak valid: {str(ex)}")
                return False

            return True
        except Exception as e:
            self.logger.error(f"Error validasi gambar: {str(e)}")
            return False
    
    def load_image(self, image_path):
        """
        Load gambar dengan fallback ke PIL jika OpenCV gagal
        
        Args:
            image_path (Path): Path ke file gambar
            
        Returns:
            numpy.ndarray: Image array atau None jika gagal
        """
        try:
            # Coba load dengan OpenCV terlebih dahulu
            image = cv2.imread(str(image_path))
            if image is None:
                # Jika OpenCV gagal, coba menggunakan PIL
                self.logger.warning(f"OpenCV gagal, mencoba PIL: {image_path}")
                try:
                    pil_image = Image.open(image_path)
                    image = np.array(pil_image)
                    if len(image.shape) == 3:
                        if image.shape[2] == 4:
                            image = cv2.cvtColor(image, cv2.COLOR_RGBA2RGB)
                        elif image.shape[2] == 3:
                            pass  # Sudah dalam format RGB
                except Exception as pil_error:
                    self.logger.error(f"PIL juga gagal: {str(pil_error)}")
                    return None
            else:
                image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
            
            if image is not None:
                self.logger.info(f"Gambar dimuat: {image.shape[1]}x{image.shape[0]} pixels")
            
            return image
        except Exception as e:
            self.logger.error(f"Error loading gambar: {str(e)}")
            return None

    def detect_orientation(self, image):
        """
        Deteksi orientasi gambar dengan pendekatan yang lebih reliabel
        
        Args:
            image (numpy.ndarray): Image array
            
        Returns:
            numpy.ndarray: Gambar yang sudah di-rotate
        """
        try:
            if len(image.shape) == 3:
                gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
            else:
                gray = image

            # Deteksi tepi dengan Canny untuk orientasi yang lebih presisi
            edges = cv2.Canny(gray, 50, 150)
            lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=100)
            
            if lines is not None:
                angles = [np.degrees(theta) for rho, theta in lines[:, 0]]
                median_angle = np.median(angles)
                
                if abs(median_angle) > 5:
                    self.logger.info(f"Rotasi gambar sebesar {median_angle:.1f} derajat")
                    return self.rotate_image(image, -median_angle)
            
            return image
        except Exception as e:
            self.logger.error(f"Error deteksi orientasi: {str(e)}")
            return image
    
    def preprocess_image(self, image, enhance_text=True, auto_rotate=True):
        """
        Preprocessing untuk OCR yang lebih akurat dengan peningkatan kontras dan rotasi otomatis
        
        Args:
            image (numpy.ndarray): Image array
            enhance_text (bool): Apakah menggunakan enhancement teks
            auto_rotate (bool): Apakah menggunakan rotasi otomatis
            
        Returns:
            numpy.ndarray: Gambar yang telah diproses
        """
        try:
            original_image = image.copy()
            
            # Auto rotation jika diperlukan
            if auto_rotate:
                image = self.detect_orientation(image)
            
            gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY) if len(image.shape) == 3 else image
            
            # Ukuran gambar untuk OCR optimal
            height, width = gray.shape
            if height < 900 or width < 900:
                scale = max(900 / height, 900 / width)
                scale = min(scale, 1.8)  # Sesuaikan scale agar hasil optimal
                new_width = int(width * scale)
                new_height = int(height * scale)
                gray = cv2.resize(gray, (new_width, new_height), interpolation=cv2.INTER_CUBIC)
                self.logger.info(f"Gambar di-resize ke: {new_width}x{new_height}")
            
            if enhance_text:
                gray = self.reduce_noise_advanced(gray)
                gray = self.enhance_contrast_adaptive(gray)
                gray = self.sharpen_text(gray)
            else:
                gray = cv2.medianBlur(gray, 3)
            
            gray = np.clip(gray, 0, 255).astype(np.uint8)
            self.logger.info("Preprocessing selesai dengan enhancement")
            return gray
        except Exception as e:
            self.logger.error(f"Error preprocessing gambar: {str(e)}")
            return original_image if 'original_image' in locals() else image