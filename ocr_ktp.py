import streamlit as st
import cv2
import pytesseract
import re
import numpy as np
import pandas as pd
from PIL import Image, ImageDraw
import io
import base64
from datetime import datetime
import os

# Set page config
st.set_page_config(
    page_title="KTP OCR Dashboard",
    page_icon="🆔",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        color: #2c3e50;
        text-align: center;
        margin-bottom: 2rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    
    .sub-header {
        font-size: 1.2rem;
        color: #7f8c8d;
        text-align: center;
        margin-bottom: 3rem;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin: 0.5rem 0;
    }
    
    .success-msg {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    
    .error-msg {
        background: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 5px;
        border: 1px solid #f5c6cb;
        margin: 1rem 0;
    }
    
    .info-msg {
        background: #cce7ff;
        color: #004085;
        padding: 1rem;
        border-radius: 5px;
        border: 1px solid #b3d7ff;
        margin: 1rem 0;
    }
    
    .warning-msg {
        background: #fff3cd;
        color: #856404;
        padding: 1rem;
        border-radius: 5px;
        border: 1px solid #ffeaa7;
        margin: 1rem 0;
    }
    
    .stDataFrame {
        border: 2px solid #e1e5e9;
        border-radius: 10px;
    }
    
    .crop-container {
        border: 2px dashed #667eea;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
        background: #f8f9fa;
    }
</style>
""", unsafe_allow_html=True)

class ImageCropper:
    def __init__(self):
        pass
    
    def create_crop_interface(self, image):
        """Create cropping interface using Streamlit components"""
        st.markdown('<div class="crop-container">', unsafe_allow_html=True)
        st.subheader("✂ Crop Area Selection")
        
        # Get image dimensions
        width, height = image.size
        
        # Create sliders for crop coordinates
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("📐 Horizontal (Width) Selection:")
            x_start = st.slider(
                "Start X (Left)",
                min_value=0,
                max_value=width-1,
                value=0,
                key="x_start",
                help="Drag untuk mengatur titik awal horizontal"
            )
            x_end = st.slider(
                "End X (Right)",
                min_value=x_start+1,
                max_value=width,
                value=width,
                key="x_end",
                help="Drag untuk mengatur titik akhir horizontal"
            )
        
        with col2:
            st.write("📏 Vertical (Height) Selection:")
            y_start = st.slider(
                "Start Y (Top)",
                min_value=0,
                max_value=height-1,
                value=0,
                key="y_start",
                help="Drag untuk mengatur titik awal vertikal"
            )
            y_end = st.slider(
                "End Y (Bottom)",
                min_value=y_start+1,
                max_value=height,
                value=height,
                key="y_end",
                help="Drag untuk mengatur titik akhir vertikal"
            )
        
        # Show crop dimensions
        crop_width = x_end - x_start
        crop_height = y_end - y_start
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Original Size", f"{width}×{height}")
        with col2:
            st.metric("Crop Size", f"{crop_width}×{crop_height}")
        with col3:
            percentage = round((crop_width * crop_height) / (width * height) * 100, 1)
            st.metric("Area Percentage", f"{percentage}%")
        with col4:
            if crop_width > 100 and crop_height > 50:
                st.metric("Status", "✅ Valid")
            else:
                st.metric("Status", "❌ Too Small")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        return (x_start, y_start, x_end, y_end)
    
    def create_preview_with_overlay(self, image, crop_coords):
        """Create preview image with crop overlay"""
        x_start, y_start, x_end, y_end = crop_coords
        
        # Create a copy of the image for overlay
        preview_image = image.copy()
        draw = ImageDraw.Draw(preview_image)
        
        # Draw crop rectangle
        draw.rectangle(
            [x_start, y_start, x_end, y_end],
            outline="red",
            width=3
        )
        
        # Add semi-transparent overlay outside crop area
        overlay = Image.new('RGBA', image.size, (0, 0, 0, 128))
        overlay_draw = ImageDraw.Draw(overlay)
        
        # Draw overlay areas (outside crop)
        if y_start > 0:  # Top area
            overlay_draw.rectangle([0, 0, image.width, y_start], fill=(0, 0, 0, 128))
        if y_end < image.height:  # Bottom area
            overlay_draw.rectangle([0, y_end, image.width, image.height], fill=(0, 0, 0, 128))
        if x_start > 0:  # Left area
            overlay_draw.rectangle([0, y_start, x_start, y_end], fill=(0, 0, 0, 128))
        if x_end < image.width:  # Right area
            overlay_draw.rectangle([x_end, y_start, image.width, y_end], fill=(0, 0, 0, 128))
        
        # Combine images
        preview_image = Image.alpha_composite(preview_image.convert('RGBA'), overlay)
        
        return preview_image.convert('RGB')
    
    def crop_image(self, image, crop_coords):
        """Crop image based on coordinates"""
        x_start, y_start, x_end, y_end = crop_coords
        
        # Validate crop coordinates
        if x_end <= x_start or y_end <= y_start:
            return None
        
        # Crop image
        cropped = image.crop((x_start, y_start, x_end, y_end))
        
        return cropped
class EnhancedKTPPreprocessor:
    """Enhanced preprocessing untuk KTP dengan security pattern dan noise tinggi"""
    
    def __init__(self):
        pass
    
    def remove_security_patterns(self, image):
        """Menghilangkan pola keamanan/watermark pada KTP"""
        # Convert PIL to OpenCV format
        img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 1. Median filter untuk mengurangi noise grain
        denoised = cv2.medianBlur(gray, 3)
        
        # 2. Morphological opening untuk menghilangkan pola kecil
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2, 2))
        opened = cv2.morphologyEx(denoised, cv2.MORPH_OPEN, kernel)
        
        # 3. Gaussian blur untuk meratakan background noise
        blurred = cv2.GaussianBlur(opened, (3, 3), 0.5)
        
        return blurred
    
    def enhance_text_contrast(self, gray_image):
        """Meningkatkan kontras khusus untuk teks"""
        # 1. CLAHE dengan parameter yang lebih agresif
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray_image)
        
        # 2. Histogram stretching
        min_val, max_val = np.percentile(enhanced, [2, 98])  # Ignore extreme outliers
        if max_val > min_val:
            stretched = np.clip((enhanced - min_val) * 255 / (max_val - min_val), 0, 255)
            enhanced = stretched.astype(np.uint8)
        
        # 3. Gamma correction untuk mencerahkan teks
        gamma = 0.7  # Lebih agresif untuk teks gelap
        inv_gamma = 1.0 / gamma
        table = np.array([((i / 255.0) ** inv_gamma) * 255 for i in np.arange(0, 256)]).astype("uint8")
        enhanced = cv2.LUT(enhanced, table)
        
        return enhanced
    
    def advanced_thresholding(self, gray_image):
        """Thresholding yang lebih canggih untuk KTP dengan noise"""
        results = []
        
        # 1. Adaptive Gaussian dengan parameter yang disesuaikan untuk noise
        adaptive_gauss = cv2.adaptiveThreshold(
            gray_image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
            cv2.THRESH_BINARY, 21, 8  # Block size lebih besar untuk noise
        )
        results.append(adaptive_gauss)
        
        # 2. Adaptive Mean threshold
        adaptive_mean = cv2.adaptiveThreshold(
            gray_image, 255, cv2.ADAPTIVE_THRESH_MEAN_C, 
            cv2.THRESH_BINARY, 19, 6
        )
        results.append(adaptive_mean)
        
        # 3. Otsu dengan Gaussian blur preprocessing
        blurred = cv2.GaussianBlur(gray_image, (5, 5), 0)
        _, otsu = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        results.append(otsu)
        
        # 4. Kombinasi CLAHE + Otsu
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        clahe_img = clahe.apply(gray_image)
        _, clahe_otsu = cv2.threshold(clahe_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        results.append(clahe_otsu)
        
        return results
    
    def preprocess_ktp_enhanced(self, image):
        """Pipeline preprocessing yang dioptimalkan untuk KTP dengan security pattern"""
        processed_images = []
        
        # Step 1: Remove security patterns and noise
        clean_gray = self.remove_security_patterns(image)
        processed_images.append(clean_gray)
        
        # Step 2: Enhance text contrast
        enhanced = self.enhance_text_contrast(clean_gray)
        processed_images.append(enhanced)
        
        # Step 3: Focus on upper region (lebih penting untuk KTP)
        height, width = enhanced.shape
        upper_region = enhanced[:int(height * 0.75), :]  # 75% bagian atas
        processed_images.append(upper_region)
        
        # Step 4: Advanced thresholding
        threshold_results = self.advanced_thresholding(enhanced)
        processed_images.extend(threshold_results)
        
        # Step 5: Threshold untuk upper region
        upper_threshold_results = self.advanced_thresholding(upper_region)
        processed_images.extend(upper_threshold_results)
        
        # Step 6: Edge enhancement untuk teks yang blur
        try:
            # Unsharp masking yang lebih agresif
            gaussian = cv2.GaussianBlur(enhanced, (5, 5), 2.0)
            unsharp = cv2.addWeighted(enhanced, 2.0, gaussian, -1.0, 0)
            unsharp = np.clip(unsharp, 0, 255).astype(np.uint8)
            processed_images.append(unsharp)
        except:
            pass
        
        # Step 7: Bilateral filter untuk noise reduction sambil preserve edges
        try:
            bilateral = cv2.bilateralFilter(enhanced, 9, 80, 80)
            processed_images.append(bilateral)
            
            # Bilateral + threshold
            bilateral_thresh = self.advanced_thresholding(bilateral)
            processed_images.extend(bilateral_thresh[:2])  # Ambil 2 yang terbaik
        except:
            pass
        
        return processed_images

class KTPExtractor:
    
    def __init__(self):
        # Konfigurasi Tesseract (sesuaikan path jika diperlukan)
        if os.name == 'nt':  # Windows
            try:
                pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            except:
                pass  # Let pytesseract use default path
        
        # Patterns yang diperbaiki untuk ekstraksi data KTP Indonesia
        self.patterns = {
    "Provinsi": [
        r'PROVINSI\s+([A-Z\s]{3,40})(?=\n|KOTA|KABUPATEN|$)',
    ],
    "Kabupaten": [
        r'KOTA\s+([A-Z\s]{3,40})(?=\n|NIK|$)',
        r'KABUPATEN\s+([A-Z\s]{3,40})(?=\n|NIK|$)',
    ],
    "NIK": [
    # Matches 'NIK' + optional colon/dash + exactly 16 digits
    r'NIK\s*[:\-]?\s*([0-9]{16})(?=\s|\n|$)'
],


    "Nama": [
        r'NAMA\s*[:\-]?\s*([A-Z\s]+?)(?=\n|TEMPAT|Tempat|TGL|LAHIR|JENIS)'  # extended stop words
    ],
    "Tempat Tgl Lahir": [
        r'TEMPAT\/TGL\s*LAHIR\s*:?\s*([A-Z\s]+,\s*\d{2}[\-\/]\d{2}[\-\/]\d{4})',
        r'Tempat\/Tgl\s*Lahir\s*:?\s*([A-Z\s]+,\s*\d{2}[\-\/]\d{2}[\-\/]\d{4})',
    ],
    "Jenis Kelamin": [
        r'JENIS\s*KELAMIN\s*:?\s*(LAKI[\-\s]*LAKI|PEREMPUAN)',
        r'Jenis\s*Kelamin\s*:?\s*(LAKI[\-\s]*LAKI|PEREMPUAN)',
    ],
    "Gol Darah": [
        r'GOL\.?\s*DARAH\s*:?\s*([ABO]{1,2}[+\-]?)',
    ],
    "Alamat": [
        r'ALAMAT\s*[:\-]?\s*([A-Z0-9\s\/\.\-]+?)(?=\n|RT|Rw|RW|KEL|KEL\/DESA|DESA|KEC|AGAMA|PEKERJAAN)'
    ],
    "RT RW": [
        r'RT\s*[:\-]?\s*0*([0-9]{1,3})\s*[\/\-]\s*RW\s*[:\-]?\s*0*([0-9]{1,3})',
        r'RT\s*[:\-]?\s*0*([0-9]{1,3})\s*RW\s*[:\-]?\s*0*([0-9]{1,3})'
    ],
    "Kel Desa": [
        r'KEL\/DESA\s*[:\-]?\s*([A-Z\s]+?)(?=\n|KECAMATAN|KEC|AGAMA|PEKERJAAN)',
        r'KELURAHAN\s*[:\-]?\s*([A-Z\s]+?)(?=\n|KECAMATAN|KEC|AGAMA|PEKERJAAN)',
        r'DESA\s*[:\-]?\s*([A-Z\s]+?)(?=\n|KECAMATAN|KEC|AGAMA|PEKERJAAN)'
    ],
    "Kecamatan": [
        r'KEC(?:AMATAN)?\s*[:\-]?\s*([A-Z\s]+?)(?=\n|AGAMA|PEKERJAAN|STATUS|Kewarganegaraan)'
    ],
    "Agama": [
        r'AGAMA\s*[:\-]?\s*(ISLAM|KRISTEN|KATOLIK|HINDU|BUDDHA|KONGHUCU)'
    ],
    "Status Perkawinan": [
        r'STATUS\s*PERKAWINAN\s*:?\s*(BELUM\s*KAWIN|KAWIN|MENIKAH|CERAI\s*HIDUP|CERAI\s*MATI)',
    ],
    "Pekerjaan": [
        r'PEKERJAAN\s*[:\-]?\s*([A-Z\s\/\.]+?)(?=\n|STATUS|KEWARGANEGARAAN|BERLAKU)'
    ],
    "Kewarganegaraan": [
        r'KEWARGANEGARAAN\s*:?\s*(WNI|WNA)',
    ],
    "Berlaku Hingga": [
        r'BERLAKU\s*HINGGA\s*:?\s*([A-Z\s]+)',
    ]
}


    def preprocess_image(self, image):
        """Preprocessing khusus untuk bagian atas KTP dengan noise reduction dan kontras enhancement"""
        img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        processed_images = []
        
        # 1. Fokus pada bagian atas KTP (crop 70% bagian atas)
        height, width = gray.shape
        top_region = gray[:int(height * 0.7), :]
        processed_images.append(top_region)
        
        # 2. Noise reduction dengan bilateral filter untuk text clarity
        try:
            bilateral = cv2.bilateralFilter(gray, 9, 75, 75)
            processed_images.append(bilateral)
            
            # Bilateral untuk top region
            bilateral_top = cv2.bilateralFilter(top_region, 9, 75, 75)
            processed_images.append(bilateral_top)
        except:
            pass
        
        # 3. Morphological operations untuk menghilangkan noise
        try:
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 1))
            opening = cv2.morphologyEx(gray, cv2.MORPH_OPEN, kernel)
            processed_images.append(opening)
            
            # Opening untuk top region
            opening_top = cv2.morphologyEx(top_region, cv2.MORPH_OPEN, kernel)
            processed_images.append(opening_top)
        except:
            pass
        
        # 4. CLAHE - Contrast enhancement khusus untuk text
        try:
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(gray)
            processed_images.append(enhanced)
            
            enhanced_top = clahe.apply(top_region)
            processed_images.append(enhanced_top)
        except:
            pass
        
        # 5. Adaptive threshold yang lebih clean
        try:
            # Gaussian adaptive threshold
            thresh_gauss = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 4)
            processed_images.append(thresh_gauss)
            
            thresh_gauss_top = cv2.adaptiveThreshold(top_region, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 4)
            processed_images.append(thresh_gauss_top)
        except:
            pass
        
        # 6. Otsu threshold untuk clean text
        try:
            _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_images.append(otsu)
            
            _, otsu_top = cv2.threshold(top_region, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_images.append(otsu_top)
        except:
            pass
        
        
        # 7. Histogram Equalization untuk kontras global
        try:
            hist_eq = cv2.equalizeHist(gray)
            processed_images.append(hist_eq)
            
            hist_eq_top = cv2.equalizeHist(top_region)
            processed_images.append(hist_eq_top)
        except:
            pass
        
        # 8. Contrast Stretching untuk maksimalkan range dinamis
        try:
            # Stretch kontras ke range penuh 0-255
            min_val, max_val = np.min(gray), np.max(gray)
            if max_val != min_val:
                stretched = ((gray - min_val) / (max_val - min_val)) * 255
                stretched = stretched.astype(np.uint8)
                processed_images.append(stretched)
                
                # Stretch untuk top region
                min_val_top, max_val_top = np.min(top_region), np.max(top_region)
                if max_val_top != min_val_top:
                    stretched_top = ((top_region - min_val_top) / (max_val_top - min_val_top)) * 255
                    stretched_top = stretched_top.astype(np.uint8)
                    processed_images.append(stretched_top)
        except:
            pass
        
        # 9. Gamma Correction untuk brightening teks
        try:
            # Gamma < 1 untuk mencerahkan area gelap (teks)
            gamma = 0.8
            inv_gamma = 1.0 / gamma
            table = np.array([((i / 255.0) ** inv_gamma) * 255 for i in np.arange(0, 256)]).astype("uint8")
            gamma_corrected = cv2.LUT(gray, table)
            processed_images.append(gamma_corrected)
            
            gamma_corrected_top = cv2.LUT(top_region, table)
            processed_images.append(gamma_corrected_top)
        except:
            pass
        
        # 10. Unsharp Masking untuk ketajaman teks
        try:
            # Gaussian blur
            blurred = cv2.GaussianBlur(gray, (3, 3), 1.0)
            # Unsharp mask
            unsharp = cv2.addWeighted(gray, 1.5, blurred, -0.5, 0)
            processed_images.append(unsharp)
            
            # Unsharp untuk top region
            blurred_top = cv2.GaussianBlur(top_region, (3, 3), 1.0)
            unsharp_top = cv2.addWeighted(top_region, 1.5, blurred_top, -0.5, 0)
            processed_images.append(unsharp_top)
        except:
            pass
        
        # 11. Kombinasi terbaik: CLAHE + Gamma + Unsharp
        try:
            # Step 1: CLAHE dengan parameter yang sama seperti #4
            clahe_combo = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            combo = clahe_combo.apply(gray)
            
            # Step 2: Gamma correction
            gamma = 0.8
            inv_gamma = 1.0 / gamma
            table = np.array([((i / 255.0) ** inv_gamma) * 255 for i in np.arange(0, 256)]).astype("uint8")
            combo = cv2.LUT(combo, table)
            
            # Step 3: Unsharp masking
            blurred = cv2.GaussianBlur(combo, (3, 3), 1.0)
            combo = cv2.addWeighted(combo, 1.3, blurred, -0.3, 0)
            
            processed_images.append(combo)
            
            # Combo untuk top region
            combo_top = clahe_combo.apply(top_region)
            combo_top = cv2.LUT(combo_top, table)
            blurred_top = cv2.GaussianBlur(combo_top, (3, 3), 1.0)
            combo_top = cv2.addWeighted(combo_top, 1.3, blurred_top, -0.3, 0)
            processed_images.append(combo_top)
            
        except:
            pass
        
        # 12. Edge enhancement untuk teks yang lebih sharp
        try:
            # Laplacian edge detection dan combine dengan original
            laplacian = cv2.Laplacian(gray, cv2.CV_64F)
            laplacian = np.uint8(np.absolute(laplacian))
            edge_enhanced = cv2.add(gray, laplacian)
            processed_images.append(edge_enhanced)
            
            # Edge enhancement untuk top region
            laplacian_top = cv2.Laplacian(top_region, cv2.CV_64F)
            laplacian_top = np.uint8(np.absolute(laplacian_top))
            edge_enhanced_top = cv2.add(top_region, laplacian_top)
            processed_images.append(edge_enhanced_top)
        except:
            pass
        
        return processed_images

    def clean_text_advanced(self, text):
        """Pembersihan teks yang lebih agresif untuk mengurangi noise"""
        if not text:
            return ""
        
        # Hapus karakter non-printable dan symbol aneh
        text = re.sub(r'[^\x20-\x7E\n]', '', text)
        
        # Hapus garis horizontal dan vertical artifacts
        text = re.sub(r'[\-_]{3,}', ' ', text)
        text = re.sub(r'[\|]{2,}', ' ', text)
        
        # Bersihkan multiple spaces dan newlines
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\n\s*\n', '\n', text)
        
        # Hapus baris yang hanya berisi karakter repetitif
        lines = text.split('\n')
        clean_lines = []
        
        for line in lines:
            line = line.strip()
            if len(line) < 2:  # Skip baris terlalu pendek
                continue
            
            # Skip baris yang hanya berisi karakter repetitif
            if len(set(line.replace(' ', ''))) < 2:
                continue
            
            # Skip baris yang kebanyakan angka tanpa konteks
            if re.match(r'^\d+\s*$', line) and len(line) < 5:
                continue
            
            # Hanya ambil huruf, angka, spasi, dan tanda baca penting
            cleaned_line = re.sub(r'[^A-Za-z0-9\s:\-\/,.]', '', line)
            if len(cleaned_line.strip()) > 1:
                clean_lines.append(cleaned_line.strip())
        
        return '\n'.join(clean_lines)

    def extract_text_from_image(self, processed_images):
        """OCR dengan konfigurasi yang dioptimalkan untuk KTP Indonesia"""
        try:
            # Konfigurasi OCR yang fokus pada text accuracy
            ocr_configs = [
                # Bahasa Indonesia dengan whitelist karakter umum KTP
                r'--oem 3 --psm 6 -l ind -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789:/-., ',
                r'--oem 3 --psm 4 -l ind -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789:/-., ',
                r'--oem 3 --psm 3 -l ind',  # Fully automatic tanpa whitelist
                r'--oem 3 --psm 6 -l ind',  # Block text
                
                # Mixed language dengan whitelist
                r'--oem 3 --psm 6 -l ind+eng -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789:/-., ',
                r'--oem 3 --psm 4 -l ind+eng -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789:/-., ',
                
                # Fallback tanpa whitelist
                r'--oem 3 --psm 6 -l eng',
                r'--oem 3 --psm 4 -l eng'
            ]
            
            all_texts = []
            
            # Try each processed image dengan prioritas pada top region
            for i, img in enumerate(processed_images):
                for config in ocr_configs:
                    try:
                        text = pytesseract.image_to_string(img, config=config)
                        if text and len(text.strip()) > 10:  # Minimum length untuk text yang bermakna
                            cleaned = self.clean_text_advanced(text)
                            if cleaned:
                                # Berikan priority score untuk top region images
                                priority = 2 if i < len(processed_images)//2 else 1
                                all_texts.append((cleaned, priority))
                    except Exception as e:
                        continue
            
            if not all_texts:
                return ""
            
            # Pilih text terbaik berdasarkan score
            best_text = ""
            best_score = 0
            
            for text, priority in all_texts:
                # Score berdasarkan: panjang text + keyword KTP + priority
                ktp_keywords = ['NIK', 'NAMA', 'TEMPAT', 'LAHIR', 'ALAMAT', 'AGAMA', 'PEKERJAAN', 'PROVINSI', 'KABUPATEN']
                keyword_count = sum(1 for keyword in ktp_keywords if keyword in text.upper())
                
                # Penalti untuk text dengan terlalu banyak noise
                noise_count = len(re.findall(r'[^A-Za-z0-9\s:\-\/,.]', text))
                noise_penalty = noise_count * 2
                
                score = len(text) + (keyword_count * 100) + (priority * 50) - noise_penalty
                
                if score > best_score:
                    best_score = score
                    best_text = text
            
            return best_text
            
        except Exception as e:
            st.error(f"Error dalam OCR: {str(e)}")
            return ""

    def extract_rt_rw(self, text):
        """Ekstraksi khusus untuk RT/RW dengan deteksi garis pemisah yang diperbaiki"""
        rt_rw_patterns = [
            # Pattern utama dengan berbagai format
            r'RT[\s\/]RW\s*:?\s*(\d{1,3})[\s\/\-]+(\d{1,3})',
            r'RT\s*:?\s*(\d{1,3})[\s\/\-]*RW\s*:?\s*(\d{1,3})',
            r'RT\s*(\d{1,3})\s*[\/\-]\s*RW\s*(\d{1,3})',
            r'RT\s*(\d{1,3})\s*RW\s*(\d{1,3})',
            
            # Pattern untuk format 3 digit
            r'(\d{3})[\s\/\-](\d{3})(?=\s*\n.*(?:KEL|DESA|Kel|Desa))',
            r'(\d{2,3})[\s\/\-](\d{2,3})(?=\s*\n.*(?:KEL|DESA|Kel|Desa))',
            
            # Pattern setelah alamat
            r'(?:ALAMAT.*\n.*?)(\d{2,3})[\s\/\-](\d{2,3})(?=\s*\n)',
            
            # Pattern umum RT/RW
            r'(?:RT|Rt)\s*(\d{1,3})\s*(?:RW|Rw)\s*(\d{1,3})'
        ]
        
        for pattern in rt_rw_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
            for match in matches:
                if isinstance(match, tuple) and len(match) == 2:
                    rt, rw = match[0].strip(), match[1].strip()
                    
                    # Validasi: harus berupa angka dan panjang wajar
                    if rt.isdigit() and rw.isdigit() and 1 <= len(rt) <= 3 and 1 <= len(rw) <= 3:
                        # Skip jika angka terlalu besar (kemungkinan bukan RT/RW)
                        if int(rt) <= 999 and int(rw) <= 999:
                            return f"{rt.zfill(3)}/{rw.zfill(3)}"
                            
                elif isinstance(match, str) and '/' in match:
                    parts = re.split(r'[\s\/\-]+', match)
                    if len(parts) == 2 and all(p.isdigit() for p in parts):
                        rt, rw = parts[0], parts[1]
                        if int(rt) <= 999 and int(rw) <= 999:
                            return f"{rt.zfill(3)}/{rw.zfill(3)}"
        
        return "Tidak terdeteksi"

    def extract_field_value(self, text, field_name):
        """Ekstrak nilai field dengan validasi yang lebih ketat"""
        if field_name not in self.patterns:
            return "Tidak terdeteksi"
        
        # Special handling untuk RT/RW
        if field_name == "RT RW":
            return self.extract_rt_rw(text)
        
        # Coba setiap pattern
        for pattern in self.patterns[field_name]:
            try:
                matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
                for match in matches:
                    if isinstance(match, tuple):
                        value = match[0].strip() if match[0] else ""
                    else:
                        value = match.strip()
                    
                    # Validasi dan pembersihan berdasarkan field type
                    cleaned_value = self.validate_field_value(field_name, value)
                    if cleaned_value and cleaned_value != "Tidak terdeteksi":
                        return cleaned_value
                        
            except Exception:
                continue
        
        # Fallback ke keyword search yang diperbaiki
        return self.improved_keyword_search(text, field_name)

    def validate_field_value(self, field_name, value):
        """Validasi dan pembersihan nilai berdasarkan jenis field - DIPERBAIKI"""
        if not value or len(value.strip()) < 1:
            return "Tidak terdeteksi"
        
        value = value.strip()
        
        # Validasi berdasarkan field
        if field_name == "NIK":
            # Ekstrak hanya digit dari value
            digits = re.findall(r'\d', value)
            digit_string = ''.join(digits)
            
            # Harus tepat 16 digit
            if len(digit_string) == 16:
                return digit_string
            elif len(digit_string) > 16:
                # Ambil 16 digit pertama jika ada lebih dari 16
                return digit_string[:16]
            else:
                # Jika kurang dari 16 digit, cari pattern 16 digit dalam teks
                sixteen_digit_pattern = re.search(r'(\d{16})', value)
                if sixteen_digit_pattern:
                    return sixteen_digit_pattern.group(1)
                
            return "Tidak terdeteksi"
                    
        elif field_name in ["Provinsi", "Kabupaten", "Kel Desa", "Kecamatan"]:
            # Hapus prefix yang tidak diinginkan seperti "PROVINSI:", "KABUPATEN:", dll
            prefixes_to_remove = [
                "PROVINSI", "KABUPATEN", "KAB", "KOTA", "KEL", "DESA", 
                "KELURAHAN", "KECAMATAN", "KEC", ":"
            ]
            
            cleaned = value.upper()
            for prefix in prefixes_to_remove:
                cleaned = re.sub(f'^{prefix}\\s*:?\\s*', '', cleaned)
            
            # Hanya ambil huruf dan spasi
            cleaned = re.sub(r'[^A-Z\s]', '', cleaned)
            cleaned = re.sub(r'\s+', ' ', cleaned).strip()
            
            if len(cleaned) >= 3 and cleaned[0].isalpha():
                return cleaned
                    
        elif field_name == "Nama":
            # Hapus prefix "NAMA:" jika ada
            cleaned = re.sub(r'^NAMA\s*:?\s*', '', value.upper())
            
            # Hanya ambil huruf dan spasi
            cleaned = re.sub(r'[^A-Z\s]', '', cleaned)
            cleaned = re.sub(r'\s+', ' ', cleaned).strip()
            
            # Validasi: nama harus 3-40 karakter, tidak boleh mengandung judul field
            forbidden_words = ["TEMPAT", "TGL", "LAHIR", "KELAMIN", "ALAMAT"]
            if 3 <= len(cleaned) <= 40 and not any(word in cleaned for word in forbidden_words):
                return cleaned
                    
        elif field_name == "Tempat Tgl Lahir":
            # Hapus prefix jika ada
            cleaned = re.sub(r'^(?:TEMPAT|TGL|LAHIR|\s|:)+', '', value, flags=re.IGNORECASE)
            
            # Cari pattern tempat dan tanggal lahir
            pattern = r'([A-Z][A-Z\s]{2,20}),?\s*(\d{1,2}[\-\/]\d{1,2}[\-\/]\d{4})'
            match = re.search(pattern, cleaned.upper())
            
            if match:
                tempat = match.group(1).strip()
                tanggal = match.group(2).strip()
                
                # Validasi tempat tidak mengandung kata kunci field lain
                forbidden_words = ["KELAMIN", "ALAMAT", "AGAMA", "PEKERJAAN"]
                if not any(word in tempat for word in forbidden_words):
                    return f"{tempat}, {tanggal}"
            
            return "Tidak terdeteksi"
                    
        elif field_name == "Jenis Kelamin":
            # Standarisasi gender
            upper_val = value.upper()
            if 'LAKI' in upper_val:
                return "LAKI-LAKI"
            elif 'PEREMPUAN' in upper_val:
                return "PEREMPUAN"
            elif upper_val in ['L', 'P']:
                return "LAKI-LAKI" if upper_val == 'L' else "PEREMPUAN"
                    
        elif field_name == "Gol Darah":
            # Validasi golongan darah
            cleaned = re.sub(r'[^ABCO\+\-]', '', value.upper())
            if cleaned and len(cleaned) <= 3:
                return cleaned
                    
        elif field_name == "Agama":
            # Validasi agama yang diakui
            agama_valid = ['ISLAM', 'KRISTEN', 'KATOLIK', 'HINDU', 'BUDDHA', 'KONGHUCU']
            upper_val = value.upper()
            for agama in agama_valid:
                if agama in upper_val:
                    return agama
                        
        elif field_name == "Status Perkawinan":
            # Validasi status perkawinan
            status_valid = ['BELUM KAWIN', 'KAWIN', 'CERAI HIDUP', 'CERAI MATI']
            upper_val = value.upper()
            for status in status_valid:
                if status.replace(' ', '') in upper_val.replace(' ', ''):
                    return status
                        
        elif field_name == "Kewarganegaraan":
            # Validasi kewarganegaraan
            upper_val = value.upper()
            if 'WNI' in upper_val:
                return "WNI"
            elif 'WNA' in upper_val:
                return "WNA"
        
        # Default cleaning untuk field lain
        cleaned = re.sub(r'[^A-Za-z0-9\s:\-\/,.]', '', value)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        
        if len(cleaned) >= 2:
            return cleaned
            
        return "Tidak terdeteksi"

    def improved_keyword_search(self, text, field_name):
        """Keyword search yang diperbaiki dengan validasi"""
        keywords_map = {
            "NIK": ["NIK", "NOMOR", "INDUK"],
            "Nama": ["NAMA", "NAME"],
            "Tempat Tgl Lahir": ["TEMPAT", "LAHIR", "TGL"],
            "Jenis Kelamin": ["KELAMIN", "JENIS"],
            "Alamat": ["ALAMAT", "ADDRESS", "JALAN"],
            "Agama": ["AGAMA", "RELIGION"],
            "Pekerjaan": ["PEKERJAAN", "KERJA", "PROFESI"],
            "Provinsi": ["PROVINSI", "PROV"],
            "Kabupaten": ["KABUPATEN", "KAB", "KOTA"]
        }
        
        if field_name not in keywords_map:
            return "Tidak terdeteksi"
        
        keywords = keywords_map[field_name]
        text_upper = text.upper()
        
        for keyword in keywords:
            if keyword in text_upper:
                keyword_pos = text_upper.find(keyword)
                if keyword_pos != -1:
                    # Ambil teks setelah keyword dalam radius yang wajar
                    after_keyword = text[keyword_pos + len(keyword):keyword_pos + len(keyword) + 100]
                    
                    # Extract value setelah keyword
                    lines = after_keyword.split('\n')
                    for line in lines[:3]:  # Cek 3 baris pertama
                        line = line.strip()
                        if not line:
                            continue
                            
                        # Skip line yang hanya berisi simbol
                        if re.match(r'^[\s:\-\.]+$', line):
                            continue
                            
                        # Ambil bagian value
                        value = re.sub(r'^[\s:\-\.]+', '', line)  # Hapus prefix
                        value = value.split()[0] if field_name == "NIK" else ' '.join(value.split()[:5])
                        
                        validated = self.validate_field_value(field_name, value)
                        if validated and validated != "Tidak terdeteksi":
                            return validated
        
        return "Tidak terdeteksi"

    def extract_ktp_data(self, image):
        """Ekstrak data KTP dengan akurasi tinggi dan output yang lebih baik"""
        try:
            # Preprocessing dengan fokus pada bagian atas
            processed_images = self.preprocess_image(image)
            
            # OCR dengan konfigurasi optimal
            full_text = self.extract_text_from_image(processed_images)
            
            if not full_text.strip():
                # Last resort: OCR simple pada gambar asli
                try:
                    img_array = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                    gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
                    full_text = pytesseract.image_to_string(gray, config=r'--oem 3 --psm 3 -l ind')
                    full_text = self.clean_text_advanced(full_text)
                except:
                    pass
            
            if not full_text.strip():
                error_data = self._create_empty_result()
                return error_data, "Tidak ada teks yang dapat diekstrak dari gambar"
            
            # Extract fields dengan validasi ketat
            extracted_data = {}
            fields_found = 0
            confidence_scores = {}
            
            # Define field order for better presentation
            field_order = [
                "Provinsi", "Kabupaten", "NIK", "Nama", 
                "Tempat Tgl Lahir", "Jenis Kelamin", "Gol Darah",
                "Alamat", "RT RW", "Kel Desa", "Kecamatan",
                "Agama", "Status Perkawinan", "Pekerjaan", "Kewarganegaraan"
            ]
            
            # Extract each field and calculate confidence
            for field_name in field_order:
                value, confidence = self.extract_field_value_with_confidence(full_text, field_name)
                extracted_data[field_name] = value
                confidence_scores[field_name] = confidence
                
                if value and value != "Tidak terdeteksi" and value != "Error dalam ekstraksi":
                    fields_found += 1
            
            # Calculate overall accuracy based on confidence scores
            valid_scores = [score for score in confidence_scores.values() if score > 0]
            accuracy = round(sum(valid_scores) / len(valid_scores) * 100 if valid_scores else 0, 1)
            
            # Add quality indicators
            quality_indicator = self._determine_quality_indicator(accuracy, fields_found)
            
            # Add comprehensive metadata with safe key access
            extracted_data['_metadata'] = {
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'accuracy_percentage': f"{accuracy}%",
                'fields_detected': f"{fields_found}/{len(field_order)}",
                'quality_indicator': quality_indicator,
                'confidence_scores': confidence_scores,
                'extraction_method': 'Enhanced OCR with Validation',
                'text_length': len(full_text),
                'processing_status': 'Success' if fields_found > 5 else 'Partial'
            }
            
            # For backward compatibility with existing UI - use safe get method
            extracted_data['Timestamp'] = extracted_data['_metadata'].get('timestamp', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            extracted_data['Accuracy'] = extracted_data['_metadata'].get('accuracy_percentage', '0%')
            extracted_data['Fields Found'] = extracted_data['_metadata'].get('fields_detected', f'0/{len(field_order)}')
            
            return extracted_data, full_text
            
        except Exception as e:
            # Enhanced error handling with detailed error info
            error_data = self._create_empty_result()
            error_data['_metadata'] = {
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'accuracy_percentage': "0%",
                'fields_detected': "0/15",
                'quality_indicator': "Error",
                'error_details': str(e),
                'extraction_method': 'Failed',
                'processing_status': 'Error'
            }
            
            # For backward compatibility with safe defaults
            error_data['Timestamp'] = error_data['_metadata'].get('timestamp', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            error_data['Accuracy'] = error_data['_metadata'].get('accuracy_percentage', '0%')
            error_data['Fields Found'] = error_data['_metadata'].get('fields_detected', '0/15')
            
            return error_data, f"Error dalam ekstraksi: {str(e)}"
    
    def extract_field_value_with_confidence(self, text, field_name):
        """Ekstrak nilai field dengan confidence score"""
        if field_name not in self.patterns:
            return "Tidak terdeteksi", 0.0
        
        # Special handling untuk RT/RW
        if field_name == "RT RW":
            value = self.extract_rt_rw(text)
            confidence = 0.8 if value != "Tidak terdeteksi" else 0.0
            return value, confidence
        
        best_value = "Tidak terdeteksi"
        max_confidence = 0.0
        
        # Try each pattern and calculate confidence
        for i, pattern in enumerate(self.patterns[field_name]):
            try:
                matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
                for match in matches:
                    if isinstance(match, tuple):
                        value = match[0].strip() if match[0] else ""
                    else:
                        value = match.strip()
                    
                    # Validate and get confidence
                    cleaned_value = self.validate_field_value(field_name, value)
                    if cleaned_value and cleaned_value != "Tidak terdeteksi":
                        # Calculate confidence based on pattern index and value quality
                        pattern_confidence = max(0.9 - (i * 0.1), 0.3)  # First patterns are more reliable
                        value_confidence = self._calculate_value_confidence(field_name, cleaned_value, value)
                        total_confidence = (pattern_confidence + value_confidence) / 2
                        
                        if total_confidence > max_confidence:
                            best_value = cleaned_value
                            max_confidence = total_confidence
                        
            except Exception:
                continue
        
        # Fallback to keyword search if no pattern match
        if best_value == "Tidak terdeteksi":
            fallback_value = self.improved_keyword_search(text, field_name)
            if fallback_value != "Tidak terdeteksi":
                best_value = fallback_value
                max_confidence = 0.4  # Lower confidence for keyword search
        
        return best_value, round(max_confidence, 2)
    
    def _calculate_value_confidence(self, field_name, cleaned_value, original_value):
        """Hitung confidence score berdasarkan kualitas value"""
        if not cleaned_value or cleaned_value == "Tidak terdeteksi":
            return 0.0
        
        base_confidence = 0.5
        
        # Field-specific confidence calculation
        if field_name == "NIK":
            if len(cleaned_value) == 16 and cleaned_value.isdigit():
                return 0.95
            elif len(cleaned_value) >= 14:
                return 0.7
            else:
                return 0.3
                
        elif field_name in ["Provinsi", "Kabupaten", "Kel Desa", "Kecamatan"]:
            if len(cleaned_value) >= 5 and cleaned_value.replace(' ', '').isalpha():
                return 0.9
            elif len(cleaned_value) >= 3:
                return 0.7
            else:
                return 0.4
                
        elif field_name == "Nama":
            if 5 <= len(cleaned_value) <= 40 and cleaned_value.replace(' ', '').isalpha():
                return 0.9
            elif len(cleaned_value) >= 3:
                return 0.6
            else:
                return 0.3
                
        elif field_name == "Jenis Kelamin":
            if cleaned_value in ["LAKI-LAKI", "PEREMPUAN"]:
                return 0.95
            else:
                return 0.3
                
        elif field_name == "Gol Darah":
            valid_types = ['A', 'B', 'AB', 'O', 'A+', 'B+', 'AB+', 'O+', 'A-', 'B-', 'AB-', 'O-']
            if cleaned_value in valid_types:
                return 0.9
            else:
                return 0.4
                
        elif field_name == "Agama":
            valid_agama = ['ISLAM', 'KRISTEN', 'KATOLIK', 'HINDU', 'BUDDHA', 'KONGHUCU']
            if cleaned_value in valid_agama:
                return 0.95
            else:
                return 0.3
        
        # General confidence based on value quality
        if len(cleaned_value) >= 5:
            base_confidence += 0.3
        if cleaned_value == original_value.strip():  # No cleaning needed
            base_confidence += 0.2
        if len(cleaned_value.split()) > 1:  # Multiple words
            base_confidence += 0.1
            
        return min(base_confidence, 1.0)
    
    def _determine_quality_indicator(self, accuracy, fields_found):
        """Tentukan indikator kualitas ekstraksi"""
        if accuracy >= 80 and fields_found >= 10:
            return "Excellent ⭐⭐⭐⭐⭐"
        elif accuracy >= 70 and fields_found >= 8:
            return "Very Good ⭐⭐⭐⭐"
        elif accuracy >= 60 and fields_found >= 6:
            return "Good ⭐⭐⭐"
        elif accuracy >= 40 and fields_found >= 4:
            return "Fair ⭐⭐"
        elif fields_found >= 2:
            return "Poor ⭐"
        else:
            return "Failed ❌"
    
    def _create_empty_result(self):
        """Buat hasil kosong dengan struktur yang konsisten"""
        field_order = [
            "Provinsi", "Kabupaten", "NIK", "Nama", 
            "Tempat Tgl Lahir", "Jenis Kelamin", "Gol Darah",
            "Alamat", "RT RW", "Kel Desa", "Kecamatan",
            "Agama", "Status Perkawinan", "Pekerjaan", "Kewarganegaraan"
        ]
        
        return {field: "Tidak terdeteksi" for field in field_order}

def safe_get_value(data_dict, key, default="0%"):
    """Safely get value from dictionary with default fallback"""
    try:
        return data_dict.get(key, default)
    except (AttributeError, KeyError):
        return default

def main():
    """Main function untuk menjalankan KTP OCR Dashboard"""
    # Header
    st.markdown('<h1 class="main-header">🆔 KTP OCR Dashboard</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Sistem Ekstraksi Data KTP Indonesia dengan Enhanced Detection + Image Cropping</p>', unsafe_allow_html=True)
    
    # Info tambahan
    st.markdown("""
    <div class="info-msg">
        ℹ <strong>Enhanced Detection Features:</strong> 
        <ul>
            <li>✂ <strong>Image Cropping:</strong> Crop area spesifik sebelum ekstraksi untuk akurasi maksimal</li>
            <li>🎯 <strong>Fokus Bagian Atas:</strong> Prioritas pada area header KTP untuk akurasi maksimal</li>
            <li>🧹 <strong>Advanced Noise Reduction:</strong> Bilateral filtering & morphological operations</li>
            <li>✅ <strong>Field Validation:</strong> Validasi ketat untuk setiap jenis data KTP</li>
            <li>📐 <strong>Smart RT/RW Detection:</strong> Deteksi otomatis format garis pemisah (001/002)</li>
            <li>🔤 <strong>Character Whitelist:</strong> Filter karakter yang relevan untuk KTP</li>
            <li>🇮🇩 <strong>Indonesian Optimized:</strong> Konfigurasi khusus bahasa Indonesia</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'extracted_records' not in st.session_state:
        st.session_state.extracted_records = []
    if 'current_data' not in st.session_state:
        st.session_state.current_data = {}
    if 'original_image' not in st.session_state:
        st.session_state.original_image = None
    if 'cropped_image' not in st.session_state:
        st.session_state.cropped_image = None
    if 'crop_mode' not in st.session_state:
        st.session_state.crop_mode = False
    
    # Sidebar
    with st.sidebar:
        st.header("📋 Control Panel")
        
        # Statistics
        total_records = len(st.session_state.extracted_records)
        st.markdown(f"""
        <div class="metric-card">
            <h3>{total_records}</h3>
            <p>Total Records</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.current_data:
            accuracy = safe_get_value(st.session_state.current_data, 'Accuracy', '0%')
            fields_found = safe_get_value(st.session_state.current_data, 'Fields Found', '0/15')
            
            st.markdown(f"""
            <div class="metric-card">
                <h3>{accuracy}</h3>
                <p>Detection Accuracy</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="metric-card">
                <h3>{fields_found}</h3>
                <p>Fields Detected</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Crop mode toggle
        if st.session_state.original_image is not None:
            crop_enabled = st.toggle(
                "✂ Enable Cropping Mode",
                value=st.session_state.crop_mode,
                help="Aktifkan untuk crop gambar sebelum ekstraksi"
            )
            st.session_state.crop_mode = crop_enabled
            
            if crop_enabled:
                st.markdown('<div class="warning-msg">🔄 Mode cropping aktif. Atur area crop kemudian lakukan ekstraksi.</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="info-msg">📷 Mode normal aktif. Ekstraksi akan menggunakan seluruh gambar.</div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Clear all data button
        if st.button("🗑 Clear All Data", type="secondary", use_container_width=True):
            st.session_state.extracted_records = []
            st.session_state.current_data = {}
            st.session_state.original_image = None
            st.session_state.cropped_image = None
            st.session_state.crop_mode = False
            st.success("✅ Semua data telah dihapus!")
            st.rerun()
        
        # Tips section
        st.markdown("---")
        st.markdown("### 💡 Tips untuk Hasil Optimal")
        st.markdown("""
        - ✂ Gunakan fitur cropping untuk fokus pada area tertentu
        - 📷 Pastikan foto KTP tidak terlalu miring
        - 💡 Cahaya cukup pada bagian atas KTP
        - 🔍 Resolusi gambar minimal 800px
        - 📱 Hindari bayangan pada teks
        """)
    
    # Main content
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📤 Upload & Crop Foto KTP")
        
        uploaded_file = st.file_uploader(
            "Pilih foto KTP",
            type=['jpg', 'jpeg', 'png', 'webp'],
            help="Format yang didukung: JPG, PNG, WEBP (Max: 10MB)",
            accept_multiple_files=False
        )
        
        if uploaded_file is not None:
            # Load and store original image
            original_image = Image.open(uploaded_file)
            st.session_state.original_image = original_image
            
            # Image info
            width, height = original_image.size
            st.info(f"📊 Dimensi: {width}x{height} pixels | Format: {original_image.format}")
            
            # Cropping interface
            if st.session_state.crop_mode:
                # Initialize cropper
                cropper = ImageCropper()
                
                # Get crop coordinates
                crop_coords = cropper.create_crop_interface(original_image)
                
                # Create preview with overlay
                preview_image = cropper.create_preview_with_overlay(original_image, crop_coords)
                st.image(preview_image, caption="🎯 Preview dengan Area Crop (Area merah = yang akan diekstrak)", use_container_width=True)
                
                # Apply crop button
                col_crop, col_reset = st.columns(2)
                
                with col_crop:
                    if st.button("✂ Apply Crop", type="primary", use_container_width=True):
                        cropped = cropper.crop_image(original_image, crop_coords)
                        if cropped:
                            st.session_state.cropped_image = cropped
                            crop_width, crop_height = cropped.size
                            st.success(f"✅ Cropping berhasil! Dimensi baru: {crop_width}x{crop_height}")
                            st.image(cropped, caption="✂ Hasil Cropping", use_container_width=True)
                        else:
                            st.error("❌ Gagal melakukan cropping. Pastikan area crop valid.")
                
                with col_reset:
                    if st.button("🔄 Reset Crop", type="secondary", use_container_width=True):
                        st.session_state.cropped_image = None
                        st.success("✅ Crop direset!")
                        st.rerun()
            
            else:
                # Normal mode - show original image
                st.image(original_image, caption="📷 Foto KTP Asli", use_container_width=True)
        
        # Extract button
        if st.session_state.original_image is not None:
            # Determine which image to use for extraction
            image_to_extract = st.session_state.cropped_image if st.session_state.crop_mode and st.session_state.cropped_image is not None else st.session_state.original_image
            
            extract_button_text = "🔍 Extract Data dari Crop" if st.session_state.crop_mode and st.session_state.cropped_image is not None else "🔍 Extract Data KTP"
            
            if st.button(extract_button_text, type="primary", use_container_width=True):
                with st.spinner('🔄 Memproses gambar dan mengekstrak data...'):
                    try:
                        # Initialize extractor
                        extractor = KTPExtractor()
                        
                        # Extract data
                        extracted_data, full_text = extractor.extract_ktp_data(image_to_extract)
                        
                        # Store in session state
                        st.session_state.current_data = extracted_data
                        
                        # Add to records
                        record_with_id = extracted_data.copy()
                        record_with_id['ID'] = len(st.session_state.extracted_records) + 1
                        st.session_state.extracted_records.append(record_with_id)
                        
                        # Show success message
                        accuracy = safe_get_value(extracted_data, 'Accuracy', '0%')
                        fields_found = safe_get_value(extracted_data, 'Fields Found', '0/15')
                        
                        st.markdown(f"""
                        <div class="success-msg">
                            ✅ <strong>Ekstraksi berhasil!</strong><br>
                            📊 Akurasi: {accuracy} | 📋 Field terdeteksi: {fields_found}
                        </div>
                        """, unsafe_allow_html=True)
                        
                    except Exception as e:
                        st.markdown(f"""
                        <div class="error-msg">
                            ❌ <strong>Error dalam ekstraksi:</strong><br>
                            {str(e)}
                        </div>
                        """, unsafe_allow_html=True)

    with col2:
        st.subheader("📊 Hasil Ekstraksi")
        
        if st.session_state.current_data:
            # Quality indicator
            if '_metadata' in st.session_state.current_data:
                quality = st.session_state.current_data['_metadata'].get('quality_indicator', 'Unknown')
                st.markdown(f"""
                <div class="info-msg">
                    <strong>Kualitas Ekstraksi:</strong> {quality}
                </div>
                """, unsafe_allow_html=True)
            
            # Display extracted data in organized format
            st.markdown("### 🏛 Informasi Wilayah")
            col_w1, col_w2 = st.columns(2)
            with col_w1:
                st.text_input("Provinsi", value=st.session_state.current_data.get('Provinsi', ''), disabled=True)
                st.text_input("Kabupaten/Kota", value=st.session_state.current_data.get('Kabupaten', ''), disabled=True)
            with col_w2:
                st.text_input("Kecamatan", value=st.session_state.current_data.get('Kecamatan', ''), disabled=True)
                st.text_input("Kel/Desa", value=st.session_state.current_data.get('Kel Desa', ''), disabled=True)
            
            st.markdown("### 👤 Informasi Pribadi")
            st.text_input("NIK", value=st.session_state.current_data.get('NIK', ''), disabled=True)
            st.text_input("Nama", value=st.session_state.current_data.get('Nama', ''), disabled=True)
            
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.text_input("Tempat, Tgl Lahir", value=st.session_state.current_data.get('Tempat Tgl Lahir', ''), disabled=True)
                st.text_input("Jenis Kelamin", value=st.session_state.current_data.get('Jenis Kelamin', ''), disabled=True)
            with col_p2:
                st.text_input("Golongan Darah", value=st.session_state.current_data.get('Gol Darah', ''), disabled=True)
                st.text_input("Agama", value=st.session_state.current_data.get('Agama', ''), disabled=True)
            
            st.markdown("### 🏠 Informasi Alamat")
            st.text_area("Alamat", value=st.session_state.current_data.get('Alamat', ''), disabled=True, height=60)
            
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                st.text_input("RT/RW", value=st.session_state.current_data.get('RT RW', ''), disabled=True)
            with col_a2:
                st.text_input("Status Perkawinan", value=st.session_state.current_data.get('Status Perkawinan', ''), disabled=True)
            
            st.markdown("### 💼 Informasi Lainnya")
            col_l1, col_l2 = st.columns(2)
            with col_l1:
                st.text_input("Pekerjaan", value=st.session_state.current_data.get('Pekerjaan', ''), disabled=True)
            with col_l2:
                st.text_input("Kewarganegaraan", value=st.session_state.current_data.get('Kewarganegaraan', ''), disabled=True)
            
            # Edit mode toggle
            st.markdown("---")
            if st.toggle("✏ Mode Edit", help="Aktifkan untuk mengedit data hasil ekstraksi"):
                st.markdown("### ✏ Edit Data")
                with st.form("edit_form"):
                    edited_data = {}
                    
                    # Editable fields
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        edited_data['NIK'] = st.text_input("NIK *", value=st.session_state.current_data.get('NIK', ''))
                        edited_data['Nama'] = st.text_input("Nama *", value=st.session_state.current_data.get('Nama', ''))
                        edited_data['Jenis Kelamin'] = st.selectbox("Jenis Kelamin", 
                            options=['LAKI-LAKI', 'PEREMPUAN'], 
                            index=0 if st.session_state.current_data.get('Jenis Kelamin', '') == 'LAKI-LAKI' else 1)
                        edited_data['Agama'] = st.selectbox("Agama", 
                            options=['ISLAM', 'KRISTEN', 'KATOLIK', 'HINDU', 'BUDDHA', 'KONGHUCU'],
                            index=0)
                    
                    with col_e2:
                        edited_data['Tempat Tgl Lahir'] = st.text_input("Tempat, Tgl Lahir", 
                            value=st.session_state.current_data.get('Tempat Tgl Lahir', ''))
                        edited_data['Gol Darah'] = st.selectbox("Golongan Darah",
                            options=['A', 'B', 'AB', 'O', 'A+', 'B+', 'AB+', 'O+', 'A-', 'B-', 'AB-', 'O-'],
                            index=0)
                        edited_data['Status Perkawinan'] = st.selectbox("Status Perkawinan",
                            options=['BELUM KAWIN', 'KAWIN', 'CERAI HIDUP', 'CERAI MATI'],
                            index=0)
                        edited_data['Kewarganegaraan'] = st.selectbox("Kewarganegaraan",
                            options=['WNI', 'WNA'],
                            index=0)
                    
                    edited_data['Alamat'] = st.text_area("Alamat", 
                        value=st.session_state.current_data.get('Alamat', ''))
                    
                    col_e3, col_e4, col_e5 = st.columns(3)
                    with col_e3:
                        edited_data['RT RW'] = st.text_input("RT/RW", 
                            value=st.session_state.current_data.get('RT RW', ''))
                    with col_e4:
                        edited_data['Pekerjaan'] = st.text_input("Pekerjaan", 
                            value=st.session_state.current_data.get('Pekerjaan', ''))
                    with col_e5:
                        edited_data['Provinsi'] = st.text_input("Provinsi", 
                            value=st.session_state.current_data.get('Provinsi', ''))
                    
                    col_e6, col_e7, col_e8 = st.columns(3)
                    with col_e6:
                        edited_data['Kabupaten'] = st.text_input("Kabupaten/Kota", 
                            value=st.session_state.current_data.get('Kabupaten', ''))
                    with col_e7:
                        edited_data['Kecamatan'] = st.text_input("Kecamatan", 
                            value=st.session_state.current_data.get('Kecamatan', ''))
                    with col_e8:
                        edited_data['Kel Desa'] = st.text_input("Kel/Desa", 
                            value=st.session_state.current_data.get('Kel Desa', ''))
                    
                    if st.form_submit_button("💾 Simpan Perubahan", type="primary", use_container_width=True):
                        # Update current data
                        for key, value in edited_data.items():
                            st.session_state.current_data[key] = value
                        
                        # Update in records
                        if st.session_state.extracted_records:
                            st.session_state.extracted_records[-1].update(edited_data)
                        
                        st.success("✅ Data berhasil diperbarui!")
                        st.rerun()
        
        else:
            st.markdown("""
            <div class="info-msg">
                📋 <strong>Belum ada data yang diekstrak.</strong><br>
                Upload foto KTP dan klik tombol "Extract Data" untuk memulai.
            </div>
            """, unsafe_allow_html=True)

    # Records history section
    if st.session_state.extracted_records:
        st.markdown("---")
        st.subheader("📚 Riwayat Ekstraksi")
        
        # Convert to DataFrame for better display
        df_records = pd.DataFrame(st.session_state.extracted_records)
        
        # Reorder columns for better presentation
        display_columns = ['ID', 'Timestamp', 'Accuracy', 'Fields Found', 'NIK', 'Nama', 
                          'Tempat Tgl Lahir', 'Jenis Kelamin', 'Alamat', 'Provinsi', 'Kabupaten']
        
        # Only show columns that exist in the dataframe
        available_columns = [col for col in display_columns if col in df_records.columns]
        df_display = df_records[available_columns]
        
        # Display options
        col_opt1, col_opt2, col_opt3 = st.columns(3)
        
        with col_opt1:
            show_all_fields = st.toggle("📋 Tampilkan Semua Field", help="Tampilkan semua field atau hanya yang penting")
        
        with col_opt2:
            records_per_page = st.selectbox("Records per halaman", options=[5, 10, 20, 50], index=1)
        
        with col_opt3:
            search_term = st.text_input("🔍 Cari record", placeholder="Cari berdasarkan NIK/Nama...")
        
        # Filter records based on search
        if search_term:
            mask = df_records.apply(lambda x: x.astype(str).str.contains(search_term, case=False, na=False).any(), axis=1)
            df_display = df_records[mask][available_columns] if not show_all_fields else df_records[mask]
        elif show_all_fields:
            df_display = df_records
        
        # Pagination
        total_records = len(df_display)
        total_pages = max(1, (total_records + records_per_page - 1) // records_per_page)
        
        if total_pages > 1:
            col_page1, col_page2, col_page3 = st.columns([1, 2, 1])
            with col_page2:
                current_page = st.selectbox(
                    f"Halaman (Total: {total_pages})",
                    options=list(range(1, total_pages + 1)),
                    index=0
                )
        else:
            current_page = 1
        
        # Calculate pagination
        start_idx = (current_page - 1) * records_per_page
        end_idx = min(start_idx + records_per_page, total_records)
        
        if total_records > 0:
            df_page = df_display.iloc[start_idx:end_idx]
            
            # Display dataframe
            st.dataframe(
                df_page,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "ID": st.column_config.NumberColumn("ID", width="small"),
                    "Timestamp": st.column_config.DatetimeColumn("Waktu", width="medium"),
                    "Accuracy": st.column_config.TextColumn("Akurasi", width="small"),
                    "Fields Found": st.column_config.TextColumn("Field", width="small"),
                    "NIK": st.column_config.TextColumn("NIK", width="medium"),
                    "Nama": st.column_config.TextColumn("Nama", width="medium")
                }
            )
            
            st.info(f"📊 Menampilkan {len(df_page)} dari {total_records} record(s)")
        else:
            st.info("🔍 Tidak ada record yang sesuai dengan pencarian.")
        
        # Export options
        st.markdown("---")
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        
        with col_exp1:
            if st.button("📥 Export ke CSV", type="secondary", use_container_width=True):
                csv_data = df_records.to_csv(index=False)
                st.download_button(
                    label="⬇ Download CSV",
                    data=csv_data,
                    file_name=f"ktp_extraction_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        with col_exp2:
            if st.button("📊 Export ke Excel", type="secondary", use_container_width=True):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_records.to_excel(writer, sheet_name='KTP_Data', index=False)
                excel_data = output.getvalue()
                
                st.download_button(
                    label="⬇ Download Excel",
                    data=excel_data,
                    file_name=f"ktp_extraction_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        with col_exp3:
            if st.button("🗑 Hapus Semua Record", type="secondary", use_container_width=True):
                if st.button("⚠ Konfirmasi Hapus", type="primary"):
                    st.session_state.extracted_records = []
                    st.success("✅ Semua record telah dihapus!")
                    st.rerun()

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #7f8c8d; padding: 2rem;">
        <p>🆔 <strong>KTP OCR Dashboard v2.0</strong></p>
        <p>Enhanced dengan Image Cropping, Advanced Detection & Validation</p>
        <p><em>Optimized for Indonesian KTP Format</em></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()