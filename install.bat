@echo off
echo ========================================
echo EasyOCR System - Windows Installer
echo ========================================
echo.

:: Check Python
python --version
if errorlevel 1 (
    echo ERROR: Python tidak ditemukan!
    echo Install Python dari: https://python.org
    pause
    exit /b 1
)

echo.
echo Upgrading pip...
python -m pip install --upgrade pip

echo.
echo Installing packages step by step...

:: Install numpy first
echo [1/6] Installing numpy...
python -m pip install numpy --no-cache-dir
if errorlevel 1 echo WARNING: numpy installation failed

:: Install Pillow
echo [2/6] Installing Pillow...
python -m pip install Pillow --no-cache-dir
if errorlevel 1 echo WARNING: Pillow installation failed

:: Install OpenCV
echo [3/6] Installing opencv-python...
python -m pip install opencv-python --no-cache-dir
if errorlevel 1 echo WARNING: opencv-python installation failed

:: Install PyTorch (CPU version)
echo [4/6] Installing torch...
python -m pip install torch torchvision --index-url https://download.pytorch.org/whl/cpu
if errorlevel 1 echo WARNING: torch installation failed

:: Install EasyOCR
echo [5/6] Installing easyocr...
python -m pip install easyocr --no-cache-dir
if errorlevel 1 echo WARNING: easyocr installation failed

:: Install matplotlib
echo [6/6] Installing matplotlib...
python -m pip install matplotlib --no-cache-dir
if errorlevel 1 echo WARNING: matplotlib installation failed

echo.
echo ========================================
echo Testing installation...
echo ========================================
python install_packages.py

echo.
echo Installation completed!
echo You can now run: python main.py
pause