"""
File utilities - Helper functions untuk file operations
"""
import logging
import shutil
from pathlib import Path
from datetime import datetime

def create_backup(file_path, backup_dir=None):
    """
    Membuat backup file
    
    Args:
        file_path (Path): Path ke file yang akan di-backup
        backup_dir (Path, optional): Direktori backup
        
    Returns:
        Path: Path ke file backup atau None jika gagal
    """
    logger = logging.getLogger(__name__)
    
    try:
        file_path = Path(file_path)
        
        if not file_path.exists():
            logger.error(f"File tidak ditemukan: {file_path}")
            return None
        
        # Default backup directory
        if backup_dir is None:
            backup_dir = file_path.parent / "backups"
        
        backup_dir = Path(backup_dir)
        backup_dir.mkdir(parents=True, exist_ok=True)
        
        # Nama file backup dengan timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        backup_path = backup_dir / backup_name
        
        # Copy file
        shutil.copy2(file_path, backup_path)
        logger.info(f"Backup dibuat: {backup_path}")
        
        return backup_path
        
    except Exception as e:
        logger.error(f"Error membuat backup: {str(e)}")
        return None

def clean_old_files(directory, days_old=7, pattern="*"):
    """
    Membersihkan file lama dari direktori
    
    Args:
        directory (Path): Direktori target
        days_old (int): File lebih lama dari berapa hari
        pattern (str): Pattern file yang akan dihapus
        
    Returns:
        int: Jumlah file yang dihapus
    """
    logger = logging.getLogger(__name__)
    
    try:
        directory = Path(directory)
        
        if not directory.exists():
            logger.warning(f"Direktori tidak ditemukan: {directory}")
            return 0
        
        # Hitung waktu cutoff
        from datetime import datetime, timedelta
        cutoff_time = datetime.now() - timedelta(days=days_old)
        
        deleted_count = 0
        
        for file_path in directory.glob(pattern):
            if file_path.is_file():
                # Cek waktu modifikasi file
                file_time = datetime.fromtimestamp(file_path.stat().st_mtime)
                
                if file_time < cutoff_time:
                    try:
                        file_path.unlink()
                        logger.info(f"File lama dihapus: {file_path.name}")
                        deleted_count += 1
                    except Exception as e:
                        logger.error(f"Gagal menghapus {file_path.name}: {str(e)}")
        
        logger.info(f"Pembersihan selesai: {deleted_count} file dihapus")
        return deleted_count
        
    except Exception as e:
        logger.error(f"Error membersihkan file lama: {str(e)}")
        return 0

def get_file_info(file_path):
    """
    Mendapatkan informasi detail file
    
    Args:
        file_path (Path): Path ke file
        
    Returns:
        dict: Informasi file atau None jika gagal
    """
    logger = logging.getLogger(__name__)
    
    try:
        file_path = Path(file_path)
        
        if not file_path.exists():
            return None
        
        stat = file_path.stat()
        
        info = {
            'name': file_path.name,
            'size': stat.st_size,
            'size_mb': round(stat.st_size / (1024 * 1024), 2),
            'created': datetime.fromtimestamp(stat.st_ctime),
            'modified': datetime.fromtimestamp(stat.st_mtime),
            'extension': file_path.suffix,
            'is_file': file_path.is_file(),
            'is_dir': file_path.is_dir()
        }
        
        return info
        
    except Exception as e:
        logger.error(f"Error mendapatkan info file: {str(e)}")
        return None