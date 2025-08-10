"""
FILE 1: backend/core/config.py
Configuration settings for the document processing service
"""

from pathlib import Path
import os

class Config:
    """Configuration settings for document processing"""
    
    # Base paths
    BASE_DIR = Path(__file__).parent.parent
    ROOT_DIR = BASE_DIR.parent
    
    # Directory paths
    TEMP_DIR = BASE_DIR / "temp"
    OUTPUT_DIR = BASE_DIR / "output"
    LOGS_DIR = BASE_DIR / "logs"
    DOCUMENTS_DIR = ROOT_DIR / "documents"
    INPUT_DIR = DOCUMENTS_DIR / "input"
    PROCESSED_DIR = DOCUMENTS_DIR / "processed"
    
    # File settings
    MAX_FILE_SIZE = 25 * 1024 * 1024  # 25MB
    ALLOWED_EXTENSIONS = ['.docx', '.doc']
    MAX_FILES_PER_REQUEST = 10
    
    # API settings
    API_HOST = "0.0.0.0"
    API_PORT = 8000
    API_TITLE = "Document Processing Service"
    API_VERSION = "1.0.0"
    
    # Processing settings
    BATCH_SIZE = 5  # Process 5 files at a time
    JOB_TIMEOUT = 600  # 10 minutes timeout
    CLEANUP_AFTER_HOURS = 24  # Clean temp files after 24 hours
    
    # Excel output settings
    EXCEL_ENGINE = 'openpyxl'
    EXCEL_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
    
    # HTML conversion settings
    HTML_ENCODING = 'utf-8'
    HTML_DIR = 'rtl'  # For Arabic documents
    HTML_LANG = 'ar'
    
    # Logging settings
    LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
    LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
    LOG_FILE_PREFIX = 'document_processing'
    
    # Mammoth settings (for Word to HTML conversion)
    MAMMOTH_STYLE_MAP = """
        p[style-name='Heading 1'] => h1:fresh
        p[style-name='Heading 2'] => h2:fresh
        p[style-name='Heading 3'] => h3:fresh
        p[style-name='Heading 4'] => h4:fresh
        p[style-name='Title'] => h1.title
        p[style-name='Subtitle'] => h2.subtitle
        p[style-name='Quote'] => blockquote
        p[style-name='Code'] => pre.code
    """
    
    # Analysis settings
    MIN_WORD_COUNT_FOR_VALID = 100
    MIN_SECTIONS_FOR_VALID = 1
    
    @classmethod
    def ensure_directories(cls):
        """Create all required directories if they don't exist"""
        directories = [
            cls.TEMP_DIR,
            cls.OUTPUT_DIR,
            cls.LOGS_DIR,
            cls.DOCUMENTS_DIR,
            cls.INPUT_DIR,
            cls.PROCESSED_DIR
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
    
    @classmethod
    def get_temp_path(cls, job_id: str) -> Path:
        """Get temp directory path for a specific job"""
        path = cls.TEMP_DIR / job_id
        path.mkdir(parents=True, exist_ok=True)
        return path
    
    @classmethod
    def get_output_path(cls, job_id: str) -> Path:
        """Get output directory path for a specific job"""
        path = cls.OUTPUT_DIR / job_id
        path.mkdir(parents=True, exist_ok=True)
        return path
    
    @classmethod
    def is_allowed_file(cls, filename: str) -> bool:
        """Check if file extension is allowed"""
        return any(filename.lower().endswith(ext) for ext in cls.ALLOWED_EXTENSIONS)
    
    @classmethod
    def get_file_size_mb(cls, size_bytes: int) -> float:
        """Convert bytes to megabytes"""
        return size_bytes / (1024 * 1024)

# Ensure all directories exist when module is imported
Config.ensure_directories()

