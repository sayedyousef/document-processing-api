# ==============================================================
# FILE 2: backend/core/logger.py
# ==============================================================
"""
Logging setup for the document processing service
"""

import logging
import sys
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Optional

# Import config
from .config import Config

class ColoredFormatter(logging.Formatter):
    """Custom formatter with colors for console output"""
    
    grey = "\x1b[38;21m"
    yellow = "\x1b[33;21m"
    red = "\x1b[31;21m"
    bold_red = "\x1b[31;1m"
    green = "\x1b[32;21m"
    reset = "\x1b[0m"
    
    COLORS = {
        logging.DEBUG: grey,
        logging.INFO: green,
        logging.WARNING: yellow,
        logging.ERROR: red,
        logging.CRITICAL: bold_red
    }
    
    def format(self, record):
        log_color = self.COLORS.get(record.levelno, self.grey)
        record.levelname = f"{log_color}{record.levelname}{self.reset}"
        return super().format(record)

def setup_logger(
    name: str,
    level: Optional[str] = None,
    log_to_file: bool = True,
    log_to_console: bool = True
) -> logging.Logger:
    """
    Setup a logger with file and console handlers
    
    Args:
        name: Logger name (usually __name__)
        level: Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_to_file: Whether to log to file
        log_to_console: Whether to log to console
        
    Returns:
        Configured logger instance
    """
    
    # Create logger
    logger = logging.getLogger(name)
    
    # Set level
    log_level = getattr(logging, level or Config.LOG_LEVEL)
    logger.setLevel(log_level)
    
    # Remove existing handlers to avoid duplicates
    logger.handlers.clear()
    
    # File handler with rotation
    if log_to_file:
        log_filename = f"{Config.LOG_FILE_PREFIX}_{datetime.now().strftime('%Y%m%d')}.log"
        log_path = Config.LOGS_DIR / log_filename
        
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=10 * 1024 * 1024,  # 10MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(log_level)
        
        # File formatter (detailed)
        file_formatter = logging.Formatter(
            Config.LOG_FORMAT,
            datefmt=Config.LOG_DATE_FORMAT
        )
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)
    
    # Console handler
    if log_to_console:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(log_level)
        
        # Console formatter (with colors)
        if sys.stdout.isatty():  # Check if running in terminal
            console_formatter = ColoredFormatter(
                Config.LOG_FORMAT,
                datefmt=Config.LOG_DATE_FORMAT
            )
        else:
            console_formatter = logging.Formatter(
                Config.LOG_FORMAT,
                datefmt=Config.LOG_DATE_FORMAT
            )
        
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)
    
    return logger

def get_logger(name: str) -> logging.Logger:
    """
    Get or create a logger with default settings
    
    Args:
        name: Logger name
        
    Returns:
        Logger instance
    """
    return setup_logger(name)

# Create a default logger for the module
logger = setup_logger(__name__)

# Utility functions for common logging patterns

def log_function_call(func):
    """Decorator to log function calls"""
    def wrapper(*args, **kwargs):
        logger.debug(f"Calling {func.__name__} with args={args}, kwargs={kwargs}")
        try:
            result = func(*args, **kwargs)
            logger.debug(f"{func.__name__} returned successfully")
            return result
        except Exception as e:
            logger.error(f"{func.__name__} failed with error: {e}")
            raise
    return wrapper

def log_processing_stats(job_id: str, stats: dict):
    """Log processing statistics"""
    logger.info(f"Job {job_id} statistics:")
    for key, value in stats.items():
        logger.info(f"  {key}: {value}")

def log_error_with_context(error: Exception, context: dict):
    """Log error with additional context"""
    logger.error(f"Error: {str(error)}")
    logger.error("Context:")
    for key, value in context.items():
        logger.error(f"  {key}: {value}")

# Example usage in other modules:
"""
from core.logger import get_logger

logger = get_logger(__name__)

logger.info("Starting document processing")
logger.debug("Debug information")
logger.warning("Warning message")
logger.error("Error occurred")
"""