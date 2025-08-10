#backend/processors/processor_factory.py
# ==============================================================

from .scan_verify_processor import ScanVerifyProcessor
from .word_to_html_processor import WordToHtmlProcessor

def get_processor(processor_type: str):
    """Factory to get the right processor"""
    processors = {
        "scan_verify": ScanVerifyProcessor(),
        "word_to_html": WordToHtmlProcessor()
    }
    
    if processor_type not in processors:
        raise ValueError(f"Unknown processor type: {processor_type}")
    
    return processors[processor_type]