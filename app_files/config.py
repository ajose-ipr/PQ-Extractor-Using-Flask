import os

class Config:
    """Application configuration"""
    # Flask settings
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-secret-key-change-in-production-please'
    
    # File upload settings
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max file size
    ALLOWED_EXTENSIONS = {'pdf'}
    
    # Create upload folder if it doesn't exist
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    
    # Processing constants (from utils.processing)
    EXPECTED_HARMONICS = set(range(2, 51))
    FAIL_COLOR = 'color: darkred; font-weight: bold; background-color: #ffebeb'
    HIGHLIGHT_COLOR = 'font-weight: bold; background-color: #fff3cd'
    
    # Column definitions for all table types
    VOLTAGE_COLUMNS = [
        "Harmonic", "Time Percent Limit[%]", "Reg Max[%]",
        "Measured_V1N", "Measured_V2N", "Measured_V3N", 
        "Result_V1N", "Result_V2N", "Result_V3N"
    ]
    CURRENT_COLUMNS = [
        "Harmonic", "Time Percent Limit[%]", "Reg Max[%]",
        "Measured_I1", "Measured_I2", "Measured_I3", 
        "Result_I1", "Result_I2", "Result_I3"
    ]
    
    # Section boundaries for table extraction
    SECTION_BOUNDARIES = {
        "HARMONIC VOLTAGE FULL TIME RANGE": [
            "SUMMARY", "TOTAL HARMONIC VOLTAGE FULL TIME RANGE", 
            "TOTAL HARMONIC DISTORTION FULL TIME RANGE", "HARMONIC CURRENT FULL TIME RANGE"
        ],
        "HARMONIC CURRENT FULL TIME RANGE": [
            "TOTAL HARMONIC DISTORTION DAILY", "TDD FULL TIME RANGE",
            "HARMONIC VOLTAGE DAILY", "TRANSIENT"
        ],
        "HARMONIC VOLTAGE DAILY": [
            "TOTAL HARMONIC DISTORTION FULL TIME RANGE", 
            "TOTAL HARMONIC VOLTAGE FULL TIME RANGE", "HARMONIC CURRENT DAILY", 
            "TOTAL HARMONIC DISTORTION DAILY"
        ],
        "HARMONIC CURRENT DAILY": [
            "TDD FULL TIME RANGE", "TDD DAILY", "TRANSIENT", "FLICKER SEVERITY"
        ]
    }
    
    # All supported table names
    SUPPORTED_TABLES = [
        "Harmonic Voltage Full Time Range",
        "Harmonic Current Full Time Range", 
        "Harmonic Voltage Daily",
        "Harmonic Current Daily"
    ]
    
    # Application information
    APP_NAME = "PQ Report Harmonics Extractor"
    APP_VERSION = "2.0.0"
    
    # Logging configuration
    LOG_LEVEL = os.environ.get('LOG_LEVEL', 'INFO')
    LOG_FILE = 'app.log'
    
    # Development settings
    DEBUG = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    
    # Production settings
    HOST = os.environ.get('FLASK_HOST', '0.0.0.0')
    PORT = int(os.environ.get('FLASK_PORT', 5000))