# API and Request Settings
from enum import Enum


MAX_TOKENS_PER_REQUEST = 400
MAX_RETRIES = 10
RETRY_DELAY = 1
QUOTA_EXHAUST_ERROR_DELAY = 10
REQUEST_DELAY = 0.5
TRANSLATION_DELIMITER = "|//|"

# Supported Languages
SUPPORTED_LANGUAGES = [
    "English",
    "Vietnamese",
    "Japanese",
    "Chinese",
    "French",
    "German",
    "Spanish",
]


class Language(str, Enum):
    English = "English"
    Vietnamese = "Vietnamese"
    Japanese = "Japanese"
    Chinese = "Chinese"
    French = "French"
    German = "German"
    Spanish = "Spanish"


# Token Estimation
CHARS_PER_TOKEN = 4
