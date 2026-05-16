from app.i18n.en import TRANSLATIONS as EN_TRANSLATIONS
from app.i18n.th import TRANSLATIONS as TH_TRANSLATIONS


TRANSLATIONS = {
    "English": EN_TRANSLATIONS,
    "ไทย": TH_TRANSLATIONS,
}

THAI_LANGUAGE = "ไทย"


def translate(language: str, key: str) -> str:
    return TRANSLATIONS.get(language, EN_TRANSLATIONS).get(key, key)


def translate_value(language: str, value: object) -> object:
    if not isinstance(value, str):
        return value
    return translate(language, value)
