from deep_translator import GoogleTranslator

def translate_text(text: str, target_lang: str = 'pt') -> str:
    """
    Translates text from English to the specified target language (default Portuguese).
    
    Parameters:
        text (str): The text to translate.
        target_lang (str): The target language code ('pt' for Portuguese, 'en' for English, etc.)

    Returns:
        str: Translated text.
    """
    if not text or not isinstance(text, str):
        return text
    try:
        return GoogleTranslator(source='en', target=target_lang).translate(text)
    except Exception as e:
        print(f"Translation error: {e}")
        return text