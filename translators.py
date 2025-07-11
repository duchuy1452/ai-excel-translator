import google.generativeai as genai
import json
from time import sleep
from typing import List
import re

from constants import (
    MAX_RETRIES,
    QUOTA_EXHAUST_ERROR_DELAY,
    RETRY_DELAY,
    REQUEST_DELAY,
    Language,
)
from utils import logger

QUOTA_EXHAUST_ERROR = "429 Resource has been exhausted"


class GeminiTranslator:
    def __init__(self, api_key: str, file_description: str):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel("gemini-2.0-flash")
        self.file_description = file_description
        self.translation_cache = {}  # Add a translation cache
        logger.info("GeminiTranslator initialized.")

    def translate_batch(self, texts: List[str], target_lang: str) -> List[str]:
        """Translate a batch of texts using Gemini API with retry mechanism"""
        if not texts:
            return []

        num_texts = len(texts)
        translations = [None] * num_texts
        # Remove \n, \t at heading and trailing
        texts = [self._preprocess_text(text) for text in texts]

        uncached_indices, uncached_texts = self._get_cached_translations(
            texts, target_lang, translations
        )

        if uncached_texts:
            new_translations = self._call_translation_api(uncached_texts, target_lang)
            if new_translations:
                # Store new translations in cache and update results
                for i, translated_text in enumerate(new_translations):
                    original_index = uncached_indices[i]
                    original_text = uncached_texts[i]
                    self.translation_cache[(original_text, target_lang)] = (
                        translated_text
                    )
                    translations[original_index] = translated_text
            else:
                self._handle_failed_translations(uncached_indices, texts, translations)

        return translations

    def _get_cached_translations(self, texts, target_lang, translations):
        uncached_indices = []
        uncached_texts = []
        for i, text in enumerate(texts):
            if (text, target_lang) in self.translation_cache:
                translations[i] = self.translation_cache[(text, target_lang)]
            else:
                uncached_indices.append(i)
                uncached_texts.append(text)
        return uncached_indices, uncached_texts

    def _call_translation_api(self, texts, target_lang):
        prompt = self._create_translation_prompt(texts, target_lang)
        logger.info(f"Sending translation request for {len(texts)} texts.")
        for attempt in range(MAX_RETRIES):
            try:
                response = self.model.generate_content(
                    prompt,
                    generation_config=genai.GenerationConfig(
                        response_mime_type="application/json",
                        response_schema=list[str],
                        temperature=0.1,
                    ),
                )
                logger.debug(
                    f"Received API response (attempt {attempt + 1}): {response.text}"
                )
                new_translations = json.loads(response.text)
                sleep(REQUEST_DELAY)
                if len(new_translations) != len(texts):
                    error_msg = f"Translation count mismatch: expected {len(texts)}, got {len(new_translations)} (attempt {attempt + 1})"
                    logger.error(error_msg)
                    logger.info("Ouput %s", new_translations)
                    raise ValueError(error_msg)
                logger.info(f"Successfully translated {len(texts)} texts.")
                return new_translations
            except ValueError as e:
                if "Translation count mismatch" in str(e):
                    logger.warning(f"Retrying with smaller batches due to: {e}")
                    midpoint = len(texts) // 2
                    first_half = self._call_translation_api(texts[:midpoint], target_lang)
                    second_half = self._call_translation_api(texts[midpoint:], target_lang)
                    return first_half + second_half
                else:
                    logger.error(f"Translation error: {e}")
                    return [""] * len(texts)
            except Exception as e:
                logger.info("Fail to translate: %s", texts)
                logger.error(
                    f"Error during translation attempt {attempt + 1}/{MAX_RETRIES}: {e}"
                )
                if QUOTA_EXHAUST_ERROR in str(e):
                    sleep(QUOTA_EXHAUST_ERROR_DELAY)
                if attempt == MAX_RETRIES - 1:
                    logger.error(
                        "Max retries reached. Returning None for failed translations."
                    )
                    return None
                sleep(RETRY_DELAY)
        return None

    def _handle_failed_translations(self, uncached_indices, texts, translations):
        logger.error(
            "Max retries reached. Returning original text for failed translations."
        )
        for i in uncached_indices:
            if translations[i] is None:
                translations[i] = texts[i]

    def _preprocess_text(self, text: str):
        text = re.sub(r"\n+", "\n", text)
        text = re.sub(r"\t+", "\t", text).strip()
        return text

    def _create_translation_prompt(self, texts, target_lang: str) -> str:
        json_text = json.dumps(texts, ensure_ascii=False, indent=2)
        return f"""
You are professional translator.
You will receive a JSON string containing a list of {len(texts)} texts.
Your task is to translate all strings in the list to {target_lang}.
{f"""Please translate the texts to fit the context described in the file description below.
File description: {self.file_description}
""" if self.file_description else ""}
Note:
- Maintain any special characters, tab, enter, bullet list, unicode characters, numbers, and formatting within each text segment.
- Ensure that there are no extra newline characters, spaces, tabs, newlines
- If a text segment appears to be programming code, do not translate it.
- The output must be array of {len(texts)} translated texts, and must preserve any duplicate texts.
- Return only a JSON format representing the list of translated texts and can be directly parsed as a Python list of strings.
{"- Must translate texts in the dictionary form (plain form (辞書形 - jishokei))\n"if target_lang == Language.Japanese.value else ""}
Input JSON:
{json_text}
"""
