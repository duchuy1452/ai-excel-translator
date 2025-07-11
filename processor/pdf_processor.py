from io import BytesIO
from typing import Callable
from pdf2docx import Converter
import tempfile
import os

from translators import GeminiTranslator
from processor.docx_processor import DocxProcessor # Import the new DocxProcessor
from utils import logger

class PdfProcessor:
    """
    Processes PDF (.pdf) files by converting them to DOCX format first,
    then translating the DOCX content.
    """

    def __init__(self, translator: GeminiTranslator):
        self.translator = translator
        # PdfProcessor now needs a DocxProcessor to handle the actual translation
        self.docx_processor = DocxProcessor(translator)

    def process_pdf(
        self, file_data: BytesIO, target_lang: str, progress_callback: Callable = None
    ) -> BytesIO:
        """
        Converts the input PDF to DOCX, translates the DOCX content,
        and returns the translated DOCX file as a BytesIO stream.

        Args:
            file_data (BytesIO): The input PDF file as a BytesIO stream.
            target_lang (str): The target language for translation.
            progress_callback (Callable, optional): A callback function to track progress
                                                   (will be passed to DocxProcessor).

        Returns:
            BytesIO: The translated content as a DOCX BytesIO stream.
        """
        logger.info("Starting PDF to DOCX conversion...")

        # pdf2docx works with file paths, so save BytesIO to a temporary file
        pdf_temp_file = None
        docx_temp_file = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(file_data.getvalue())
                pdf_temp_file = temp_pdf.name

            # Create a temporary path for the output docx
            docx_temp_file = tempfile.mktemp(suffix=".docx")

            # Perform the conversion
            cv = Converter(pdf_temp_file)
            cv.convert(docx_temp_file, start=0, end=None)
            cv.close()
            logger.info("PDF to DOCX conversion successful.")

            # Read the converted DOCX file into BytesIO
            with open(docx_temp_file, "rb") as f_docx:
                docx_data = BytesIO(f_docx.read())

            # --- Delegate translation to DocxProcessor ---
            logger.info("Processing converted DOCX file for translation...")
            # Pass the progress callback directly to the docx processor
            translated_docx_stream = self.docx_processor.process_docx(
                docx_data, target_lang, progress_callback
            )

            return translated_docx_stream

        except Exception as e:
            logger.error(f"Error during PDF processing (conversion or translation): {e}", exc_info=True)
            # Consider more specific error handling for conversion vs. translation
            raise RuntimeError(f"Failed to process PDF. Conversion or translation failed. Error: {e}") from e

        finally:
            # Clean up temporary files
            if pdf_temp_file and os.path.exists(pdf_temp_file):
                os.remove(pdf_temp_file)
            if docx_temp_file and os.path.exists(docx_temp_file):
                os.remove(docx_temp_file)