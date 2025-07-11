import streamlit as st
from dotenv import load_dotenv
import os
import time
from translators import GeminiTranslator
from processor.excel_processor import ExcelProcessor
from processor.powerpoint_processor import PowerPointProcessor
from processor.pdf_processor import PdfProcessor  # Add PDF processor import
from processor.docx_processor import DocxProcessor
from constants import SUPPORTED_LANGUAGES
from utils import logger


def initialize_app(file_description):
    load_dotenv()
    api_key = os.getenv("GEMINI_API_KEY")
    translator = GeminiTranslator(api_key, file_description)
    return translator


def update_excel_progress(
    sheet_name: str, processed: int, total: int, sheet_idx: int, no_sheet: int
):
    progress = processed / total if total > 0 else 0
    st.write(
        f"Processing sheet [{sheet_idx + 1}/{no_sheet}] {sheet_name}: {progress:.2%} complete"
    )
    logger.info(
        f"Processing sheet [{sheet_idx + 1}/{no_sheet}] {sheet_name}: {progress:.2%} complete"
    )


def update_powerpoint_progress(processed: int, total: int):
    progress = processed / total if total > 0 else 0
    st.write(f"Processing slide: {progress:.2%} complete")
    logger.info(f"Processing slide: {progress:.2%} complete")


def update_pdf_progress(processed_batches: int, total_batches: int):
    # Progress is now based on batches processed by DocxProcessor after PDF->DOCX conversion
    # Note: Conversion time is not explicitly tracked here, progress starts during translation.
    if total_batches > 0:
        progress = processed_batches / total_batches
        st.write(f"Translating content (batch {processed_batches}/{total_batches}): {progress:.1%} complete")
        logger.info(f"Translating content (batch {processed_batches}/{total_batches}): {progress:.1%} complete")
    else:
        st.write("Preparing translation...") # Or handle zero batches case
        logger.info("Preparing translation (0 batches)...")


def update_docx_progress(processed_batches: int, total_batches: int):
    # Progress is now based on batches processed by DocxProcessor after PDF->DOCX conversion
    # Note: Conversion time is not explicitly tracked here, progress starts during translation.
    if total_batches > 0:
        progress = processed_batches / total_batches
        st.write(f"Translating content (batch {processed_batches}/{total_batches}): {progress:.1%} complete")
        logger.info(f"Translating content (batch {processed_batches}/{total_batches}): {progress:.1%} complete")
    else:
        st.write("Preparing translation...") # Or handle zero batches case
        logger.info("Preparing translation (0 batches)...")

def main():
    st.set_page_config(
        page_title="Gemibara - Translate files",
        page_icon="./assets/gemibara.png",
        layout="centered",
    )

    logo_col, title_col = st.columns([1, 7])

    with logo_col:
        st.image("./assets/gemibara.png", width=80)

    with title_col:
        st.title("Gemibara")

    uploaded_file = st.file_uploader(
        "Choose a file", type=["xlsx", "xls", "pptx", "pdf", "docx"], accept_multiple_files=False, # Add 'pdf' type
    )

    file_description = st.text_area(
        label="Describe your file for better translation (optional):",
        placeholder="For example: This is a test case file for a web application's recruitment module.",
        max_chars=5000,
    )

    target_lang = st.selectbox("Select target language", SUPPORTED_LANGUAGES)

    if uploaded_file:
        file_extension = uploaded_file.name.split(".")[-1].lower()

        if file_extension in ["xlsx", "xls"]:
            import pandas as pd

            try:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                selected_sheets = st.multiselect(
                    "Select sheets to translate", sheet_names, default=sheet_names
                )

                if len(selected_sheets) == 1:
                    cell_range = st.text_input(
                        "Enter cell range (optional, A1:B10):", ""
                    )
                else:
                    cell_range = None

            except Exception as e:
                st.error(f"Error reading Excel file: {e}")
                selected_sheets = []
                cell_range = None

        elif file_extension == "pptx":
            selected_sheets = ["PowerPoint"]
            cell_range = None
        elif file_extension == "pdf": # Add handler for PDF
            selected_sheets = ["PDF Content"] # Placeholder, not really used for selection
            cell_range = None
        elif file_extension == "docx":
            selected_sheets = ["Word Document"]
            cell_range = None
        else:
            selected_sheets = []
            cell_range = None
            st.error("Unsupported file type")

        if uploaded_file and st.button("Translate"):
            translator = initialize_app(file_description)

            with st.spinner("Translating..."):
                try:
                    start_time = time.time()
                    if file_extension in ["xlsx", "xls"]:
                        processor = ExcelProcessor(translator)
                        output = processor.process_workbook(
                            uploaded_file,
                            target_lang,
                            selected_sheets,
                            cell_range,
                            update_excel_progress,
                        )
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    elif file_extension == "pptx":
                        processor = PowerPointProcessor(translator)
                        output = processor.process_powerpoint(
                            uploaded_file, target_lang, update_powerpoint_progress
                        )
                        mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    elif file_extension == "pdf": # Add PDF processing logic
                        processor = PdfProcessor(translator)
                        output = processor.process_pdf(
                            uploaded_file, target_lang, update_pdf_progress # Pass the same progress callback
                        )
                        # Output is now DOCX after conversion and translation
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    elif file_extension == "docx":
                        processor = DocxProcessor(translator)
                        output = processor.process_docx(
                            uploaded_file, target_lang, update_docx_progress
                        )
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    else:
                        raise ValueError("Unsupported file type")

                    total_time = int(time.time() - start_time)
                    logger.info(
                        "Translation completed! Total time: %s seconds", total_time
                    )
                    st.success(
                        f"""Translation completed! Total time: {total_time} seconds"""
                    )
                    st.download_button(
                        label="Download translated file",
                        data=output,
                        # Change extension to .docx for PDF input, keep others same
                        file_name=f"({target_lang})_{os.path.splitext(uploaded_file.name)[0]}.docx" if file_extension == 'pdf' else f"({target_lang})_{uploaded_file.name}",
                        mime=mime_type,
                    )
                except Exception as e:
                    logger.error(e)
                    st.error(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    main()
