# Gemibara - Document Translation Suite

A comprehensive tool for translating Excel, Word, PDF, and PowerPoint files while preserving formatting and structure. Powered by Google's Gemini AI API for high-quality multilingual translations.

## Features

- Translate multiple file formats:
  - Excel (.xlsx, .xls): Preserves formulas, formatting, and worksheets
  - Word (.docx): Maintains styles, images, and document structure
  - PDF: Preserves text layout and formatting (limited editing capability)
  - PowerPoint (.pptx): Keeps slide layouts, animations, and embedded media
- Smart detection of translatable content
- Batch processing for large files
- Progress tracking during translation
- Simple web interface built with Streamlit
- Support for 7 languages: Vietnamese, English, French, German, Spanish, Chinese, and Japanese

## Prerequisites

- Python 3.8 or higher
- Google Cloud account with Gemini API access
- Gemini API key

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/gemibara.git
cd gemibara
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root and add your Gemini API key:
```bash
GEMINI_API_KEY=your_api_key_here
```

## Usage

1. Start the application:
```bash
streamlit run app.py
```

2. Open your web browser and navigate to the provided URL (typically http://localhost:8501)

3. Upload an Excel file using the file uploader

4. Select the target language from the dropdown menu

5. Click "Translate" to start the translation process

6. Download the translated file when processing is complete