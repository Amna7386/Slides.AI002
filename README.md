# AI Presentation Generator

A beginner-friendly Streamlit app that creates editable PowerPoint presentations.

## Features
- Input by topic, pasted notes, uploaded PDF/TXT/Markdown/DOCX, or a combination
- Exports editable PowerPoint (`.pptx`) plus a text outline
- Professional blue/teal design
- Works without an API key using a built-in structured generator
- Optional Google Gemini API integration; use only if you have access to its free tier

## Run locally (Windows)
1. Install Python 3.10 or newer from https://www.python.org/downloads/
2. Open Command Prompt in this folder.
3. Create a virtual environment:
   `python -m venv .venv`
4. Activate it:
   `.venv\Scripts\activate`
5. Install packages:
   `pip install -r requirements.txt`
6. Start the app:
   `streamlit run app.py`
7. Open the local URL shown in the terminal, usually http://localhost:8501

## Quick test
1. Choose Topic only.
2. Enter `Generative AI in Education`.
3. Keep 7 slides and click Generate presentation.
4. Download the PowerPoint and open it in Microsoft PowerPoint, LibreOffice Impress, or Google Slides.
5. Try Paste notes with 4-5 short sentences, then try uploading a text-based PDF or DOCX.
6. Confirm the downloaded deck has editable titles and bullet text.

## Optional Gemini mode
The app does not require an API key. If you choose to use Gemini:
1. Obtain a Gemini API key through Google's AI developer platform and check current free-tier availability and terms.
2. Paste it into the optional key field inside the app.
3. Check “Use Gemini AI if a key is entered” and generate.
Never publish or commit your API key. Free-tier limits and model availability can change.

## Important limitations
- The no-key generator is a structured template-based generator, not a large language model. It creates a useful starting deck but may need editing.
- Scanned/image-only PDFs do not yield text without OCR.
- Always fact-check AI-generated content and add source citations before presenting.
