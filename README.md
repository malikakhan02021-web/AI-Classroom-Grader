# AI Classroom Grader 🤖🎓

An automated AI Teaching Assistant that connects to Google Classroom, downloads student lab reports, extracts text using OCR (for scanned PDFs and screenshots), and grades them against a teacher's rubric using Groq AI.

## Features
- **Google Classroom Integration:** Fetches submissions and student names automatically using the Google API.
- **OCR Support:** Uses Tesseract OCR to read text from images, screenshots inside Word docs, and scanned PDF pages.
- **AI Grading:** Uses CrewAI and Groq (Llama 3) to act as a strict professor, grading the report out of 100 and providing specific feedback.
- **Excel Output:** Generates a clean `Final_Grades.xlsx` file with student names, grades, and feedback.

## Tech Stack
- **Python 3.12**
- **Google Classroom & Drive APIs**
- **CrewAI & Groq (GPT-OSS-120B)**
- **Tesseract OCR & PyMuPDF**
- **Pandas**

## Prerequisites
1. Python 3.12+
2. Tesseract OCR installed on your system.
3. A Google Cloud Project with the Classroom API and Drive API enabled (`credentials.json`).
4. A free Groq API Key.

## How to Run
1. Install the required packages:
   `py -3.12 -m pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib crewai pandas openpyxl python-docx pymupdf pytesseract pillow`
2. Place your `credentials.json` and `rubric.docx` in the folder.
3. Add your Groq API key to `app.py` (Line 42).
4. Run the script:
   `py -3.12 app.py`

*Note: Never commit your `credentials.json`, `token.json`, or API keys to GitHub.*
