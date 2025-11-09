# **NLP Text Smell Analysis**

This project uses **Natural Language Processing (NLP)** and **Google Gemini AI** to find and fix unclear writing patterns such as **passive voice** and **conditional words** (*should, could, might, may, must*).  
It supports multiple file types and creates a **PDF report** with suggested sentence improvements.

---

## **Dependencies**

Install all required libraries:

```bash
pip install spacy python-docx beautifulsoup4 PyMuPDF reportlab pywin32 striprtf google-generativeai
python -m spacy download en_core_web_sm



Note:

pywin32 is only needed on Windows for .doc → .docx conversion.

Set your Gemini API key in the script:

genai.configure(api_key="YOUR_API_KEY_HERE")


How to Run

Place your files inside the req/ folder.
Supported: .pdf, .doc, .docx, .html, .htm, .rtf, .txt

Run the program:

python NLP.py


The report will be saved in the outputs/ folder as:

NLP_Analysis_Report.pdf

Report Details

The generated report includes:

File name and page number

Sentence found

Issue type (Passive / Conditional)

AI-suggested rewrite

Summary table with total counts

Example

Input:

The report was written by the student.

Output:

The student wrote the report.

Folder Structure:

NLP_Text_Smell_Analysis/
├── req/          # Input files
├── outputs/      # Generated reports
├── nlp_analysis.py
└── README.md