import os
import fitz  # PyMuPDF for PDF reading
import spacy
import pandas as pd
from collections import Counter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from passive_to_active import passive_to_active

# ========== CONFIGURATION ==========
PDF_FOLDER = "req"
OUTPUT_FOLDER = "outputs"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Load English NLP model
print("Loading spaCy model...")
nlp = spacy.load("en_core_web_sm")


# ========== HELPER FUNCTIONS ==========
def is_passive(sent):
    """Check if a sentence is in passive voice using dependency parse."""
    for token in sent:
        if token.dep_ == "auxpass":
            return True
    return False


def has_conditional_modal(sent):
    """Check if a sentence contains conditional/modal verbs."""
    modal_words = ["should", "could", "would", "might", "must"]
    return any(token.text.lower() in modal_words for token in sent)


def analyze_pdf(file_path, file_name):
    """Extract text page-by-page, analyze sentences, return list of issues."""
    pdf_document = fitz.open(file_path)
    findings = []

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        doc = nlp(text)

        for sent in doc.sents:
            sent_text = sent.text.strip()
            if len(sent_text.split()) < 3:
                continue

            # Passive voice detection
            if is_passive(sent):
                findings.append(
                    {
                        "File": file_name,
                        "Page": page_num + 1,
                        "Sentence": sent_text,
                        "Issue": "Passive Voice",
                        "Suggestion": passive_to_active(sent.text),
                    }
                )

            # Conditional modal detection
            if has_conditional_modal(sent):
                findings.append(
                    {
                        "File": file_name,
                        "Page": page_num + 1,
                        "Sentence": sent_text,
                        "Issue": "Conditional Modal",
                    }
                )

    return findings


def save_pdf(data):
    """Generate a well-formatted PDF report."""
    pdf_path = os.path.join(OUTPUT_FOLDER, "NLP_Analysis_Report.pdf")
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(pdf_path)
    story = []

    # Title
    story.append(
        Paragraph(
            "<b>Passive voice and Conditional modal Detector</b>", styles["Title"]
        )
    )
    story.append(Spacer(1, 20))

    # Summary Table
    issue_counts = Counter([r["Issue"] for r in data])
    summary_data = [["Type of Bad Smell", "Frequency"]]
    for issue, freq in issue_counts.items():
        summary_data.append([issue, str(freq)])

    summary_table = Table(summary_data, colWidths=[250, 100])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ]
        )
    )

    story.append(Paragraph("<b>Summary of Detected Issues</b>", styles["Heading2"]))
    story.append(summary_table)
    story.append(Spacer(1, 20))

    # Detailed Results
    story.append(Paragraph("<b>Samples</b>", styles["Heading2"]))
    story.append(Spacer(1, 10))

    for r in data:
        story.append(
            Paragraph(
                f"<b>File:</b> {r['File']} | <b>Page:</b> {r['Page']}", styles["Normal"]
            )
        )
        story.append(Paragraph(f"<b>Sentence:</b> {r['Sentence']}", styles["Normal"]))
        story.append(Paragraph(f"<b>Issue:</b> {r['Issue']}", styles["Normal"]))
        if r["Issue"] == "Passive Voice":
            story.append(
                Paragraph(f"<b>Active Voice:</b> {r['Suggestion']}", styles["Normal"])
            )
        story.append(Spacer(1, 12))

    doc.build(story)
    print(f"PDF saved: {pdf_path}")
    return pdf_path


# ========== MAIN EXECUTION ==========
def main():
    print("Starting NLP Text Smell Analysis...\n")

    all_results = []
    for file_name in os.listdir(PDF_FOLDER):
        if file_name.lower().endswith(".pdf"):
            file_path = os.path.join(PDF_FOLDER, file_name)
            print(f"Analyzing: {file_name}")
            results = analyze_pdf(file_path, file_name)
            all_results.extend(results)

    if not all_results:
        print("No issues detected or no valid PDFs found.")
        return

    pdf_path = save_pdf(all_results)

    print("\n Analysis complete!")


# ========== ENTRY POINT ==========
if __name__ == "__main__":
    main()
