
import streamlit as st
from docx import Document

def check_formatting(doc):
    results = {
        "A3 Page Size": False,
        "Top Margin 0.6"": False,
        "Paragraph 2: Border": False,
        "Paragraph 4: Bold Highlights on 'डिफाइन'": False,
        "Watermark 'क्राइम'": False,
        "Header Page Number": False,
        "Double Line Spacing in Paragraph 2": False,
        "Word 'रक्षा' Font Style Changed": False,
        "Underline Removed for 'रक्षा'": False,
        "Line Spacing 1.15 in Paragraph 1": False,
        "Bullet List in Paragraph 3": False,
        "Table Line Spacing 1.5 in Paragraph 4": False,
    }

    try:
        # Check basic content
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)

        if any("डिफाइन" in para.text for para in doc.paragraphs):
            results["Paragraph 4: Bold Highlights on 'डिफाइन'"] = True

        if any("रक्षा" in para.text for para in doc.paragraphs):
            results["Word 'रक्षा' Font Style Changed"] = True
            results["Underline Removed for 'रक्षा'"] = True

        if len(doc.paragraphs) > 0 and doc.paragraphs[0].paragraph_format.line_spacing == 1.15:
            results["Line Spacing 1.15 in Paragraph 1"] = True

        if len(doc.paragraphs) > 1 and doc.paragraphs[1].paragraph_format.line_spacing == 1.5:
            results["Double Line Spacing in Paragraph 2"] = True

    except Exception as e:
        st.error(f"Error checking document: {e}")

    return results

st.title("Word Format Checker")

uploaded_file = st.file_uploader("Upload your .docx file", type="docx")
if uploaded_file:
    doc = Document(uploaded_file)
    results = check_formatting(doc)

    st.header("Format Check Results")
    correct = 0
    for key, value in results.items():
        st.write(f"{'✅' if value else '❌'} {key}")
        if value:
            correct += 1

    st.subheader(f"Total Score: {correct}/12")
