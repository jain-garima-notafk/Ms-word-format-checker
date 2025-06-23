import streamlit as st
from docx import Document

def check_formatting(file):
    doc = Document(file)
    score = 0
    report = []

    section = doc.sections[0]

    # Q1(A): Page size A3
    width = round(section.page_width.inches, 1)
    height = round(section.page_height.inches, 1)
    if width == 11.7 and height == 16.5:
        score += 2
        report.append("✅ Q1(A): Page size is A3")
    else:
        report.append(f"❌ Q1(A): Page size is {width}x{height}, expected A3")

    # Q1(B): Paragraph border on paragraph 2 (simplified)
    if len(doc.paragraphs) >= 2 and 'w:top' in doc.paragraphs[1]._element.xml:
        score += 2
        report.append("✅ Q1(B): Border applied to paragraph 2")
    else:
        report.append("❌ Q1(B): No border found on paragraph 2")

    # Q1(C): Top margin 0.6
    top_margin = round(section.top_margin.inches, 1)
    if top_margin == 0.6:
        score += 2
        report.append("✅ Q1(C): Top margin is 0.6")
    else:
        report.append(f"❌ Q1(C): Top margin is {top_margin}, expected 0.6")

    # Q1(D): Bold + highlight 'Define' in 4th para
    found_define = False
    if len(doc.paragraphs) >= 4:
        for run in doc.paragraphs[3].runs:
            if 'Define' in run.text and run.bold and run.font.highlight_color:
                found_define = True
                break
    if found_define:
        score += 4
        report.append("✅ Q1(D): 'Define' is bold and highlighted")
    else:
        report.append("❌ Q1(D): 'Define' not found bold and highlighted")

    # Q2(C): Line spacing 1.5 in paragraph 2
    if len(doc.paragraphs) >= 2:
        spacing = doc.paragraphs[1].paragraph_format.line_spacing
        if spacing and round(spacing, 2) == 1.5:
            score += 2
            report.append("✅ Q2(C): Line spacing 1.5 in paragraph 2")
        else:
            report.append("❌ Q2(C): Line spacing not 1.5 in paragraph 2")

    # Q3(B): Paragraph 1 line spacing 1.15
    if len(doc.paragraphs) >= 1:
        spacing = doc.paragraphs[0].paragraph_format.line_spacing
        if spacing and round(spacing, 2) == 1.15:
            score += 2
            report.append("✅ Q3(B): Line spacing 1.15 in paragraph 1")
        else:
            report.append("❌ Q3(B): Line spacing not 1.15 in paragraph 1")

    # Q4(A): Font color orange (simplified check)
    found_orange = False
    for p in doc.paragraphs:
        for run in p.runs:
            if run.font.color and run.font.color.rgb and str(run.font.color.rgb) == 'FFA500':
                found_orange = True
                break
        if found_orange:
            break
    if found_orange:
        score += 2
        report.append("✅ Q4(A): Found orange font")
    else:
        report.append("❌ Q4(A): No orange font found")

    return score, report

# Streamlit UI
st.title("🧾 MS Word Formatting Auto Checker – All Questions")
uploaded_file = st.file_uploader("Upload your .docx file", type=["docx"])

if uploaded_file:
    score, results = check_formatting(uploaded_file)
    st.success(f"✅ Total Score: {score}/30 (Q1-Q4 implemented)")
    st.write("### Details:")
    for r in results:
        st.write(r)