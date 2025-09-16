# app.py
import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import tempfile, os, zipfile, re
from xml.sax.saxutils import escape

st.set_page_config(layout="wide")
st.title("MCQ Question Paper Generator")

# ---------- FONT HANDLING ----------
# Point to your DejaVu extracted folder
FONTS_DIR = r"C:\Users\10568\OneDrive - PHN TECHNOLOGY PRIVATE LIMITED\Desktop\webdevpro\pdfConvertor\dejavu-fonts-ttf-2.37\ttf"

registered_font = None
if os.path.isdir(FONTS_DIR):
    try:
        # Register all .ttf files in the folder
        for fname in os.listdir(FONTS_DIR):
            if fname.lower().endswith(".ttf"):
                font_path = os.path.join(FONTS_DIR, fname)
                font_name = os.path.splitext(fname)[0]
                pdfmetrics.registerFont(TTFont(font_name, font_path))
        # Pick base font
        if "DejaVuSans" in pdfmetrics.getRegisteredFontNames():
            registered_font = "DejaVuSans"
        elif "DejaVuSans.ttf" in pdfmetrics.getRegisteredFontNames():
            registered_font = "DejaVuSans.ttf"
    except Exception as e:
        st.error(f"Font registration failed: {e}")

if not registered_font:
    st.error("⚠️ DejaVu fonts not found. Superscripts/subscripts may still fail. "
             "Please check your `ttf` folder path.")

# ---------- Text cleaning / conversion helpers ----------
SUPER_MAP = {"²":"2","³":"3","¹":"1","⁴":"4","⁵":"5","⁶":"6","⁷":"7","⁸":"8","⁹":"9","⁰":"0","⁻":"-"}
SUB_MAP   = {"₂":"2","₃":"3","₁":"1","₄":"4","₅":"5","₆":"6","₇":"7","₈":"8","₉":"9","₀":"0"}

def replace_unicode_super_sub(s: str) -> str:
    for k,v in SUPER_MAP.items():
        s = s.replace(k, f"<super>{v}</super>")
    for k,v in SUB_MAP.items():
        s = s.replace(k, f"<sub>{v}</sub>")
    return s

def preserve_tags_escape(s: str) -> str:
    placeholders = {
        "<super>":"[[SUPER_OPEN]]", "</super>":"[[SUPER_CLOSE]]",
        "<sub>":"[[SUB_OPEN]]", "</sub>":"[[SUB_CLOSE]]",
        "<b>":"[[B_OPEN]]","</b>":"[[B_CLOSE]]",
        "<i>":"[[I_OPEN]]","</i>":"[[I_CLOSE]]"
    }
    for k,v in placeholders.items():
        s = s.replace(k, v)
    s = escape(s)
    for k,v in placeholders.items():
        s = s.replace(v, k)
    return s

def clean_text(raw_text: str) -> str:
    if raw_text is None:
        return ""
    s = str(raw_text)
    s = s.replace("■", "")
    s = replace_unicode_super_sub(s)
    s = re.sub(r'([0-9\.\-]+)\^(-?[0-9]+)', r'\1<super>\2</super>', s)
    s = re.sub(r'm/s\^?([0-9]+)', r'm/s<super>\1</super>', s, flags=re.IGNORECASE)
    s = re.sub(r'([A-Za-z]+)/s\^?([0-9]+)', r'\1/s<super>\2</super>', s, flags=re.IGNORECASE)
    s = re.sub(r'([A-Za-z\)\]])_([0-9]+)', r'\1<sub>\2</sub>', s)
    s = re.sub(r'(?<![A-Za-z0-9])([A-Z][a-z]?)(\d+)', r'\1<sub>\2</sub>', s)
    s = re.sub(r'(10)\s*[×xX]\s*10\^?(-?[0-9]+)', r'\1×10<super>\2</super>', s)
    s = preserve_tags_escape(s)
    return s

# ---------- Draw header/footer ----------
def draw_header_footer(canvas, doc, set_name, is_first_page=False):
    page_width, page_height = A4
    margin = 1.5 * cm
    fontnames = pdfmetrics.getRegisteredFontNames()
    heading_font = "DejaVuSans-Bold" if "DejaVuSans-Bold" in fontnames else ("DejaVuSans" if "DejaVuSans" in fontnames else "Helvetica-Bold")
    normal_font = "DejaVuSans" if "DejaVuSans" in fontnames else "Helvetica"
    dark_blue = colors.HexColor("#001F4D")

    canvas.setFont(heading_font, 14)
    canvas.setFillColor(dark_blue)
    line1_y = page_height - 1.5 * cm
    line2_y = line1_y - 0.6 * cm
    line3_y = line2_y - 0.6 * cm
    line4_y = line3_y - 0.5 * cm

    canvas.drawCentredString(page_width / 2, line1_y, "PHN Scholar Exam 2025-26")
    canvas.setFont(heading_font, 12)
    canvas.drawCentredString(page_width / 2, line2_y, "Question Paper - Class 9")
    canvas.setFont(normal_font, 11)
    canvas.drawCentredString(page_width / 2, line3_y, "Maharashtra State Board")
    canvas.drawCentredString(page_width / 2, line4_y, f"SET: {set_name}")

    try:
        logo_left = ImageReader("scholar.png")
        logo_right = ImageReader("logo.png")
        logo_size = 2.5 * cm
        header_center_y = (line1_y + line4_y) / 2
        y_position = header_center_y - (logo_size / 2)
        canvas.drawImage(logo_left, margin, y_position, width=logo_size, height=logo_size, preserveAspectRatio=True, mask="auto")
        canvas.drawImage(logo_right, page_width - margin - logo_size, y_position, width=logo_size, height=logo_size, preserveAspectRatio=True, mask="auto")
    except Exception:
        pass

    line_y = line4_y - 0.8 * cm
    canvas.setStrokeColor(colors.black)
    canvas.line(margin, line_y, page_width - margin, line_y)

    canvas.setFont(normal_font, 9)
    canvas.setFillColor(colors.black)
    canvas.drawRightString(page_width - margin, 0.7 * cm, f"Page {doc.page}")

# ---------- Header Section ----------
def create_header_section(total_marks, time_duration, sections, no_of_questions, marks_per_question, instructions):
    story = []
    styles = getSampleStyleSheet()
    base_font = "DejaVuSans" if "DejaVuSans" in pdfmetrics.getRegisteredFontNames() else styles['Normal'].fontName

    marks_data = [['Total Marks:', total_marks, '', 'Time Duration:', time_duration]]
    marks_table = Table(marks_data, colWidths=[2.5*cm, 2*cm, 8*cm, 3*cm, 2*cm])
    marks_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), base_font),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(marks_table)
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("<b><font size=11>PATTERN & MARKING SCHEME</font></b>", styles['Normal']))
    story.append(Spacer(1, 0.3*cm))

    pattern_data = [
        ['Section', *[f"({i+1}) {s.strip()}" for i, s in enumerate(sections)]],
        ['No. Of Questions', *no_of_questions],
        ['Marks Per Ques.', *marks_per_question]
    ]
    col_widths = [4.5*cm] + [5*cm] * len(sections)
    pattern_table = Table(pattern_data, colWidths=col_widths)
    pattern_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,0), base_font),
        ('FONTNAME', (0,1), (0,-1), base_font),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(pattern_table)
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("<b><font size=11>INSTRUCTIONS</font></b>", styles['Normal']))
    story.append(Spacer(1, 0.3*cm))
    instruction_style = ParagraphStyle(name='InstructionStyle', alignment=TA_LEFT, fontSize=10, fontName=base_font, leftIndent=10)
    for instruction in instructions:
        if instruction.strip():
            story.append(Paragraph(f"• {escape(instruction.strip())}", instruction_style))
            story.append(Spacer(1, 0.1*cm))

    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("____________________________________________________________________________________", styles['Normal']))
    story.append(Spacer(1, 0.5*cm))
    return story

# ---------- PDF Generator ----------
def generate_pdf_for_set(df_set, set_name, sections, total_marks, time_duration, no_of_questions, marks_per_question, instructions):
    story = []
    styles = getSampleStyleSheet()
    base_font = "DejaVuSans" if "DejaVuSans" in pdfmetrics.getRegisteredFontNames() else styles['Normal'].fontName

    styles.add(ParagraphStyle(name='QuestionText', alignment=TA_LEFT, fontSize=11, fontName=base_font))
    styles.add(ParagraphStyle(name='OptionText', alignment=TA_LEFT, fontSize=10, fontName=base_font))
    styles.add(ParagraphStyle(name='SectionHeading', alignment=TA_CENTER, fontSize=12, fontName=base_font, textColor=colors.white, backColor=colors.HexColor("#001F4D"), borderPadding=5))

    first_page_story = create_header_section(total_marks, time_duration, sections, no_of_questions, marks_per_question, instructions)
    story.extend(first_page_story)

    question_number = 1
    col_idx = 1
    current_section_idx = 0
    questions_in_section = 0
    total_questions = 50

    if sections:
        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph(f"Section 1 - {sections[0].strip().upper()}", styles['SectionHeading']))
        story.append(Spacer(1, 0.5*cm))

    while col_idx < len(df_set.columns) and question_number <= total_questions:
        if current_section_idx < len(sections) and questions_in_section >= no_of_questions[current_section_idx]:
            questions_in_section = 0
            current_section_idx += 1
            if current_section_idx < len(sections):
                story.append(Spacer(1, 0.5*cm))
                story.append(Paragraph(f"Section {current_section_idx+1} - {sections[current_section_idx].strip().upper()}", styles['SectionHeading']))
                story.append(Spacer(1, 0.5*cm))

        if col_idx + 4 >= len(df_set.columns):
            break

        q_col = df_set.columns[col_idx]
        opts_cols = df_set.columns[col_idx + 1: col_idx + 5]
        question_text = clean_text(df_set.iloc[0][q_col])
        options = [clean_text(df_set.iloc[0][oc]) for oc in opts_cols]

        question_block = []
        question_block.append(Paragraph(f"<b>{question_number}) {question_text}</b>", styles['QuestionText']))

        if max(len(re.sub(r'<.*?>','',opt)) for opt in options) < 25:
            row = [[Paragraph(f"A) {options[0]}", styles['OptionText']),
                    Paragraph(f"B) {options[1]}", styles['OptionText']),
                    Paragraph(f"C) {options[2]}", styles['OptionText']),
                    Paragraph(f"D) {options[3]}", styles['OptionText'])]]
            options_table = Table(row, colWidths=[4*cm, 4*cm, 4*cm, 4*cm])
        else:
            row = [
                [Paragraph(f"A) {options[0]}", styles['OptionText']),
                 Paragraph(f"B) {options[1]}", styles['OptionText'])],
                [Paragraph(f"C) {options[2]}", styles['OptionText']),
                 Paragraph(f"D) {options[3]}", styles['OptionText'])]
            ]
            options_table = Table(row, colWidths=[8*cm, 8*cm])

        options_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))

        question_block.append(options_table)
        question_block.append(Spacer(1, 0.3*cm))
        story.append(KeepTogether(question_block))

        col_idx += 5
        question_number += 1
        questions_in_section += 1

    pdf_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
    doc = SimpleDocTemplate(pdf_file, pagesize=A4,
                            rightMargin=cm, leftMargin=cm, topMargin=4.5*cm, bottomMargin=cm)

    def on_page(canvas, doc):
        draw_header_footer(canvas, doc, set_name)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return pdf_file

# ---------- Helper ----------
def get_unique_sets(df):
    if "Set" in df.columns:
        return sorted(df["Set"].dropna().unique().tolist())
    else:
        return []

# ---------- Streamlit UI ----------
excel_file = st.file_uploader("Upload MCQ Excel File", type=["xlsx"])
st.markdown("**Tip:** Use `<super>` and `<sub>` tags in Excel cells for perfect control, e.g. `10<super>-11</super>` or `H<sub>2</sub>O`. The app will also auto-convert common notations.")

if excel_file:
    with st.expander("Exam Details"):
        school_name = st.text_input("School Name", "PHN Scholar Exam 2025-26")
        board_name = st.text_input("Board", "Maharashtra State Board")
        exam_name = st.text_input("Exam Name", "Question Paper")
        class_name = st.text_input("Class", "Class 9")
        total_marks = st.number_input("Total Marks", value=50)
        time_duration = st.text_input("Time Duration", "60 Mins")

    with st.expander("Section & Marking Schema"):
        sections = st.text_area("Sections (comma separated)", "Maths,Science,English").split(",")
        no_of_questions = st.text_area("No. of Questions (comma separated)", "20,20,10").split(",")
        marks_per_question = st.text_area("Marks Per Question (comma separated)", "1,1,1").split(",")
        try:
            no_of_questions = list(map(int, no_of_questions))
            marks_per_question = list(map(int, marks_per_question))
        except ValueError:
            st.error("Please enter numbers for 'No. of Questions' and 'Marks Per Question'.")

    with st.expander("Instructions"):
        instructions = st.text_area(
            "Exam Instructions (each line will be separate instruction)",
            "Check your Roll Number, Class, and other details carefully before starting.\n"
            "Find the correct answer and darken the circle completely in the OMR sheet with a black/blue ball pen.\n"
            "Do not use a pencil, gel pen, or sketch pen.\n"
            "Darken only one circle for each question. If more than one circle is darkened, it will not be counted.\n"
            "Do not make any extra marks, ticks, or scratches on the OMR sheet.\n"
            "Answers marked on the question paper will not be checked. Only OMR sheet answers will be considered.\n"
            "Follow the instructions given by the invigilator in the exam hall.\n"
            "Kindly do not write on the question paper. Return it to the examiner after the exam.\n",
        ).split("\n")

    temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_excel.write(excel_file.read())
    temp_excel.close()
    df_raw = pd.read_excel(temp_excel.name)
    df_raw.columns = df_raw.columns.str.strip()
    unique_sets = get_unique_sets(df_raw)

    if not unique_sets:
        st.error("No 'Set' column found in your Excel file.")
    else:
        selected_set = st.selectbox("Select a Set to download", unique_sets)
        col1, col2 = st.columns(2)

        with col1:
            if st.button("Download Selected Set"):
                df_set = df_raw[df_raw["Set"] == selected_set].reset_index(drop=True)
                pdf_path = generate_pdf_for_set(
                    df_set, selected_set, sections, total_marks, time_duration,
                    no_of_questions, marks_per_question, instructions
                )
                with open(pdf_path, "rb") as f:
                    st.download_button("Download PDF", f, file_name=f"MCQ_Set_{selected_set}.pdf")

        with col2:
            if st.button("Download All Sets as ZIP"):
                with tempfile.TemporaryDirectory() as temp_dir:
                    zip_file_path = os.path.join(temp_dir, "All_MCQ_Sets.zip")
                    with zipfile.ZipFile(zip_file_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                        for current_set in unique_sets:
                            df_set = df_raw[df_raw["Set"] == current_set].reset_index(drop=True)
                            pdf_path = generate_pdf_for_set(
                                df_set, current_set, sections, total_marks, time_duration,
                                no_of_questions, marks_per_question, instructions
                            )
                            zipf.write(pdf_path, f"MCQ_Set_{current_set}.pdf")
                    with open(zip_file_path, "rb") as f:
                        st.download_button("Download ZIP", f, file_name="All_MCQ_Sets.zip")
