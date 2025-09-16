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
import tempfile, os, zipfile

# --- Streamlit UI ---
st.set_page_config(layout="wide")
st.title("MCQ Question Paper Generator")

# --- Draw Header & Footer ---
def draw_header_footer(canvas, doc, set_name, is_first_page=False):
    page_width, page_height = A4
    margin = 1.5 * cm

    dark_blue = colors.HexColor("#001F4D")
    canvas.setFont("Helvetica-Bold", 14)
    canvas.setFillColor(dark_blue)

    # Line positions
    line1_y = page_height - 1.5 * cm
    line2_y = line1_y - 0.6 * cm
    line3_y = line2_y - 0.6 * cm
    line4_y = line3_y - 0.5 * cm

    # Draw header text
    canvas.drawCentredString(page_width / 2, line1_y, "PHN Scholar Exam 2025-26")
    canvas.setFont("Helvetica-Bold", 12)
    canvas.drawCentredString(page_width / 2, line2_y, "Question Paper - Class 9")
    canvas.setFont("Helvetica-Bold", 11)
    canvas.drawCentredString(page_width / 2, line3_y, "Maharashtra State Board")
    canvas.drawCentredString(page_width / 2, line4_y, f"SET: {set_name}")

    # --- Compute header block center ---
    header_center_y = (line1_y + line4_y) / 2

    # --- Logos aligned to center with spacing ---
    try:
        logo_left = ImageReader("scholar.png")
        logo_right = ImageReader("logo.png")
        logo_size = 2.5 * cm

        y_position = header_center_y - (logo_size / 2)

        canvas.drawImage(
            logo_left,
            margin,
            y_position,
            width=logo_size,
            height=logo_size,
            preserveAspectRatio=True,
            mask="auto",
        )
        canvas.drawImage(
            logo_right,
            page_width - margin - logo_size,
            y_position,
            width=logo_size,
            height=logo_size,
            preserveAspectRatio=True,
            mask="auto",
        )
    except:
        pass

    # --- Horizontal line after header ---
    line_y = line4_y - 0.8 * cm
    canvas.setStrokeColor(colors.black)
    canvas.line(margin, line_y, page_width - margin, line_y)

    # --- Footer ---
    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(colors.black)
    canvas.drawRightString(page_width - margin, 0.7 * cm, f"Page {doc.page}")

# --- Create Header Section ---
def create_header_section(total_marks, time_duration, sections, no_of_questions, marks_per_question, instructions):
    story = []
    styles = getSampleStyleSheet()

    # --- Total Marks & Duration Table ---
    marks_data = [['Total Marks:', total_marks, '', 'Time Duration:', time_duration]]
    marks_table = Table(marks_data, colWidths=[2.5*cm, 2*cm, 8*cm, 3*cm, 2*cm])
    marks_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(marks_table)
    story.append(Spacer(1, 0.5*cm))

    # --- Pattern & Marking Scheme Title ---
    story.append(Paragraph("<b><font size=11>PATTERN & MARKING SCHEME</font></b>", styles['Normal']))
    story.append(Spacer(1, 0.3*cm))

    # --- Pattern & Marking Scheme Table ---
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
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,1), (0,-1), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(pattern_table)
    story.append(Spacer(1, 0.5*cm))

    # --- Instructions Section ---
    story.append(Paragraph("<b><font size=11>INSTRUCTIONS</font></b>", styles['Normal']))
    story.append(Spacer(1, 0.3*cm))
    
    instruction_style = ParagraphStyle(name='InstructionStyle', alignment=TA_LEFT, fontSize=10, fontName='Helvetica', leftIndent=10)
    for instruction in instructions:
        if instruction.strip():
            story.append(Paragraph(f"• {instruction.strip()}", instruction_style))
            story.append(Spacer(1, 0.1*cm))
    
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("____________________________________________________________________________________", styles['Normal']))
    story.append(Spacer(1, 0.5*cm))
    
    return story

# --- PDF Generator ---
def generate_pdf_for_set(df_set, set_name, sections, total_marks, time_duration, no_of_questions, marks_per_question, instructions):
    story = []
    styles = getSampleStyleSheet()

    # Custom styles
    styles.add(ParagraphStyle(name='QuestionText', alignment=TA_LEFT, fontSize=11, fontName='Helvetica'))
    styles.add(ParagraphStyle(name='OptionText', alignment=TA_LEFT, fontSize=10, fontName='Helvetica'))
    styles.add(ParagraphStyle(
        name='SectionHeading',
        alignment=TA_CENTER,
        fontSize=12,
        fontName='Helvetica-Bold',
        textColor=colors.white,
        backColor=colors.HexColor("#001F4D"),
        borderPadding=5
    ))

    # --- Add Header Section ---
    first_page_story = create_header_section(total_marks, time_duration, sections, no_of_questions, marks_per_question, instructions)
    story.extend(first_page_story)

    question_number = 1
    col_idx = 1
    current_section_idx = 0
    questions_in_section = 0
    total_questions = 50

    # Section 1 Heading
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

        question_text = str(df_set.iloc[0][q_col]).strip()
        options = [str(df_set.iloc[0][oc]).strip() for oc in opts_cols]

        # --- KeepTogether block for question + options ---
        question_block = []
        question_block.append(Paragraph(f"<b>{question_number}) {question_text}</b>", styles['QuestionText']))

        if max(len(opt) for opt in options) < 25:
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

# --- FIXED Function ---
def get_unique_sets(df):
    if "Set" in df.columns:
        return sorted(df["Set"].dropna().unique().tolist())
    else:
        return []

# --- Upload Excel ---
excel_file = st.file_uploader("Upload MCQ Excel File", type=["xlsx"])

# --- Main Logic ---
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
