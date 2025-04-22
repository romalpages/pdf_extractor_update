from pydoc import doc
import pdfplumber
import pandas as pd
import io
from flask import Flask, request, send_file, jsonify
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.platypus import Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import re

app = Flask(__name__)

def extract_headings_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        if not pdf.pages:
            return "", ""
        lines = pdf.pages[0].extract_text().split('\n')
        main_heading = lines[0].strip() if len(lines) > 0 else ""
        sub_heading = lines[1].strip() if len(lines) > 1 else ""
    return main_heading, sub_heading

def extract_footer_elements(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            page_match = re.search(r'Page\s+\d+\s+of\s+(\d+)', text)
            printed_on_match = re.search(r'Printed On:\s+(\d{1,2}/\d{1,2}/\d{4})', text)
            time_match = re.search(r'(\d{1,2}:\d{2}:\d{2}(AM|PM))', text)

            if page_match or printed_on_match or time_match:
                return {
                    "total_pages": page_match.group(1) if page_match else "1",
                    "printed_on": printed_on_match.group(1) if printed_on_match else "",
                    "printed_time": time_match.group(1) if time_match else ""
                }
    return {"total_pages": "1", "printed_on": "", "printed_time": ""}

def is_number(token):
    cleaned = token.replace(',', '')
    if cleaned.replace('.', '', 1).isdigit() and cleaned.count('.') <= 1:
        try:
            float(cleaned)
            return True
        except ValueError:
            return False
    return False

def extract_tables_from_ESIC(pdf_file, search_terms=None):
    all_tables = []
    search_terms_found = set()
    search_terms_not_found = set(search_terms) if search_terms else set()

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables({
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "explicit_vertical_lines": [],
                "explicit_horizontal_lines": [],
                "snap_tolerance": 3,
                "join_tolerance": 3,
            })

            if tables:
                for table in tables:
                    cleaned_table = []
                    headers = ["SNo", "Is Disable", "IP Number", "IP Name", "No of Days", 
                              "Total Wages", "IP Contribution", "Reason"]
                    cleaned_table.append(headers)

                    for row in table:
                        row = [cell.strip() if isinstance(cell, str) else "" for cell in row]
                        flat_row = " ".join(row)
                        if not flat_row.strip():
                            continue

                        words = flat_row.split()
                        ip_number = None
                        for word in words:
                            if len(word) == 10 and word.isdigit():
                                ip_number = word
                                break
                        if not ip_number:
                            continue

                        try:
                            ip_index = words.index(ip_number)
                        except ValueError:
                            continue
                        prefix = " ".join(words[:ip_index])
                        suffix_tokens = words[ip_index+1:]

                        prefix_parts = prefix.split()
                        sno = prefix_parts[0] if prefix_parts else ""
                        is_disabled = prefix_parts[1] if len(prefix_parts) > 1 else ""
                        ip_name = " ".join(prefix_parts[2:]) if len(prefix_parts) > 2 else ""

                        numbers = [token for token in suffix_tokens if is_number(token)]
                        no_of_days = numbers[-3] if len(numbers) >= 3 else ""
                        wages = numbers[-2] if len(numbers) >= 2 else ""
                        contribution = numbers[-1] if len(numbers) >= 1 else ""

                        reason_tokens = [t for t in suffix_tokens if t not in numbers]
                        reason = " ".join(reason_tokens)

                        reason_lower = reason.lower()
                        known_reason_starts = ["on leave", "absent", "joined", "resigned", "on duty", "on training", "on tour", "left", "-", ""]

                        split_index = -1
                        for keyword in known_reason_starts:
                            if keyword in reason_lower:
                                split_index = reason_lower.index(keyword)
                                break

                        if split_index > 0:
                            possible_name = reason[:split_index].strip()
                            reason = reason[split_index:].strip()
                            ip_name += " " + possible_name
                        elif reason.strip() == "-":
                            pass

                        row_fixed = [
                            sno, 
                            is_disabled, 
                            ip_number, 
                            ip_name.strip(), 
                            no_of_days, 
                            wages, 
                            contribution, 
                            reason
                        ]
                        cleaned_table.append(row_fixed)

                        if search_terms:
                            current_values = [str(v).lower() for v in row_fixed]
                            for term in search_terms:
                                if term.lower() in current_values:
                                    search_terms_found.add(term)
                                    search_terms_not_found.discard(term)

                    if len(cleaned_table) > 1:
                        df = pd.DataFrame(cleaned_table[1:], columns=cleaned_table[0])
                        all_tables.append(df)

    return all_tables, search_terms_found, search_terms_not_found

def search_and_extract_ip(all_tables, search_terms):
    search_terms = [s.strip().lower() for s in search_terms]
    combined_rows = []
    columns = []

    for df in all_tables:
        df = df.dropna(subset=["IP Number", "IP Name"])
        df["ip_number_lower"] = df["IP Number"].astype(str).str.lower()
        df["ip_name_lower"] = df["IP Name"].astype(str).str.lower()

        filtered_df = df[df["ip_number_lower"].isin(search_terms) | df["ip_name_lower"].isin(search_terms)]
        if not filtered_df.empty:
            combined_rows.append(filtered_df.drop(columns=["ip_number_lower", "ip_name_lower"]))
            columns = df.columns.drop(["ip_number_lower", "ip_name_lower"])

    return pd.concat(combined_rows, ignore_index=True) if combined_rows else pd.DataFrame(columns=columns)

def add_page_number(canvas, doc, footer_elements={}):
    page_num = canvas.getPageNumber()
    canvas.setFont("Helvetica", 8)
    canvas.drawString(30, 15, f"Page {page_num}")
    canvas.drawRightString(doc.pagesize[0] - 30, 15, f"Printed On: {footer_elements.get('printed_on', '')}")
    canvas.drawRightString(doc.pagesize[0] - 30,38, footer_elements.get('printed_time', ''))

def generate_single_table_pdf(df, main_heading="", sub_heading="", footer_elements={}):
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), leftMargin=2, rightMargin=2, topMargin=20, bottomMargin=20)

    styles = getSampleStyleSheet()
    centered_heading = ParagraphStyle(name="CenterHeading", parent=styles['Heading2'], alignment=1, fontSize=13)
    centered_subheading = ParagraphStyle(name="CenterSubheading", parent=styles['Normal'], alignment=1, fontSize=9)

    story = []

    try:
        logo_path = "logo.jpg"
        logo = Image(logo_path, width=80, height=70)
    except Exception:
        logo = None

    heading_elements = [logo if logo else Spacer(1, 50)]

    heading_text = []
    if main_heading:
        heading_text.append(Paragraph(f"<u>{main_heading}</u>", centered_heading))
    if sub_heading:
        heading_text.append(Paragraph(sub_heading, centered_subheading))

    heading_table = Table([[heading_elements[0], heading_text]], colWidths=[60, 440])
    heading_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("LEFTPADDING", (0, 0), (-2, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(heading_table)
    story.append(Spacer(1, 12))

    column_map = {
        'sno': 'SNo',
        'is_disable': 'Is Disable',
        'ip_number': 'IP Number',
        'ip_name': 'IP Name',
        'no_of_days': 'No. Of Days',
        'total_wages': 'Total Wages',
        'ip_contribution': 'IP Contribution',
        'reason': 'Reason'
    }

    df.columns = df.columns.str.strip().str.lower()
    data = [list(column_map.values())]

    for idx, row in df.iterrows():
        data_row = [
            str(idx + 1),
            row.get('is disable', ''),
            row.get('ip number', ''),
            row.get('ip name', ''),
            row.get('no of days', ''),
            row.get('total wages', ''),
            row.get('ip contribution', ''),
            row.get('reason', '')
        ]
        data.append(data_row)

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
        ('LINEABOVE', (0, 0), (-1, 0), 0.5, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.black),
        ('LINEBEFORE', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEAFTER', (0, 0), (-2, -1), 0.5, colors.black),
        ('LINEBELOW', (0, -1), (-1, -1), 0.5, colors.black),  
    ]))

    story.append(table)

    def footer_callback(canvas, doc):
        add_page_number(canvas, doc, footer_elements)

    doc.build(story, onFirstPage=footer_callback, onLaterPages=footer_callback)
    output.seek(0)
    return output

@app.route('/extract-ip', methods=['POST'])
def extract_ip_from_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "PDF file is required"}), 400
    file = request.files['file']
    search_input = request.form.get('ip_number', '').strip()

    if not search_input:
        return jsonify({"error": "At least one IP number or name is required"}), 400

    search_terms = [term.strip().lower() for term in search_input.split('|') if term.strip()]
    if not search_terms:
        return jsonify({"error": "No valid search terms provided"}), 400

    try:
        main_heading, sub_heading = extract_headings_from_pdf(file)
        footer_elements = extract_footer_elements(file)
        file.seek(0)
        tables, found, not_found = extract_tables_from_ESIC(file, search_terms)
        combined_df = search_and_extract_ip(tables, search_terms)

        if combined_df.empty:
            return jsonify({"error": "No matching records found"}), 404

        pdf_buffer = generate_single_table_pdf(combined_df, main_heading, sub_heading, footer_elements)

        return send_file(
            pdf_buffer,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="filtered_combined_table.pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__': 
    app.run(debug=True)
