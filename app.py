from flask import Flask, request, send_file, jsonify
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import tempfile

app = Flask(__name__)
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# === Format with OpenAI (structure only, visual will be handled below) ===
def gpt_format_text(user_text):
    messages = [
        {"role": "system", "content": "You are a professional document formatter who rewrites content using clear structure with headings, bullets, and sections."},
        {"role": "user", "content": f"Format this into a business document:\n\"\"\"\n{user_text}\n\"\"\""}
    ]
    response = client.chat.completions.create(
        model="gpt-4",
        messages=messages,
        temperature=0.6
    )
    return response.choices[0].message.content.strip()

# === Add page numbers ===
def add_footer_with_page_number(section):
    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

# === Convert Text to Beautiful DOCX ===
def generate_docx_from_text(text, filename):
    doc = Document()

    # Set margins
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    # Add Title
    title = doc.add_heading("AI Generated Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)  # Blue
    run.font.size = Pt(20)

    # Subtitle
    subtitle = doc.add_paragraph("Document generated and formatted using OpenAI & Python")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle.runs[0]
    subtitle_run.font.size = Pt(12)
    subtitle_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)  # Gray

    doc.add_paragraph("")  # Spacer

    # Main Content Formatting
    for line in text.split("\n"):
        stripped = line.strip()
        if stripped.startswith("# "):
            p = doc.add_heading(stripped[2:], level=1)
            p.runs[0].font.color.rgb = RGBColor(0x00, 0x00, 0x80)
        elif stripped.startswith("## "):
            p = doc.add_heading(stripped[3:], level=2)
            p.runs[0].font.color.rgb = RGBColor(0x33, 0x66, 0x99)
        elif stripped.startswith("- "):
            p = doc.add_paragraph(stripped[2:], style='List Bullet')
        elif stripped:
            para = doc.add_paragraph(stripped)
            para.style.font.size = Pt(11)

    # Add page numbers
    add_footer_with_page_number(section)

    doc.save(filename)
    return filename

@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    data = request.get_json()
    if not data or 'text' not in data:
        return jsonify({"error": "Missing 'text' in request body"}), 400

    try:
        ai_text = gpt_format_text(data['text'])

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = generate_docx_from_text(ai_text, tmp.name)
            return send_file(output_path, as_attachment=True, download_name="beautiful_doc.docx",
                             mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/')
def home():
    return '👋 Welcome to the Beautiful DOCX Generator API!'

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
