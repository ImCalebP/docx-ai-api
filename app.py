from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openai, os, tempfile, re

# ── OpenAI key (set on Render) ────────────────────────────────────────────────
openai.api_key = os.getenv("OPENAI_API_KEY")

app = Flask(__name__)

# ── GPT: return Markdown-style structure ─────────────────────────────────────
def gpt_markdown(raw_text: str) -> str:
    sys_msg = (
        "You return structured Markdown. Use:\n"
        "- # Title (once)\n"
        "- ## Section headings\n"
        "- Bullets starting with '- '\n"
        "- Short paragraphs\n"
        "Do NOT wrap in code-blocks."
    )
    usr_msg = f"Format the following:\n\n\"\"\"\n{raw_text}\n\"\"\""
    res = openai.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "system", "content": sys_msg},
                  {"role": "user",    "content": usr_msg}],
        temperature=0.5,
    )
    return res.choices[0].message.content.strip()

# ── Helpers ──────────────────────────────────────────────────────────────────
def _add_page_numbers(section):
    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run()
    fld_begin = OxmlElement('w:fldChar'); fld_begin.set(qn('w:fldCharType'), 'begin')
    instr     = OxmlElement('w:instrText'); instr.text = " PAGE "
    fld_end   = OxmlElement('w:fldChar');  fld_end.set(qn('w:fldCharType'), 'end')
    run._r.extend((fld_begin, instr, fld_end))

def markdown_to_docx(md: str, filename: str):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)
    _add_page_numbers(sec)

    lines = md.splitlines()
    if lines and lines[0].startswith("# "):
        title = doc.add_heading(lines.pop(0)[2:].strip(), level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.size = Pt(22)
        title.runs[0].font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
        doc.add_paragraph("")

    for ln in lines:
        ln = ln.strip()
        if ln.startswith("## "):
            h = doc.add_heading(ln[3:], level=1)
            h.runs[0].font.color.rgb = RGBColor(0x00, 0x57, 0xA6)
        elif ln.startswith("- "):
            doc.add_paragraph(ln[2:], style="List Bullet")
        elif ln == "":
            doc.add_paragraph("")
        else:
            p = doc.add_paragraph(ln)
            p.style.font.size = Pt(11)

    doc.save(filename)

def extract_title(md: str) -> str:
    """Return first '# Title' line or fallback"""
    for line in md.splitlines():
        if line.startswith("# "):
            return line[2:].strip()
    return "ai_document"

def safe_filename(title: str) -> str:
    """Sanitise title → safe file name"""
    cleaned = re.sub(r'[^A-Za-z0-9 _-]', '', title)      # drop weird chars
    cleaned = cleaned.strip().replace(" ", "_")           # spaces → _
    return (cleaned[:50] or "ai_document") + ".docx"      # max 50 chars

# ── Flask route ───────────────────────────────────────────────────────────────
@app.route("/generate-docx", methods=["POST"])
def generate_docx():
    data = request.get_json(silent=True)
    if not data or "text" not in data:
        return jsonify({"error": "JSON body must contain 'text'"}), 400
    try:
        md       = gpt_markdown(data["text"])
        title    = extract_title(md)
        filename = safe_filename(title)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            markdown_to_docx(md, tmp.name)
            return send_file(
                tmp.name,
                as_attachment=True,
                download_name=filename,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/")
def root():
    return "✅ DOCX API running"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
