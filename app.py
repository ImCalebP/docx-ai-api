from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import tempfile
import os

# Use API key from environment variable
openai.api_key = os.environ.get("OPENAI_API_KEY")

app = Flask(__name__)

def generate_docx_from_text(text, filename):
    doc = Document()
    doc.add_heading("Generated Document", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Styled and structured with AI.\n")

    for paragraph in text.split("\n"):
        if paragraph.strip():
            p = doc.add_paragraph(paragraph.strip())
            p.style.font.size = Pt(11)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.save(filename)
    return filename

def gpt_format_text(user_text):
    prompt = f"""
You are a professional writing assistant. Take the following raw text and transform it into a polished, structured document with headings, bullet points (if needed), and clean formatting.

Text:
\"\"\"{user_text}\"\"\"
"""

    client = openai.OpenAI(api_key=os.environ["OPENAI_API_KEY"])
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[
            { "role": "system", "content": "You generate polished, structured business documents." },
            { "role": "user", "content": prompt }
        ],
        temperature=0.7
    )

    return response.choices[0].message.content.strip()

@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    data = request.get_json()
    if not data or 'text' not in data:
        return jsonify({"error": "Missing 'text' in request body"}), 400

    try:
        ai_text = gpt_format_text(data['text'])

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = generate_docx_from_text(ai_text, tmp.name)
            return send_file(output_path, as_attachment=True, download_name="document.docx")

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/')
def home():
    return '👋 Welcome to the AI DOCX Generator API!'

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
