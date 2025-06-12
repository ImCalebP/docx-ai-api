from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openai
import tempfile
import os

# Use OpenAI key from environment variable
openai.api_key = os.environ.get("OPENAI_API_KEY")

app = Flask(__name__)

def gpt_generate_structure(user_text):
    system_message = {
        "role": "system",
        "content": "You return structured JSON with a document title, and a list of sections. Each section must have a heading and either content or bullet points."
    }

    user_message = {
        "role": "user",
        "content": f"""
Format this raw text into a structured JSON object like:
{{
  "title": "Document Title",
  "sections": [
    {{
      "heading": "Introduction",
      "content": "Some paragraph."
    }},
    {{
      "heading": "Features",
      "bullets": ["Point 1", "Point 2"]
    }}
  ]
}}

Text:
\"\"\"{user_text}\"\"\"
"""
    }

    client = openai.OpenAI(api_key=openai.api_key)
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[system_message, user_message],
        response_format="json"
    )
    return response.choices[0].message.content.strip()

def generate_styled_docx(structured, filename):
    import json
    data = json.loads(structured)
    doc = Document()

    # Set title style
    title = doc.add_heading(data.get("title", "Untitled Document"), level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.size = Pt(26)
    run.font.color.rgb = RGBColor(0, 102, 204)

    doc.add_paragraph("")  # Spacer

    for section in data.get("sections", []):
        heading = doc.add_heading(section["heading"], level=1)
        heading.runs[0].font.color.rgb = RGBColor(255, 140, 0)  # Orange

        if "content" in section:
            para = doc.add_paragraph(section["content"])
            para.runs[0].font.size = Pt(11)
        elif "bullets" in section:
            for item in section["bullets"]:
                doc.add_paragraph(item, style='List Bullet')

        doc.add_paragraph("")  # Spacer between sections

    # Add footer/pagination
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "Page "
    footer_para.add_run().add_field("PAGE")
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.save(filename)
    return filename

@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    data = request.get_json()
    if not data or 'text' not in data:
        return jsonify({"error": "Missing 'text' in request body"}), 400

    try:
        structured_json = gpt_generate_structure(data['text'])

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = generate_styled_docx(structured_json, tmp.name)
            return send_file(output_path, as_attachment=True, download_name="ai-styled.docx")

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/')
def home():
    return '📄 AI-Styled DOCX Generator API is running!'

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
