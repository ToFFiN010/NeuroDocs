from flask import Flask, request, jsonify, render_template_string, send_file
import os
import io
import zipfile
import re
import uuid
from html import escape
from dotenv import load_dotenv
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Preformatted
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

load_dotenv()

app = Flask(__name__)

# Temporary in-memory cache for generated documentation downloads
DOC_CACHE = {}

def get_active_api_key(custom_key=None):
    if custom_key and custom_key.strip():
        return custom_key.strip()
    return os.getenv("GEMINI_API_KEY", "").strip()

def fetch_gemini_models(api_key):
    fallback_models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-2.5-pro"]
    if not api_key:
        return fallback_models
    try:
        genai.configure(api_key=api_key)
        models = []
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                name = m.name.replace("models/", "")
                models.append(name)
        if models:
            return models
    except Exception:
        pass
    return fallback_models

def add_docx_markdown_run(paragraph, text):
    pattern = re.compile(r'(\*\*.*?\*\*|\*.*?\*|`.*?`)')
    tokens = pattern.split(text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) > 4:
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("*") and token.endswith("*") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        elif token.startswith("`") and token.endswith("`") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run.font.name = "Consolas"
        else:
            paragraph.add_run(token)

def generate_docx_bytes(text):
    document = Document()
    document.add_heading("AI Generated Documentation", level=1)
    lines = text.split("\n")
    in_code_block = False
    code_lines = []

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```"):
            if in_code_block:
                p = document.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                run = p.add_run("\n".join(code_lines))
                run.font.name = "Consolas"
                run.font.size = Pt(9.5)
                code_lines = []
                in_code_block = False
            else:
                in_code_block = True
                code_lines = []
            continue

        if in_code_block:
            code_lines.append(line)
            continue

        if not stripped:
            continue

        if stripped.startswith("# "):
            document.add_heading(stripped[2:], level=1)
        elif stripped.startswith("## "):
            document.add_heading(stripped[3:], level=2)
        elif stripped.startswith("### "):
            document.add_heading(stripped[4:], level=3)
        elif re.match(r'^[\*\-\+]\s+', stripped):
            content = re.sub(r'^[\*\-\+]\s+', '', stripped)
            p = document.add_paragraph(style='List Bullet')
            add_docx_markdown_run(p, content)
        elif re.match(r'^\d+\.\s+', stripped):
            content = re.sub(r'^\d+\.\s+', '', stripped)
            p = document.add_paragraph(style='List Number')
            add_docx_markdown_run(p, content)
        else:
            p = document.add_paragraph()
            add_docx_markdown_run(p, stripped)

    bio = io.BytesIO()
    document.save(bio)
    bio.seek(0)
    return bio

def md_to_pdf_html(line):
    line = line.replace('’', "'").replace('“', '"').replace('”', '"').replace('–', '-').replace('—', '-')
    line = line.encode('latin-1', 'ignore').decode('latin-1')
    line = escape(line)
    line = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', line)
    line = re.sub(r'\*(.*?)\*', r'<i>\1</i>', line)
    line = re.sub(r'`(.*?)`', r'<font face="Courier" color="#c7254e">\1</font>', line)
    return line

def generate_pdf_bytes(text):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(
        bio,
        pagesize=letter,
        rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Title'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=colors.HexColor('#1E293B'),
        spaceAfter=15
    )
    h1_style = ParagraphStyle(
        'H1',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=19,
        textColor=colors.HexColor('#0F172A'),
        spaceBefore=12,
        spaceAfter=6,
        keepWithNext=True
    )
    h2_style = ParagraphStyle(
        'H2',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=17,
        textColor=colors.HexColor('#2563EB'),
        spaceBefore=10,
        spaceAfter=4,
        keepWithNext=True
    )
    h3_style = ParagraphStyle(
        'H3',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor('#334155'),
        spaceBefore=8,
        spaceAfter=4,
        keepWithNext=True
    )
    body_style = ParagraphStyle(
        'Body',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#334155'),
        spaceAfter=6
    )
    bullet_style = ParagraphStyle(
        'Bullet',
        parent=body_style,
        leftIndent=15,
        spaceAfter=3
    )
    code_style = ParagraphStyle(
        'Code',
        fontName='Courier',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#0F172A'),
        backColor=colors.HexColor('#F1F5F9'),
        borderColor=colors.HexColor('#E2E8F0'),
        borderWidth=1,
        borderPadding=6,
        spaceBefore=6,
        spaceAfter=8
    )

    story = [Paragraph("AI Generated Documentation", title_style), Spacer(1, 10)]
    lines = text.split("\n")
    in_code_block = False
    code_lines = []

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```"):
            if in_code_block:
                code_text = escape("\n".join(code_lines))
                story.append(Preformatted(code_text, code_style))
                code_lines = []
                in_code_block = False
            else:
                in_code_block = True
                code_lines = []
            continue

        if in_code_block:
            code_lines.append(line)
            continue

        if not stripped:
            continue

        if stripped.startswith("# "):
            story.append(Paragraph(md_to_pdf_html(stripped[2:]), h1_style))
        elif stripped.startswith("## "):
            story.append(Paragraph(md_to_pdf_html(stripped[3:]), h2_style))
        elif stripped.startswith("### "):
            story.append(Paragraph(md_to_pdf_html(stripped[4:]), h3_style))
        elif re.match(r'^[\*\-\+]\s+', stripped):
            content = re.sub(r'^[\*\-\+]\s+', '', stripped)
            story.append(Paragraph(f"• {md_to_pdf_html(content)}", bullet_style))
        elif re.match(r'^\d+\.\s+', stripped):
            story.append(Paragraph(md_to_pdf_html(stripped), bullet_style))
        else:
            story.append(Paragraph(md_to_pdf_html(stripped), body_style))

    doc.build(story)
    bio.seek(0)
    return bio

@app.route('/api/validate-key', methods=['POST'])
def validate_key():
    data = request.get_json(silent=True) or {}
    user_key = data.get('api_key', '').strip()
    active_key = get_active_api_key(user_key)
    
    if not active_key:
        return jsonify({
            "valid": False,
            "message": "No API key configured. Enter a key or configure server GEMINI_API_KEY."
        }), 400
    
    try:
        genai.configure(api_key=active_key)
        models = []
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                models.append(m.name.replace("models/", ""))
        return jsonify({
            "valid": True,
            "is_custom": bool(user_key),
            "model_count": len(models),
            "message": f"API Key Validated! {len(models)} models available."
        })
    except Exception as e:
        return jsonify({
            "valid": False,
            "message": f"API Key Validation Error: {str(e)}"
        }), 400

@app.route('/api/models', methods=['POST'])
def get_models():
    data = request.get_json(silent=True) or {}
    user_key = data.get('api_key', '')
    active_key = get_active_api_key(user_key)
    models = fetch_gemini_models(active_key)
    return jsonify({
        "models": models,
        "has_key": bool(active_key),
        "is_custom": bool(user_key and user_key.strip())
    })

@app.route('/api/generate', methods=['POST'])
def generate_docs():
    data = request.get_json(silent=True) or {}
    code = data.get('code', '')
    model_name = data.get('model', 'gemini-2.5-flash')
    user_key = data.get('api_key', '')

    if not code or not code.strip():
        return jsonify({"error": "No source code provided"}), 400

    active_key = get_active_api_key(user_key)
    if not active_key:
        return jsonify({"error": "No Gemini API Key found. Please click 'Key Settings' in the top bar to add your key."}), 400

    try:
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel(model_name)

        prompt = f"""
You are an expert software documentation engineer.

Analyze the following source code and generate detailed professional documentation.

Include the following sections:
1. Project Overview
2. Purpose
3. Features
4. Technologies Used
5. Function Description
6. Class Description
7. Inputs and Outputs
8. Installation Guide
9. Example Usage
10. Conclusion

Source Code:

{code}
"""
        response = model.generate_content(prompt)
        documentation = response.text

        doc_id = str(uuid.uuid4())
        DOC_CACHE[doc_id] = documentation

        return jsonify({
            "success": True,
            "token": doc_id,
            "documentation": documentation
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/download/<format_type>', methods=['GET'])
def download_file(format_type):
    token = request.args.get('token', '')
    documentation = DOC_CACHE.get(token)

    if not documentation:
        return "Invalid or expired download token.", 404

    if format_type == 'docx':
        bio = generate_docx_bytes(documentation)
        return send_file(
            bio,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="documentation.docx"
        )
    elif format_type == 'pdf':
        bio = generate_pdf_bytes(documentation)
        return send_file(
            bio,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="documentation.pdf"
        )
    elif format_type == 'md':
        bio = io.BytesIO(documentation.encode('utf-8'))
        bio.seek(0)
        return send_file(
            bio,
            mimetype="text/markdown",
            as_attachment=True,
            download_name="documentation.md"
        )
    elif format_type == 'zip':
        docx_bio = generate_docx_bytes(documentation)
        pdf_bio = generate_pdf_bytes(documentation)

        zip_bio = io.BytesIO()
        with zipfile.ZipFile(zip_bio, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.writestr("documentation.md", documentation.encode('utf-8'))
            zipf.writestr("documentation.docx", docx_bio.getvalue())
            zipf.writestr("documentation.pdf", pdf_bio.getvalue())
        zip_bio.seek(0)

        return send_file(
            zip_bio,
            mimetype="application/zip",
            as_attachment=True,
            download_name="documentation.zip"
        )

    return "Invalid format specified.", 400

@app.route('/')
def home():
    html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NeuroDocs AI — Next-Gen Code Documentation Generator</title>
    <meta name="description" content="Transform source code into high-grade developer documentation instantly using Gemini AI. Export to PDF, DOCX, Markdown & ZIP.">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600&family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --bg-canvas: #050811;
            --bg-card: rgba(15, 23, 42, 0.75);
            --bg-card-hover: rgba(30, 41, 59, 0.85);
            --border-subtle: rgba(255, 255, 255, 0.08);
            --border-glow: rgba(56, 189, 248, 0.3);
            
            --primary-cyan: #38bdf8;
            --primary-glow: rgba(56, 189, 248, 0.35);
            --accent-purple: #818cf8;
            --accent-pink: #f472b6;
            --emerald-green: #34d399;
            --amber-gold: #fbbf24;
            
            --text-heading: #f8fafc;
            --text-body: #cbd5e1;
            --text-muted: #64748b;
            --editor-bg: #090d16;
            --radius-lg: 18px;
            --radius-md: 12px;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: var(--bg-canvas);
            color: var(--text-body);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            overflow-x: hidden;
            background-image: 
                radial-gradient(circle at 10% 10%, rgba(56, 189, 248, 0.1) 0%, transparent 40%),
                radial-gradient(circle at 90% 20%, rgba(129, 140, 248, 0.12) 0%, transparent 45%),
                radial-gradient(circle at 50% 90%, rgba(244, 114, 182, 0.08) 0%, transparent 50%);
            background-attachment: fixed;
        }

        /* Ambient Top Glow Line */
        .ambient-bar {
            height: 3px;
            width: 100%;
            background: linear-gradient(90deg, var(--primary-cyan), var(--accent-purple), var(--accent-pink));
            box-shadow: 0 0 12px var(--primary-cyan);
        }

        /* Header Navigation */
        header {
            background: rgba(9, 13, 22, 0.85);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border-bottom: 1px solid var(--border-subtle);
            padding: 1rem 2rem;
            position: sticky;
            top: 0;
            z-index: 100;
        }

        .header-wrap {
            max-width: 1440px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1rem;
        }

        .brand-logo {
            display: flex;
            align-items: center;
            gap: 0.85rem;
            text-decoration: none;
        }

        .logo-box {
            width: 44px;
            height: 44px;
            border-radius: 12px;
            background: linear-gradient(135deg, var(--primary-cyan), var(--accent-purple));
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.4rem;
            color: white;
            box-shadow: 0 0 20px var(--primary-glow);
            position: relative;
        }

        .brand-text h1 {
            font-size: 1.4rem;
            font-weight: 800;
            letter-spacing: -0.02em;
            background: linear-gradient(135deg, #ffffff 30%, var(--primary-cyan));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .brand-text p {
            font-size: 0.78rem;
            color: var(--text-muted);
            font-weight: 500;
        }

        .header-actions {
            display: flex;
            align-items: center;
            gap: 0.85rem;
            flex-wrap: wrap;
        }

        /* Sleek Control Pills */
        .ctrl-pill {
            background: rgba(15, 23, 42, 0.9);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            padding: 0.45rem 0.9rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
            font-size: 0.88rem;
            color: var(--text-heading);
            cursor: pointer;
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .ctrl-pill:hover {
            border-color: var(--primary-cyan);
            box-shadow: 0 0 15px var(--primary-glow);
            transform: translateY(-1px);
        }

        .ctrl-pill select {
            background: transparent;
            border: none;
            color: var(--text-heading);
            font-family: inherit;
            font-size: 0.88rem;
            font-weight: 600;
            outline: none;
            cursor: pointer;
        }

        .ctrl-pill select option {
            background: #0f172a;
            color: white;
        }

        /* Pulsing Key Status Badge Button */
        .key-badge-btn {
            background: rgba(15, 23, 42, 0.9);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            padding: 0.5rem 1.1rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
            font-size: 0.88rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
        }

        .key-badge-btn.custom {
            border-color: rgba(52, 211, 153, 0.4);
            background: rgba(52, 211, 153, 0.08);
            color: var(--emerald-green);
        }

        .key-badge-btn.server {
            border-color: rgba(56, 189, 248, 0.4);
            background: rgba(56, 189, 248, 0.08);
            color: var(--primary-cyan);
        }

        .key-badge-btn.missing {
            border-color: rgba(251, 191, 36, 0.4);
            background: rgba(251, 191, 36, 0.08);
            color: var(--amber-gold);
        }

        .pulse-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            display: inline-block;
            box-shadow: 0 0 8px currentColor;
            animation: pulse-glow 2s infinite;
        }

        @keyframes pulse-glow {
            0% { opacity: 0.4; transform: scale(0.9); }
            50% { opacity: 1; transform: scale(1.1); }
            100% { opacity: 0.4; transform: scale(0.9); }
        }

        /* Modal Overlay & Card */
        .modal-overlay {
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(3, 7, 18, 0.85);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            z-index: 1000;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 1.5rem;
            opacity: 0;
            pointer-events: none;
            transition: opacity 0.3s ease;
        }

        .modal-overlay.open {
            opacity: 1;
            pointer-events: auto;
        }

        .modal-card {
            background: #0f172a;
            border: 1px solid var(--border-glow);
            border-radius: 20px;
            max-width: 540px;
            width: 100%;
            padding: 2rem;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7), 0 0 30px var(--primary-glow);
            transform: scale(0.95);
            transition: transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
            position: relative;
        }

        .modal-overlay.open .modal-card {
            transform: scale(1);
        }

        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1.25rem;
        }

        .modal-title {
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--text-heading);
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }

        .modal-close-btn {
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-subtle);
            color: var(--text-muted);
            width: 32px; height: 32px;
            border-radius: 50%;
            display: flex; align-items: center; justify-content: center;
            cursor: pointer;
            transition: all 0.2s;
        }
        .modal-close-btn:hover { color: white; background: rgba(255, 255, 255, 0.15); }

        .key-input-box {
            background: var(--editor-bg);
            border: 1.5px solid var(--border-subtle);
            border-radius: 12px;
            padding: 0.75rem 1rem;
            display: flex;
            align-items: center;
            gap: 0.75rem;
            margin: 1rem 0;
            transition: border-color 0.2s;
        }

        .key-input-box:focus-within {
            border-color: var(--primary-cyan);
            box-shadow: 0 0 0 3px var(--primary-glow);
        }

        .key-input-box input {
            background: transparent;
            border: none;
            color: var(--text-heading);
            font-family: 'Fira Code', monospace;
            font-size: 0.95rem;
            width: 100%;
            outline: none;
        }

        .eye-toggle {
            color: var(--text-muted);
            cursor: pointer;
            transition: color 0.2s;
        }
        .eye-toggle:hover { color: var(--primary-cyan); }

        .info-callout {
            background: rgba(56, 189, 248, 0.06);
            border-left: 3.5px solid var(--primary-cyan);
            padding: 0.85rem 1rem;
            border-radius: 0 10px 10px 0;
            font-size: 0.85rem;
            color: var(--text-body);
            line-height: 1.5;
            margin-bottom: 1.5rem;
        }

        .info-callout a {
            color: var(--primary-cyan);
            text-decoration: none;
            font-weight: 600;
        }

        .modal-actions {
            display: flex;
            gap: 0.75rem;
        }

        /* Buttons */
        .btn {
            border: none;
            border-radius: 12px;
            padding: 0.75rem 1.25rem;
            font-family: inherit;
            font-size: 0.92rem;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
            transition: all 0.25s;
        }

        .btn-accent {
            background: linear-gradient(135deg, var(--primary-cyan), var(--accent-purple));
            color: white;
            box-shadow: 0 4px 18px var(--primary-glow);
        }

        .btn-accent:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 24px var(--primary-glow);
            filter: brightness(1.1);
        }

        .btn-outline {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border-subtle);
            color: var(--text-heading);
        }

        .btn-outline:hover {
            background: rgba(255, 255, 255, 0.08);
            border-color: var(--text-muted);
        }

        /* Main Workspace Grid */
        main {
            max-width: 1440px;
            width: 100%;
            margin: 2rem auto;
            padding: 0 1.5rem;
            flex: 1;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1.75rem;
        }

        @media (max-width: 1024px) {
            main { grid-template-columns: 1fr; }
        }

        /* Panel Cards */
        .workspace-panel {
            background: var(--bg-card);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-lg);
            padding: 1.75rem;
            display: flex;
            flex-direction: column;
            gap: 1.25rem;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4);
            transition: border-color 0.3s;
        }

        .panel-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-bottom: 0.85rem;
            border-bottom: 1px solid var(--border-subtle);
        }

        .panel-heading {
            font-size: 1.15rem;
            font-weight: 700;
            color: var(--text-heading);
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }

        .panel-heading i { color: var(--primary-cyan); }

        /* Drag & Drop Upload Zone */
        .drop-zone {
            border: 2px dashed rgba(56, 189, 248, 0.25);
            background: rgba(9, 13, 22, 0.6);
            border-radius: var(--radius-md);
            padding: 1.5rem;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .drop-zone:hover, .drop-zone.drag-active {
            border-color: var(--primary-cyan);
            background: rgba(56, 189, 248, 0.08);
            box-shadow: 0 0 20px var(--primary-glow);
        }

        .drop-zone i {
            font-size: 2.2rem;
            color: var(--primary-cyan);
            margin-bottom: 0.6rem;
            display: inline-block;
            transition: transform 0.3s;
        }

        .drop-zone:hover i { transform: translateY(-4px); }

        .drop-zone p { font-size: 0.92rem; color: var(--text-body); }
        .drop-zone span { color: var(--primary-cyan); font-weight: 600; }

        /* Code Editor */
        .editor-wrap {
            position: relative;
            flex: 1;
        }

        textarea.code-area {
            width: 100%;
            height: 420px;
            background: var(--editor-bg);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
            padding: 1.2rem;
            color: #e2e8f0;
            font-family: 'Fira Code', monospace;
            font-size: 0.9rem;
            line-height: 1.6;
            outline: none;
            resize: vertical;
            transition: all 0.25s;
        }

        textarea.code-area:focus {
            border-color: var(--primary-cyan);
            box-shadow: 0 0 0 2px var(--primary-glow);
        }

        /* Large Generate Action Button */
        .btn-generate {
            width: 100%;
            padding: 1rem;
            font-size: 1.05rem;
            font-weight: 700;
            letter-spacing: 0.01em;
            border-radius: var(--radius-md);
            background: linear-gradient(135deg, var(--primary-cyan), var(--accent-purple));
            color: white;
            border: none;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.75rem;
            box-shadow: 0 6px 25px var(--primary-glow);
            transition: all 0.3s;
            position: relative;
            overflow: hidden;
        }

        .btn-generate:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 30px var(--primary-glow);
            filter: brightness(1.15);
        }

        .btn-generate:disabled {
            opacity: 0.65;
            cursor: not-allowed;
            transform: none;
        }

        /* Output Tabs & Box */
        .tab-group {
            display: flex;
            gap: 0.5rem;
            background: rgba(9, 13, 22, 0.7);
            padding: 0.3rem;
            border-radius: 10px;
            border: 1px solid var(--border-subtle);
        }

        .tab-btn {
            background: transparent;
            border: none;
            color: var(--text-muted);
            padding: 0.45rem 1rem;
            font-family: inherit;
            font-size: 0.88rem;
            font-weight: 600;
            border-radius: 7px;
            cursor: pointer;
            transition: all 0.2s;
        }

        .tab-btn.active {
            background: rgba(56, 189, 248, 0.18);
            color: var(--primary-cyan);
        }

        .output-viewport {
            height: 420px;
            overflow-y: auto;
            background: var(--editor-bg);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
            padding: 1.5rem;
            font-size: 0.95rem;
            line-height: 1.7;
            color: var(--text-body);
        }

        .output-viewport h1, .output-viewport h2, .output-viewport h3 {
            color: var(--text-heading);
            margin-top: 1.4rem;
            margin-bottom: 0.7rem;
            font-weight: 700;
        }
        .output-viewport h1 {
            font-size: 1.5rem;
            color: var(--primary-cyan);
            border-bottom: 1px solid var(--border-subtle);
            padding-bottom: 0.5rem;
        }
        .output-viewport h2 { font-size: 1.25rem; color: var(--accent-purple); }
        .output-viewport h3 { font-size: 1.1rem; color: var(--text-heading); }
        .output-viewport code {
            background: rgba(255, 255, 255, 0.08);
            color: var(--accent-pink);
            padding: 0.2rem 0.45rem;
            border-radius: 6px;
            font-family: 'Fira Code', monospace;
            font-size: 0.86rem;
        }
        .output-viewport pre {
            background: #050811;
            padding: 1.2rem;
            border-radius: 10px;
            overflow-x: auto;
            margin: 1rem 0;
            border: 1px solid var(--border-subtle);
        }
        .output-viewport pre code { background: transparent; color: #e2e8f0; padding: 0; }
        .output-viewport ul, .output-viewport ol { padding-left: 1.5rem; margin: 0.8rem 0; }
        .output-viewport li { margin-bottom: 0.4rem; }

        /* Export Dock */
        .export-dock {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 0.75rem;
        }

        @media (max-width: 640px) {
            .export-dock { grid-template-columns: repeat(2, 1fr); }
        }

        .export-card {
            background: rgba(15, 23, 42, 0.8);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            padding: 0.75rem;
            text-align: center;
            text-decoration: none;
            color: var(--text-heading);
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 0.3rem;
            transition: all 0.25s;
        }

        .export-card i { font-size: 1.4rem; margin-bottom: 0.2rem; }
        .export-card span { font-size: 0.85rem; font-weight: 600; }
        .export-card small { font-size: 0.72rem; color: var(--text-muted); }

        .export-card:hover {
            border-color: var(--primary-cyan);
            background: rgba(56, 189, 248, 0.12);
            transform: translateY(-2px);
            box-shadow: 0 6px 18px var(--primary-glow);
        }

        .export-card.disabled {
            opacity: 0.35;
            pointer-events: none;
        }

        /* Footer */
        footer {
            text-align: center;
            padding: 1.75rem;
            color: var(--text-muted);
            font-size: 0.85rem;
            border-top: 1px solid var(--border-subtle);
            background: rgba(5, 8, 17, 0.95);
        }

        /* Loading Spinner */
        .spin { animation: spin 0.8s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>

    <div class="ambient-bar"></div>

    <header>
        <div class="header-wrap">
            <a href="#" class="brand-logo">
                <div class="logo-box"><i class="fa-solid fa-brain"></i></div>
                <div class="brand-text">
                    <h1>NeuroDocs AI</h1>
                    <p>Documentation Engine v2.0</p>
                </div>
            </a>

            <div class="header-actions">
                <!-- Interactive Key Badge Button -->
                <button id="keyBadgeBtn" class="key-badge-btn missing" onclick="openKeyModal()">
                    <span class="pulse-dot"></span>
                    <span id="keyBadgeText">Checking Key...</span>
                    <i class="fa-solid fa-gear" style="font-size: 0.85rem; opacity: 0.7;"></i>
                </button>

                <!-- Model Selector -->
                <div class="ctrl-pill" title="Selected Gemini Model">
                    <i class="fa-solid fa-microchip" style="color: var(--accent-purple);"></i>
                    <select id="modelSelect">
                        <option value="gemini-2.5-flash">gemini-2.5-flash</option>
                        <option value="gemini-2.0-flash">gemini-2.0-flash</option>
                        <option value="gemini-1.5-flash">gemini-1.5-flash</option>
                        <option value="gemini-2.5-pro">gemini-2.5-pro</option>
                    </select>
                </div>
            </div>
        </div>
    </header>

    <!-- Key Settings Modal -->
    <div id="keyModal" class="modal-overlay">
        <div class="modal-card">
            <div class="modal-header">
                <div class="modal-title"><i class="fa-solid fa-key" style="color: var(--primary-cyan);"></i> Gemini API Key Settings</div>
                <button class="modal-close-btn" onclick="closeKeyModal()"><i class="fa-solid fa-xmark"></i></button>
            </div>

            <p style="font-size: 0.9rem; color: var(--text-body);">Enter your personal Gemini API Key below. Your key is stored securely in your local browser storage and used for processing requests.</p>

            <div class="key-input-box">
                <i class="fa-solid fa-lock" style="color: var(--text-muted);"></i>
                <input type="password" id="userApiKey" placeholder="Paste AIzaSy... key here">
                <i id="eyeToggle" class="fa-solid fa-eye eye-toggle" onclick="toggleKeyVisibility()"></i>
            </div>

            <div id="validationAlert" style="display: none; font-size: 0.85rem; padding: 0.6rem 0.8rem; border-radius: 8px; margin-bottom: 1rem;"></div>

            <div class="info-callout">
                💡 <strong>Don't have a Gemini API key?</strong><br>
                You can generate a free API Key instantly at <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener">Google AI Studio <i class="fa-solid fa-arrow-up-right-from-square" style="font-size: 0.75rem;"></i></a>.
            </div>

            <div class="modal-actions">
                <button class="btn btn-outline" style="flex: 1;" onclick="testKeyConnection()"><i class="fa-solid fa-vial"></i> Test Key</button>
                <button class="btn btn-accent" style="flex: 1;" onclick="saveKeyAndClose()"><i class="fa-solid fa-floppy-disk"></i> Save & Apply</button>
                <button class="btn btn-outline" onclick="clearVisitorKey()"><i class="fa-solid fa-trash-can"></i></button>
            </div>
        </div>
    </div>

    <main>
        <!-- Left: Code Input Panel -->
        <section class="workspace-panel">
            <div class="panel-top">
                <div class="panel-heading"><i class="fa-solid fa-code-commit"></i> Source Code Input</div>
                <button class="tab-btn" onclick="loadSampleCode()"><i class="fa-solid fa-wand-magic-sparkles"></i> Sample Code</button>
            </div>

            <div class="drop-zone" id="dropArea">
                <input type="file" id="fileInput" accept=".py,.js,.java,.cpp,.c,.html,.css,.txt,.go,.rs,.sql" style="display: none;">
                <i class="fa-solid fa-file-arrow-up"></i>
                <p>Drag & drop code file here or <span>browse file</span></p>
            </div>

            <div class="editor-wrap">
                <textarea id="codeEditor" class="code-area" placeholder="// Paste source code here or drop a file..."></textarea>
            </div>

            <button id="generateBtn" class="btn-generate" onclick="generateDocs()">
                <i class="fa-solid fa-bolt"></i> Generate AI Documentation
            </button>
        </section>

        <!-- Right: AI Documentation Hub Panel -->
        <section class="workspace-panel">
            <div class="panel-top">
                <div class="panel-heading"><i class="fa-solid fa-file-invoice"></i> Documentation Hub</div>
                <div class="tab-group">
                    <button class="tab-btn active" id="tabRendered" onclick="switchTab('rendered')">Preview</button>
                    <button class="tab-btn" id="tabRaw" onclick="switchTab('raw')">Raw Markdown</button>
                </div>
            </div>

            <div class="output-viewport" id="outputRendered">
                <div style="text-align: center; margin-top: 6rem; color: var(--text-muted);">
                    <i class="fa-solid fa-sparkles" style="font-size: 2.8rem; margin-bottom: 1rem; color: var(--primary-cyan); opacity: 0.8;"></i>
                    <p style="font-size: 1rem; font-weight: 500;">Ready to generate documentation.</p>
                    <p style="font-size: 0.85rem; color: var(--text-muted); margin-top: 0.4rem;">Upload or paste your code on the left and click Generate.</p>
                </div>
            </div>

            <textarea id="outputRaw" class="code-area" style="display: none; height: 420px;" readonly></textarea>

            <div>
                <p style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.6rem; font-weight: 600; letter-spacing: 0.04em;">INSTANT EXPORTS:</p>
                <div class="export-dock">
                    <a id="btnDocx" class="export-card disabled" href="#"><i class="fa-solid fa-file-word" style="color: #3b82f6;"></i><span>DOCX</span><small>Word Format</small></a>
                    <a id="btnPdf" class="export-card disabled" href="#"><i class="fa-solid fa-file-pdf" style="color: #ef4444;"></i><span>PDF</span><small>Styled Document</small></a>
                    <a id="btnMd" class="export-card disabled" href="#"><i class="fa-solid fa-file-code" style="color: #10b981;"></i><span>Markdown</span><small>Raw .md</small></a>
                    <a id="btnZip" class="export-card disabled" href="#"><i class="fa-solid fa-file-zipper" style="color: #f59e0b;"></i><span>ZIP</span><small>All Formats</small></a>
                </div>
            </div>
        </section>
    </main>

    <footer>
        NeuroDocs AI &bull; Production Engine for Vercel Serverless &bull; Developed by Dani Toffin
    </footer>

    <script>
        const apiKeyInput = document.getElementById('userApiKey');
        const keyBadgeBtn = document.getElementById('keyBadgeBtn');
        const keyBadgeText = document.getElementById('keyBadgeText');
        const modelSelect = document.getElementById('modelSelect');
        const keyModal = document.getElementById('keyModal');
        const eyeToggle = document.getElementById('eyeToggle');
        const validationAlert = document.getElementById('validationAlert');
        const codeEditor = document.getElementById('codeEditor');
        const fileInput = document.getElementById('fileInput');
        const dropArea = document.getElementById('dropArea');
        const generateBtn = document.getElementById('generateBtn');
        const outputRendered = document.getElementById('outputRendered');
        const outputRaw = document.getElementById('outputRaw');
        let currentToken = null;

        // Restore saved key from localStorage
        const savedKey = localStorage.getItem('neurodocs_gemini_key');
        if (savedKey) {
            apiKeyInput.value = savedKey;
        }

        function openKeyModal() { keyModal.classList.add('open'); }
        function closeKeyModal() { keyModal.classList.remove('open'); }
        function toggleKeyVisibility() {
            if (apiKeyInput.type === 'password') {
                apiKeyInput.type = 'text';
                eyeToggle.className = 'fa-solid fa-eye-slash eye-toggle';
            } else {
                apiKeyInput.type = 'password';
                eyeToggle.className = 'fa-solid fa-eye eye-toggle';
            }
        }

        async function fetchModels() {
            try {
                const res = await fetch('/api/models', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ api_key: apiKeyInput.value.trim() })
                });
                const data = await res.json();
                if (data.models && data.models.length > 0) {
                    modelSelect.innerHTML = data.models.map(m => `<option value="${m}">${m}</option>`).join('');
                }
                if (data.has_key) {
                    if (data.is_custom) {
                        keyBadgeBtn.className = 'key-badge-btn custom';
                        keyBadgeText.textContent = 'Visitor Key Active';
                    } else {
                        keyBadgeBtn.className = 'key-badge-btn server';
                        keyBadgeText.textContent = 'Server Key Active';
                    }
                } else {
                    keyBadgeBtn.className = 'key-badge-btn missing';
                    keyBadgeText.textContent = 'Set API Key';
                }
            } catch (err) { console.error(err); }
        }
        fetchModels();

        async function testKeyConnection() {
            validationAlert.style.display = 'block';
            validationAlert.style.background = 'rgba(56, 189, 248, 0.15)';
            validationAlert.style.color = '#38bdf8';
            validationAlert.innerHTML = `<i class="fa-solid fa-spinner spin"></i> Testing Key Connection...`;

            try {
                const res = await fetch('/api/validate-key', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ api_key: apiKeyInput.value.trim() })
                });
                const data = await res.json();
                if (data.valid) {
                    validationAlert.style.background = 'rgba(52, 211, 153, 0.15)';
                    validationAlert.style.color = '#34d399';
                    validationAlert.innerHTML = `<i class="fa-solid fa-circle-check"></i> ${data.message}`;
                } else {
                    validationAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                    validationAlert.style.color = '#f87171';
                    validationAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> ${data.message}`;
                }
            } catch (err) {
                validationAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                validationAlert.style.color = '#f87171';
                validationAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> Connection Error.`;
            }
        }

        function saveKeyAndClose() {
            localStorage.setItem('neurodocs_gemini_key', apiKeyInput.value.trim());
            fetchModels();
            closeKeyModal();
        }

        function clearVisitorKey() {
            apiKeyInput.value = '';
            localStorage.removeItem('neurodocs_gemini_key');
            validationAlert.style.display = 'none';
            fetchModels();
        }

        // File Drag and Drop
        dropArea.addEventListener('click', () => fileInput.click());
        dropArea.addEventListener('dragover', (e) => { e.preventDefault(); dropArea.classList.add('drag-active'); });
        dropArea.addEventListener('dragleave', () => dropArea.classList.remove('drag-active'));
        dropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            dropArea.classList.remove('drag-active');
            if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
        });
        fileInput.addEventListener('change', (e) => { if (e.target.files.length) handleFile(e.target.files[0]); });

        function handleFile(file) {
            const reader = new FileReader();
            reader.onload = (e) => { codeEditor.value = e.target.result; };
            reader.readAsText(file);
        }

        function loadSampleCode() {
            codeEditor.value = `def calculate_fibonacci(n):\n    \"\"\"Calculates Fibonacci sequence up to n numbers.\"\"\"\n    if n <= 0:\n        return []\n    elif n == 1:\n        return [0]\n    \n    sequence = [0, 1]\n    while len(sequence) < n:\n        sequence.append(sequence[-1] + sequence[-2])\n    return sequence\n\nif __name__ == "__main__":\n    result = calculate_fibonacci(10)\n    print("Fibonacci:", result)`;
        }

        function switchTab(tab) {
            if (tab === 'rendered') {
                document.getElementById('tabRendered').classList.add('active');
                document.getElementById('tabRaw').classList.remove('active');
                outputRendered.style.display = 'block';
                outputRaw.style.display = 'none';
            } else {
                document.getElementById('tabRaw').classList.add('active');
                document.getElementById('tabRendered').classList.remove('active');
                outputRaw.style.display = 'block';
                outputRendered.style.display = 'none';
            }
        }

        async function generateDocs() {
            const code = codeEditor.value.trim();
            if (!code) {
                alert('Please enter or upload source code first.');
                return;
            }

            generateBtn.disabled = true;
            generateBtn.innerHTML = `<i class="fa-solid fa-spinner spin"></i> Analyzing & Generating AI Documentation...`;

            try {
                const res = await fetch('/api/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        code: code,
                        model: modelSelect.value,
                        api_key: apiKeyInput.value.trim()
                    })
                });

                const data = await res.json();
                if (!res.ok) throw new Error(data.error || 'Generation failed');

                const docText = data.documentation;
                currentToken = data.token;

                outputRendered.innerHTML = marked.parse(docText);
                outputRaw.value = docText;

                ['Docx', 'Pdf', 'Md', 'Zip'].forEach(fmt => {
                    const card = document.getElementById('btn' + fmt);
                    card.href = `/api/download/${fmt.toLowerCase()}?token=${currentToken}`;
                    card.classList.remove('disabled');
                });

            } catch (err) {
                alert('Error: ' + err.message);
            } finally {
                generateBtn.disabled = false;
                generateBtn.innerHTML = `<i class="fa-solid fa-bolt"></i> Generate AI Documentation`;
            }
        }
    </script>
</body>
</html>"""
    return render_template_string(html_template)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
