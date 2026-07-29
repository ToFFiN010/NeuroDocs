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

# Temporary in-memory cache for generated documentation downloads (cleared periodically / session-scoped)
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
        return jsonify({"error": "No Gemini API Key found. Please enter your API key in the top settings bar."}), 400

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
    <title>AI Documentation Generator — NeuroDocs</title>
    <meta name="description" content="Generate professional, structured software documentation automatically from your source code using Gemini AI. Export to PDF, DOCX, Markdown, and ZIP.">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500&family=Outfit:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --bg-dark: #090d16;
            --card-bg: rgba(19, 27, 45, 0.75);
            --card-border: rgba(255, 255, 255, 0.08);
            --primary: #38bdf8;
            --primary-glow: rgba(56, 189, 248, 0.25);
            --accent: #6366f1;
            --accent-glow: rgba(99, 102, 241, 0.25);
            --success: #10b981;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --input-bg: #0f172a;
            --radius: 14px;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Outfit', sans-serif;
            background-color: var(--bg-dark);
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            background-image: 
                radial-gradient(circle at 15% 15%, rgba(99, 102, 241, 0.12) 0%, transparent 40%),
                radial-gradient(circle at 85% 85%, rgba(56, 189, 248, 0.12) 0%, transparent 40%);
            background-attachment: fixed;
        }

        /* Header */
        header {
            background: rgba(15, 23, 42, 0.8);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border-bottom: 1px solid var(--card-border);
            padding: 1rem 2rem;
            position: sticky;
            top: 0;
            z-index: 100;
        }

        .header-container {
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1rem;
        }

        .logo-group {
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        .logo-icon {
            background: linear-gradient(135deg, var(--primary), var(--accent));
            width: 42px;
            height: 42px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.3rem;
            color: white;
            box-shadow: 0 4px 14px var(--primary-glow);
        }

        .logo-text h1 {
            font-size: 1.35rem;
            font-weight: 700;
            background: linear-gradient(to right, #ffffff, var(--primary));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .logo-text p {
            font-size: 0.8rem;
            color: var(--text-muted);
        }

        .controls-group {
            display: flex;
            align-items: center;
            gap: 1rem;
            flex-wrap: wrap;
        }

        .input-pill {
            background: var(--input-bg);
            border: 1px solid var(--card-border);
            border-radius: 10px;
            padding: 0.4rem 0.8rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            font-size: 0.88rem;
            transition: all 0.2s;
        }

        .input-pill:focus-within {
            border-color: var(--primary);
            box-shadow: 0 0 0 3px var(--primary-glow);
        }

        .input-pill input, .input-pill select {
            background: transparent;
            border: none;
            color: var(--text-main);
            font-family: inherit;
            font-size: 0.88rem;
            outline: none;
        }

        .input-pill select option {
            background: #1e293b;
            color: white;
        }

        .status-badge {
            font-size: 0.78rem;
            font-weight: 600;
            padding: 0.3rem 0.75rem;
            border-radius: 20px;
            display: flex;
            align-items: center;
            gap: 0.4rem;
        }

        .status-badge.active {
            background: rgba(16, 185, 129, 0.15);
            color: #34d399;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }

        .status-badge.warning {
            background: rgba(245, 158, 11, 0.15);
            color: #fbbf24;
            border: 1px solid rgba(245, 158, 11, 0.3);
        }

        /* Main Workspace */
        main {
            max-width: 1400px;
            width: 100%;
            margin: 2rem auto;
            padding: 0 1.5rem;
            flex: 1;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1.5rem;
        }

        @media (max-width: 968px) {
            main {
                grid-template-columns: 1fr;
            }
        }

        /* Card Panels */
        .panel {
            background: var(--card-bg);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid var(--card-border);
            border-radius: var(--radius);
            padding: 1.5rem;
            display: flex;
            flex-direction: column;
            gap: 1.25rem;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
        }

        .panel-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--card-border);
            padding-bottom: 0.75rem;
        }

        .panel-title {
            font-size: 1.1rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            color: var(--text-main);
        }

        .panel-title i {
            color: var(--primary);
        }

        /* File Upload Drop Area */
        .upload-area {
            border: 2px dashed rgba(56, 189, 248, 0.3);
            background: rgba(15, 23, 42, 0.5);
            border-radius: 12px;
            padding: 1.5rem;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
        }

        .upload-area:hover, .upload-area.dragover {
            border-color: var(--primary);
            background: rgba(56, 189, 248, 0.08);
        }

        .upload-area i {
            font-size: 2.2rem;
            color: var(--primary);
            margin-bottom: 0.5rem;
        }

        .upload-area p {
            font-size: 0.9rem;
            color: var(--text-muted);
        }

        .upload-area span {
            color: var(--primary);
            font-weight: 600;
        }

        /* Code Editor Input */
        .code-container {
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
        }

        textarea.code-editor {
            width: 100%;
            height: 380px;
            background: var(--input-bg);
            border: 1px solid var(--card-border);
            border-radius: 10px;
            padding: 1rem;
            color: #e2e8f0;
            font-family: 'Fira Code', monospace;
            font-size: 0.88rem;
            line-height: 1.5;
            outline: none;
            resize: vertical;
            transition: border-color 0.2s;
        }

        textarea.code-editor:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 2px var(--primary-glow);
        }

        /* Action Buttons */
        .btn-primary {
            background: linear-gradient(135deg, var(--primary), var(--accent));
            color: white;
            border: none;
            border-radius: 10px;
            padding: 0.85rem 1.5rem;
            font-family: inherit;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.6rem;
            box-shadow: 0 4px 15px var(--primary-glow);
            transition: all 0.25s;
        }

        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px var(--primary-glow);
            filter: brightness(1.1);
        }

        .btn-primary:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }

        /* Output View Tabs */
        .tab-bar {
            display: flex;
            gap: 0.5rem;
            border-bottom: 1px solid var(--card-border);
            padding-bottom: 0.5rem;
        }

        .tab-btn {
            background: transparent;
            border: none;
            color: var(--text-muted);
            padding: 0.4rem 0.9rem;
            font-family: inherit;
            font-size: 0.88rem;
            font-weight: 500;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
        }

        .tab-btn.active {
            background: rgba(56, 189, 248, 0.15);
            color: var(--primary);
        }

        .output-box {
            flex: 1;
            height: 380px;
            overflow-y: auto;
            background: var(--input-bg);
            border: 1px solid var(--card-border);
            border-radius: 10px;
            padding: 1.25rem;
            font-size: 0.92rem;
            line-height: 1.6;
            color: #cbd5e1;
        }

        /* Markdown Styling inside Output Box */
        .output-box h1, .output-box h2, .output-box h3 {
            color: var(--text-main);
            margin-top: 1.2rem;
            margin-bottom: 0.6rem;
        }
        .output-box h1 { font-size: 1.4rem; color: var(--primary); border-bottom: 1px solid var(--card-border); padding-bottom: 0.3rem; }
        .output-box h2 { font-size: 1.2rem; color: #a5b4fc; }
        .output-box h3 { font-size: 1.05rem; }
        .output-box code {
            background: rgba(255,255,255,0.08);
            color: #f472b6;
            padding: 0.15rem 0.4rem;
            border-radius: 4px;
            font-family: 'Fira Code', monospace;
            font-size: 0.85rem;
        }
        .output-box pre {
            background: #090d16;
            padding: 1rem;
            border-radius: 8px;
            overflow-x: auto;
            margin: 0.8rem 0;
            border: 1px solid var(--card-border);
        }
        .output-box pre code {
            background: transparent;
            color: #e2e8f0;
            padding: 0;
        }
        .output-box ul, .output-box ol {
            padding-left: 1.4rem;
            margin: 0.6rem 0;
        }
        .output-box li { margin-bottom: 0.3rem; }

        /* Downloads Bar */
        .download-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 0.6rem;
        }

        @media (max-width: 600px) {
            .download-grid {
                grid-template-columns: repeat(2, 1fr);
            }
        }

        .btn-download {
            background: rgba(30, 41, 59, 0.8);
            border: 1px solid var(--card-border);
            color: var(--text-main);
            padding: 0.6rem 0.8rem;
            border-radius: 8px;
            font-family: inherit;
            font-size: 0.85rem;
            font-weight: 500;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.4rem;
            transition: all 0.2s;
            text-decoration: none;
        }

        .btn-download:hover {
            border-color: var(--primary);
            background: rgba(56, 189, 248, 0.12);
            color: var(--primary);
            transform: translateY(-1px);
        }

        .btn-download.disabled {
            opacity: 0.4;
            pointer-events: none;
        }

        /* Footer */
        footer {
            text-align: center;
            padding: 1.5rem;
            color: var(--text-muted);
            font-size: 0.85rem;
            border-top: 1px solid var(--card-border);
            background: rgba(9, 13, 22, 0.9);
        }

        /* Spinner Animation */
        .spinner {
            width: 18px;
            height: 18px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-radius: 50%;
            border-top-color: white;
            animation: spin 0.8s linear infinite;
        }
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
    </style>
</head>
<body>

    <header>
        <div class="header-container">
            <div class="logo-group">
                <div class="logo-icon"><i class="fa-solid fa-file-code"></i></div>
                <div class="logo-text">
                    <h1>NeuroDocs AI</h1>
                    <p>Production AI Documentation Generator</p>
                </div>
            </div>

            <div class="controls-group">
                <!-- Visitor API Key Input -->
                <div class="input-pill" title="Enter your Gemini API key or leave empty to use server default key">
                    <i class="fa-solid fa-key" style="color: var(--primary);"></i>
                    <input type="password" id="userApiKey" placeholder="Visitor API Key (Optional)">
                </div>

                <!-- Model Selector -->
                <div class="input-pill">
                    <i class="fa-solid fa-microchip" style="color: var(--accent);"></i>
                    <select id="modelSelect">
                        <option value="gemini-2.5-flash">gemini-2.5-flash</option>
                        <option value="gemini-2.0-flash">gemini-2.0-flash</option>
                        <option value="gemini-1.5-flash">gemini-1.5-flash</option>
                        <option value="gemini-2.5-pro">gemini-2.5-pro</option>
                    </select>
                </div>

                <!-- API Key Status Badge -->
                <div id="statusBadge" class="status-badge warning">
                    <i class="fa-solid fa-circle-dot"></i> Checking Key...
                </div>
            </div>
        </div>
    </header>

    <main>
        <!-- Left Panel: Input & Code Upload -->
        <section class="panel">
            <div class="panel-header">
                <div class="panel-title"><i class="fa-solid fa-code"></i> Source Code Input</div>
                <button class="tab-btn" onclick="loadSampleCode()"><i class="fa-solid fa-wand-magic-sparkles"></i> Load Sample</button>
            </div>

            <div class="upload-area" id="dropArea">
                <input type="file" id="fileInput" accept=".py,.js,.java,.cpp,.c,.html,.css,.txt" style="display: none;">
                <i class="fa-solid fa-cloud-arrow-up"></i>
                <p>Drag & drop code file here or <span>browse file</span></p>
            </div>

            <div class="code-container">
                <textarea id="codeEditor" class="code-editor" placeholder="// Paste source code here or upload a file..."></textarea>
            </div>

            <button id="generateBtn" class="btn-primary" onclick="generateDocs()">
                <i class="fa-solid fa-bolt"></i> Generate AI Documentation
            </button>
        </section>

        <!-- Right Panel: AI Documentation Output -->
        <section class="panel">
            <div class="panel-header">
                <div class="panel-title"><i class="fa-solid fa-book-open"></i> Generated Documentation</div>
                <div class="tab-bar">
                    <button class="tab-btn active" id="tabRendered" onclick="switchTab('rendered')">Preview</button>
                    <button class="tab-btn" id="tabRaw" onclick="switchTab('raw')">Raw Markdown</button>
                </div>
            </div>

            <div class="output-box" id="outputRendered">
                <div style="text-align: center; margin-top: 5rem; color: var(--text-muted);">
                    <i class="fa-solid fa-sparkles" style="font-size: 2.5rem; margin-bottom: 1rem; color: var(--primary);"></i>
                    <p>Your AI-generated documentation will appear here.</p>
                </div>
            </div>

            <textarea id="outputRaw" class="code-editor" style="display: none; height: 380px;" readonly></textarea>

            <!-- Export Buttons -->
            <div>
                <p style="font-size: 0.8rem; color: var(--text-muted); margin-bottom: 0.5rem; font-weight: 500;">EXPORT DOCUMENTATION:</p>
                <div class="download-grid">
                    <a id="btnDocx" class="btn-download disabled" href="#"><i class="fa-solid fa-file-word" style="color: #3b82f6;"></i> DOCX</a>
                    <a id="btnPdf" class="btn-download disabled" href="#"><i class="fa-solid fa-file-pdf" style="color: #ef4444;"></i> PDF</a>
                    <a id="btnMd" class="btn-download disabled" href="#"><i class="fa-solid fa-file-lines" style="color: #10b981;"></i> Markdown</a>
                    <a id="btnZip" class="btn-download disabled" href="#"><i class="fa-solid fa-file-zipper" style="color: #f59e0b;"></i> ZIP Bundle</a>
                </div>
            </div>
        </section>
    </main>

    <footer>
        Developed by Dani Toffin &bull; AI Documentation Generator &bull; Powered by Gemini AI &amp; Vercel Serverless
    </footer>

    <script>
        const apiKeyInput = document.getElementById('userApiKey');
        const modelSelect = document.getElementById('modelSelect');
        const statusBadge = document.getElementById('statusBadge');
        const codeEditor = document.getElementById('codeEditor');
        const fileInput = document.getElementById('fileInput');
        const dropArea = document.getElementById('dropArea');
        const generateBtn = document.getElementById('generateBtn');
        const outputRendered = document.getElementById('outputRendered');
        const outputRaw = document.getElementById('outputRaw');
        let currentToken = null;

        // Restore saved visitor API key from localStorage
        const savedKey = localStorage.getItem('neurodocs_gemini_key');
        if (savedKey) {
            apiKeyInput.value = savedKey;
        }

        apiKeyInput.addEventListener('input', () => {
            localStorage.setItem('neurodocs_gemini_key', apiKeyInput.value.trim());
            fetchModels();
        });

        // Fetch Gemini Models from Backend
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
                    statusBadge.className = 'status-badge active';
                    statusBadge.innerHTML = `<i class="fa-solid fa-circle-check"></i> ${data.is_custom ? 'Visitor Key Active' : 'Server Key Active'}`;
                } else {
                    statusBadge.className = 'status-badge warning';
                    statusBadge.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Key Required`;
                }
            } catch (err) {
                console.error(err);
            }
        }
        fetchModels();

        // Drag & Drop File Upload
        dropArea.addEventListener('click', () => fileInput.click());
        dropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropArea.classList.add('dragover');
        });
        dropArea.addEventListener('dragleave', () => dropArea.classList.remove('dragover'));
        dropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            dropArea.classList.remove('dragover');
            if (e.dataTransfer.files.length) {
                handleFile(e.dataTransfer.files[0]);
            }
        });
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) {
                handleFile(e.target.files[0]);
            }
        });

        function handleFile(file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                codeEditor.value = e.target.result;
            };
            reader.readAsText(file);
        }

        function loadSampleCode() {
            codeEditor.value = `def calculate_fibonacci(n):\n    \"\"\"Calculates Fibonacci sequence up to n numbers.\"\"\"\n    if n <= 0:\n        return []\n    elif n == 1:\n        return [0]\n    \n    sequence = [0, 1]\n    while len(sequence) < n:\n        sequence.append(sequence[-1] + sequence[-2])\n    return sequence\n\nif __name__ == "__main__":
    result = calculate_fibonacci(10)
    print("Fibonacci:", result)`;
        }

        // Tab Switching
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

        // Generate Documentation API Call
        async function generateDocs() {
            const code = codeEditor.value.trim();
            if (!code) {
                alert('Please enter or upload source code first.');
                return;
            }

            generateBtn.disabled = true;
            generateBtn.innerHTML = `<div class="spinner"></div> Generating Documentation...`;

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
                if (!res.ok) {
                    throw new Error(data.error || 'Failed to generate documentation.');
                }

                const docText = data.documentation;
                currentToken = data.token;

                outputRendered.innerHTML = marked.parse(docText);
                outputRaw.value = docText;

                // Enable download links
                ['Docx', 'Pdf', 'Md', 'Zip'].forEach(fmt => {
                    const btn = document.getElementById('btn' + fmt);
                    btn.href = `/api/download/${fmt.toLowerCase()}?token=${currentToken}`;
                    btn.classList.remove('disabled');
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
