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

import hashlib

load_dotenv()

app = Flask(__name__)

# Temporary in-memory cache for generated documentation downloads
DOC_CACHE = {}

# In-memory authentication database & active sessions
USERS_DB = {
    "demo@neurodocs.ai": {
        "name": "Dani Toffin",
        "email": "demo@neurodocs.ai",
        "password_hash": hashlib.sha256("password123".encode()).hexdigest(),
        "created_at": "2026-07-30"
    }
}
USER_SESSIONS = {}

def get_active_api_key(custom_key=None):
    if custom_key and custom_key.strip():
        return custom_key.strip()
    return os.getenv("GEMINI_API_KEY", "").strip()

def fetch_gemini_models(api_key):
    fallback_models = ["gemini-flash-latest", "gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-pro-latest"]
    if not api_key:
        return fallback_models
    try:
        genai.configure(api_key=api_key)
        raw_models = []
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                name = m.name.replace("models/", "")
                if not any(bad in name for bad in ["2.5-flash", "2.5-pro", "1.5-flash", "1.5-pro", "antigravity", "lyria", "robotics", "computer-use"]):
                    raw_models.append(name)
        ordered = [m for m in fallback_models if m in raw_models]
        others = [m for m in raw_models if m not in fallback_models]
        result = ordered + others
        if result:
            return result
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

@app.route('/api/auth/signup', methods=['POST'])
def auth_signup():
    data = request.get_json(silent=True) or {}
    email = data.get('email', '').strip().lower()
    password = data.get('password', '')
    name = data.get('name', '').strip() or email.split('@')[0].capitalize()

    if not email or '@' not in email:
        return jsonify({"success": False, "message": "Please enter a valid email address."}), 400
    if not password or len(password) < 6:
        return jsonify({"success": False, "message": "Password must be at least 6 characters long."}), 400

    if email in USERS_DB:
        return jsonify({"success": False, "message": "An account with this email already exists."}), 400

    pwd_hash = hashlib.sha256(password.encode()).hexdigest()
    USERS_DB[email] = {
        "name": name,
        "email": email,
        "password_hash": pwd_hash,
        "created_at": "2026-07-30"
    }

    session_token = f"neuro_{uuid.uuid4().hex}"
    USER_SESSIONS[session_token] = email

    return jsonify({
        "success": True,
        "message": f"Welcome to NeuroDocs AI, {name}!",
        "token": session_token,
        "user": {"name": name, "email": email}
    })

@app.route('/api/auth/login', methods=['POST'])
def auth_login():
    data = request.get_json(silent=True) or {}
    email = data.get('email', '').strip().lower()
    password = data.get('password', '')

    if not email or not password:
        return jsonify({"success": False, "message": "Email and password are required."}), 400

    user = USERS_DB.get(email)
    if not user:
        return jsonify({"success": False, "message": "Invalid email or password."}), 401

    pwd_hash = hashlib.sha256(password.encode()).hexdigest()
    if user["password_hash"] != pwd_hash:
        return jsonify({"success": False, "message": "Invalid email or password."}), 401

    session_token = f"neuro_{uuid.uuid4().hex}"
    USER_SESSIONS[session_token] = email

    return jsonify({
        "success": True,
        "message": f"Welcome back, {user['name']}!",
        "token": session_token,
        "user": {"name": user["name"], "email": user["email"]}
    })

@app.route('/api/auth/me', methods=['POST'])
def auth_me():
    data = request.get_json(silent=True) or {}
    token = data.get('token', '')
    email = USER_SESSIONS.get(token)

    if not email or email not in USERS_DB:
        return jsonify({"authenticated": False})

    user = USERS_DB[email]
    return jsonify({
        "authenticated": True,
        "user": {"name": user["name"], "email": user["email"]}
    })

@app.route('/api/auth/logout', methods=['POST'])
def auth_logout():
    data = request.get_json(silent=True) or {}
    token = data.get('token', '')
    USER_SESSIONS.pop(token, None)
    return jsonify({"success": True, "message": "Logged out successfully."})

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
    model_name = data.get('model', 'gemini-flash-latest')
    user_key = data.get('api_key', '')
    sections = data.get('sections', [])

    if not code or not code.strip():
        return jsonify({"error": "No source code provided"}), 400

    active_key = get_active_api_key(user_key)
    if not active_key:
        return jsonify({"error": "No Gemini API Key found. Please click 'Key Settings' in the top bar to add your key."}), 400

    try:
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel(model_name)

        if sections and isinstance(sections, list) and len(sections) > 0:
            sections_str = "\n".join([f"{idx+1}. {s}" for idx, s in enumerate(sections)])
        else:
            sections_str = """1. Project Overview & Purpose
2. Key Features & Capabilities
3. Architecture & Structure
4. Function & Method Descriptions
5. Class & Data Models
6. Inputs, Outputs & API Schema
7. Installation & Setup Guide
8. Example Usage & Quickstart"""

        prompt = f"""
You are an expert software documentation engineer.

Analyze the following source code and generate detailed professional documentation.

Include the following specific sections:
{sections_str}

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
    <meta name="description" content="Transform source code into high-grade developer documentation instantly using Gemini AI with Voice Explanation. Export to PDF, DOCX, Markdown & ZIP.">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600&family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --bg-canvas: #030712;
            --bg-card: rgba(15, 23, 42, 0.65);
            --bg-card-hover: rgba(30, 41, 59, 0.8);
            --border-subtle: rgba(255, 255, 255, 0.09);
            --border-glow: rgba(56, 189, 248, 0.35);
            
            --primary-cyan: #38bdf8;
            --primary-glow: rgba(56, 189, 248, 0.4);
            --accent-purple: #818cf8;
            --purple-glow: rgba(129, 140, 248, 0.4);
            --accent-pink: #f472b6;
            --emerald-green: #34d399;
            --amber-gold: #fbbf24;
            
            --text-heading: #f8fafc;
            --text-body: #cbd5e1;
            --text-muted: #64748b;
            --editor-bg: #070b14;
            --radius-lg: 20px;
            --radius-md: 14px;
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
                radial-gradient(circle at 15% 15%, rgba(56, 189, 248, 0.12) 0%, transparent 45%),
                radial-gradient(circle at 85% 25%, rgba(129, 140, 248, 0.14) 0%, transparent 50%),
                radial-gradient(circle at 50% 85%, rgba(244, 114, 182, 0.1) 0%, transparent 55%),
                linear-gradient(to right, rgba(255,255,255,0.015) 1px, transparent 1px),
                linear-gradient(to bottom, rgba(255,255,255,0.015) 1px, transparent 1px);
            background-size: 100% 100%, 100% 100%, 100% 100%, 40px 40px, 40px 40px;
            background-attachment: fixed;
        }

        .ambient-bar {
            height: 3px;
            width: 100%;
            background: linear-gradient(90deg, var(--primary-cyan), var(--accent-purple), var(--accent-pink));
            box-shadow: 0 0 15px var(--primary-cyan);
        }

        header {
            background: rgba(7, 11, 20, 0.85);
            backdrop-filter: blur(24px);
            -webkit-backdrop-filter: blur(24px);
            border-bottom: 1px solid var(--border-subtle);
            padding: 0.9rem 2rem;
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
        }

        .modal-overlay.open .modal-card { transform: scale(1); }

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

        /* Auth System UI Enhancements */
        .auth-tab-bar {
            display: flex;
            background: rgba(9, 13, 22, 0.8);
            border: 1px solid var(--border-subtle);
            border-radius: 14px;
            padding: 4px;
            gap: 4px;
            margin-bottom: 1.25rem;
        }

        .auth-tab-btn {
            flex: 1;
            padding: 0.6rem;
            border-radius: 10px;
            border: none;
            background: transparent;
            color: var(--text-muted);
            font-size: 0.88rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.25s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
        }

        .auth-tab-btn.active {
            background: linear-gradient(135deg, var(--primary-cyan), var(--accent-purple));
            color: white;
            box-shadow: 0 4px 15px var(--primary-glow);
        }

        .auth-tab-btn:hover:not(.active) {
            color: var(--text-heading);
            background: rgba(255, 255, 255, 0.05);
        }

        /* Dashboard Sub-Nav & Panels */
        .sub-nav {
            background: rgba(9, 13, 22, 0.7);
            border-bottom: 1px solid var(--border-subtle);
            padding: 0.6rem 2rem;
        }

        .sub-nav-wrap {
            max-width: 1440px;
            margin: 0 auto;
            display: flex;
            gap: 0.75rem;
            align-items: center;
        }

        .nav-mode-btn {
            background: transparent;
            border: 1px solid transparent;
            border-radius: 10px;
            padding: 0.45rem 1rem;
            font-size: 0.85rem;
            font-weight: 600;
            color: var(--text-muted);
            cursor: pointer;
            transition: all 0.25s ease;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .nav-mode-btn.active {
            background: rgba(56, 189, 248, 0.12);
            border-color: rgba(56, 189, 248, 0.4);
            color: var(--primary-cyan);
        }

        .nav-mode-btn:hover:not(.active) {
            color: var(--text-heading);
            background: rgba(255, 255, 255, 0.04);
        }

        .badge-count {
            background: rgba(129, 140, 248, 0.2);
            color: var(--accent-purple);
            padding: 2px 7px;
            border-radius: 12px;
            font-size: 0.72rem;
            font-weight: 700;
        }

        .dashboard-panel {
            max-width: 1440px;
            margin: 1.5rem auto;
            padding: 0 1.5rem;
            width: 100%;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 1.25rem;
            margin-bottom: 1.5rem;
        }

        .stat-card {
            background: var(--bg-card);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
            padding: 1.25rem;
            display: flex;
            align-items: center;
            gap: 1rem;
            transition: all 0.25s ease;
        }

        .stat-card:hover {
            border-color: var(--border-glow);
            transform: translateY(-2px);
            box-shadow: 0 8px 25px var(--primary-glow);
        }

        .stat-icon {
            width: 52px;
            height: 52px;
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.4rem;
        }

        .stat-info small {
            font-size: 0.72rem;
            font-weight: 700;
            color: var(--text-muted);
            letter-spacing: 0.05em;
        }

        .stat-info h2 {
            font-size: 1.6rem;
            font-weight: 800;
            color: var(--text-heading);
            margin: 0.15rem 0;
        }

        .stat-trend {
            font-size: 0.76rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 0.3rem;
        }

        .dashboard-grid {
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 1.25rem;
        }

        @media (max-width: 992px) {
            .dashboard-grid { grid-template-columns: 1fr; }
        }

        .dash-card {
            background: var(--bg-card);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
            padding: 1.5rem;
        }

        .dash-card-header {
            font-size: 1.05rem;
            font-weight: 700;
            color: var(--text-heading);
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }

        .lang-bar-container {
            height: 12px;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 6px;
            overflow: hidden;
            display: flex;
            margin-bottom: 1rem;
        }

        .lang-bar-segment {
            height: 100%;
            transition: width 0.4s ease;
        }

        .lang-legend {
            display: flex;
            gap: 1.25rem;
            flex-wrap: wrap;
            font-size: 0.8rem;
            color: var(--text-body);
        }

        .history-item {
            background: rgba(9, 13, 22, 0.7);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            padding: 1rem 1.25rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 0.75rem;
            gap: 1rem;
            flex-wrap: wrap;
            transition: all 0.2s ease;
        }

        .history-item:hover {
            border-color: var(--primary-cyan);
            background: rgba(15, 23, 42, 0.85);
        }

        /* Generator Column Section Configurator */
        .generator-config-box {
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
            padding: 1rem;
            margin-top: 1rem;
            margin-bottom: 1rem;
        }

        .config-title {
            font-size: 0.85rem;
            font-weight: 700;
            color: var(--text-heading);
            margin-bottom: 0.75rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .section-chips {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
            gap: 0.5rem;
            margin-bottom: 0.85rem;
        }

        .chip-item {
            background: rgba(9, 13, 22, 0.8);
            border: 1px solid var(--border-subtle);
            border-radius: 8px;
            padding: 0.45rem 0.75rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            font-size: 0.8rem;
            color: var(--text-body);
            cursor: pointer;
            transition: all 0.2s ease;
            user-select: none;
        }

        .chip-item:hover {
            border-color: var(--primary-cyan);
            background: rgba(56, 189, 248, 0.08);
        }

        .chip-item input[type="checkbox"] {
            accent-color: var(--primary-cyan);
            cursor: pointer;
            width: 14px;
            height: 14px;
        }

        .preset-row {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            flex-wrap: wrap;
            border-top: 1px solid var(--border-subtle);
            padding-top: 0.65rem;
        }

        .preset-btn {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border-subtle);
            border-radius: 6px;
            color: var(--text-muted);
            padding: 0.25rem 0.6rem;
            font-size: 0.75rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .preset-btn:hover {
            color: var(--primary-cyan);
            border-color: var(--primary-cyan);
            background: rgba(56, 189, 248, 0.1);
        }

        .key-input-box {
            background: var(--editor-bg);
            border: 1.5px solid var(--border-subtle);
            border-radius: 12px;
            padding: 0.75rem 1rem;
            display: flex;
            align-items: center;
            gap: 0.75rem;
            margin: 1rem 0;
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

        .eye-toggle { color: var(--text-muted); cursor: pointer; transition: color 0.2s; }
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

        .info-callout a { color: var(--primary-cyan); text-decoration: none; font-weight: 600; }

        .modal-actions { display: flex; gap: 0.75rem; }

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

        @media (max-width: 1024px) { main { grid-template-columns: 1fr; } }

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

        .drop-zone {
            border: 2px dashed rgba(56, 189, 248, 0.25);
            background: rgba(9, 13, 22, 0.6);
            border-radius: var(--radius-md);
            padding: 1.5rem;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
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
        }

        .drop-zone p { font-size: 0.92rem; color: var(--text-body); }
        .drop-zone span { color: var(--primary-cyan); font-weight: 600; }

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
        }

        textarea.code-area:focus {
            border-color: var(--primary-cyan);
            box-shadow: 0 0 0 2px var(--primary-glow);
        }

        .btn-generate {
            width: 100%;
            padding: 1rem;
            font-size: 1.05rem;
            font-weight: 700;
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
        }

        .btn-generate:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 30px var(--primary-glow);
            filter: brightness(1.15);
        }

        .btn-generate:disabled { opacity: 0.65; cursor: not-allowed; transform: none; }

        /* Voice Explanation Audio Control Bar */
        .voice-control-bar {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            background: rgba(9, 13, 22, 0.85);
            border: 1px solid var(--border-glow);
            padding: 0.5rem 1rem;
            border-radius: 12px;
            margin-bottom: 0.5rem;
            flex-wrap: wrap;
        }

        .btn-voice {
            background: linear-gradient(135deg, var(--accent-purple), var(--accent-pink));
            color: white;
            border: none;
            padding: 0.45rem 1rem;
            border-radius: 8px;
            font-family: inherit;
            font-size: 0.88rem;
            font-weight: 700;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            transition: all 0.25s;
            box-shadow: 0 0 14px rgba(244, 114, 182, 0.3);
        }

        .btn-voice:hover {
            filter: brightness(1.15);
            transform: translateY(-1px);
        }

        .btn-voice-stop {
            background: rgba(239, 68, 68, 0.15);
            border: 1px solid rgba(239, 68, 68, 0.3);
            color: #f87171;
            padding: 0.45rem 0.7rem;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.2s;
        }

        .btn-voice-stop:hover { background: rgba(239, 68, 68, 0.3); }

        .voice-rate-wrap {
            display: flex;
            align-items: center;
            gap: 0.4rem;
            background: var(--editor-bg);
            padding: 0.25rem 0.6rem;
            border-radius: 8px;
            border: 1px solid var(--border-subtle);
        }

        .voice-rate-wrap select {
            background: transparent;
            border: none;
            color: var(--text-heading);
            font-size: 0.82rem;
            font-family: inherit;
            font-weight: 600;
            outline: none;
            cursor: pointer;
        }

        .voice-equalizer {
            display: flex;
            align-items: flex-end;
            gap: 3px;
            height: 18px;
            margin-left: auto;
        }

        .eq-bar {
            width: 3px;
            background: var(--accent-pink);
            border-radius: 2px;
            animation: eq-pulse 0.7s infinite alternate ease-in-out;
        }

        .eq-bar:nth-child(1) { animation-delay: 0.1s; height: 50%; }
        .eq-bar:nth-child(2) { animation-delay: 0.3s; height: 100%; }
        .eq-bar:nth-child(3) { animation-delay: 0.2s; height: 35%; }
        .eq-bar:nth-child(4) { animation-delay: 0.4s; height: 80%; }

        @keyframes eq-pulse {
            0% { height: 20%; }
            100% { height: 100%; }
        }

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

        .export-dock {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 0.75rem;
        }

        @media (max-width: 640px) { .export-dock { grid-template-columns: repeat(2, 1fr); } }

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

        .export-card.disabled { opacity: 0.35; pointer-events: none; }

        footer {
            text-align: center;
            padding: 1.75rem;
            color: var(--text-muted);
            font-size: 0.85rem;
            border-top: 1px solid var(--border-subtle);
            background: rgba(5, 8, 17, 0.95);
        }

        .system-status-ticker {
            display: flex;
            align-items: center;
            gap: 0.65rem;
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid var(--border-subtle);
            padding: 0.4rem 0.9rem;
            border-radius: 30px;
            font-size: 0.78rem;
            color: var(--text-body);
        }

        .system-status-ticker strong {
            color: var(--text-heading);
        }

        @media (max-width: 992px) {
            .system-status-ticker { display: none; }
        }

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
                    <h1>NeuroDocs <span style="font-weight: 300; opacity: 0.85; font-size: 0.9em; background: linear-gradient(135deg, var(--primary-cyan), var(--accent-pink)); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">AI</span></h1>
                    <p><span class="pulse-dot" style="background: var(--emerald-green); width: 6px; height: 6px;"></span> Engine v2.5 &bull; Quantum Edition</p>
                </div>
            </a>

            <div class="system-status-ticker">
                <i class="fa-solid fa-shield-halved" style="color: var(--emerald-green);"></i>
                <span>Status: <strong>Online</strong></span>
                <span style="opacity: 0.3;">|</span>
                <i class="fa-solid fa-bolt" style="color: var(--primary-cyan);"></i>
                <span>Engine: <strong>Free Visitor Access</strong></span>
            </div>

            <div class="header-actions">
                <button id="userAuthBtn" class="key-badge-btn missing" onclick="openAuthModal()">
                    <i class="fa-solid fa-user-gear"></i>
                    <span id="userAuthText">Sign In</span>
                </button>

                <button id="keyBadgeBtn" class="key-badge-btn missing" onclick="openKeyModal()">
                    <span class="pulse-dot"></span>
                    <span id="keyBadgeText">Checking Key...</span>
                    <i class="fa-solid fa-gear" style="font-size: 0.85rem; opacity: 0.7;"></i>
                </button>

                <div class="ctrl-pill" title="Selected Gemini Model">
                    <i class="fa-solid fa-microchip" style="color: var(--accent-purple);"></i>
                    <select id="modelSelect">
                        <option value="gemini-flash-latest">gemini-flash-latest</option>
                        <option value="gemini-2.0-flash">gemini-2.0-flash</option>
                        <option value="gemini-2.0-flash-lite">gemini-2.0-flash-lite</option>
                        <option value="gemini-pro-latest">gemini-pro-latest</option>
                    </select>
                </div>
            </div>
        </div>
    </header>

    <nav class="sub-nav">
        <div class="sub-nav-wrap">
            <button class="nav-mode-btn active" id="navStudio" onclick="switchViewMode('studio')">
                <i class="fa-solid fa-bolt"></i> AI Generator Studio
            </button>
            <button class="nav-mode-btn" id="navDashboard" onclick="switchViewMode('dashboard')">
                <i class="fa-solid fa-chart-pie"></i> Analytics Dashboard
            </button>
            <button class="nav-mode-btn" id="navHistory" onclick="switchViewMode('history')">
                <i class="fa-solid fa-folder-closed"></i> History Library <span id="historyBadgeCount" class="badge-count">0</span>
            </button>
        </div>
    </nav>

    <div id="keyModal" class="modal-overlay">
        <div class="modal-card">
            <div class="modal-header">
                <div class="modal-title"><i class="fa-solid fa-key" style="color: var(--primary-cyan);"></i> Gemini API Key Settings</div>
                <button class="modal-close-btn" onclick="closeKeyModal()"><i class="fa-solid fa-xmark"></i></button>
            </div>

            <p style="font-size: 0.9rem; color: var(--text-body);"><strong>No API key required!</strong> All visitors enjoy free automatic access powered by our server. Entering a personal key below is completely optional.</p>

            <div class="key-input-box">
                <i class="fa-solid fa-lock" style="color: var(--text-muted);"></i>
                <input type="password" id="userApiKey" placeholder="Optional: Paste custom key if desired...">
                <i id="eyeToggle" class="fa-solid fa-eye eye-toggle" onclick="toggleKeyVisibility()"></i>
            </div>

            <div id="validationAlert" style="display: none; font-size: 0.85rem; padding: 0.6rem 0.8rem; border-radius: 8px; margin-bottom: 1rem;"></div>

            <div class="info-callout">
                💡 <strong>Free Visitor Access Active:</strong><br>
                You do not need to enter an API key to generate documentation. If you want your own personal API key limits, generate one at <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener">Google AI Studio <i class="fa-solid fa-arrow-up-right-from-square" style="font-size: 0.75rem;"></i></a>.
            </div>

            <div class="modal-actions">
                <button class="btn btn-outline" style="flex: 1;" onclick="testKeyConnection()"><i class="fa-solid fa-vial"></i> Test Key</button>
                <button class="btn btn-accent" style="flex: 1;" onclick="saveKeyAndClose()"><i class="fa-solid fa-floppy-disk"></i> Save & Apply</button>
                <button class="btn btn-outline" onclick="clearVisitorKey()"><i class="fa-solid fa-trash-can"></i></button>
            </div>
        </div>
    </div>

    <!-- Auth Modal -->
    <div id="authModal" class="modal-overlay">
        <div class="modal-card" style="max-width: 480px;">
            <div class="modal-header">
                <div class="modal-title"><i class="fa-solid fa-user-shield" style="color: var(--primary-cyan);"></i> Account Authentication</div>
                <button class="modal-close-btn" onclick="closeAuthModal()"><i class="fa-solid fa-xmark"></i></button>
            </div>

            <!-- Logged In Profile View -->
            <div id="loggedInView" style="display: none; text-align: center; padding: 1rem 0;">
                <div style="width: 64px; height: 64px; border-radius: 50%; background: linear-gradient(135deg, var(--primary-cyan), var(--accent-purple)); display: flex; align-items: center; justify-content: center; margin: 0 auto 1rem; font-size: 1.8rem; color: white; box-shadow: 0 0 20px var(--primary-glow);">
                    <i class="fa-solid fa-user-check"></i>
                </div>
                <h3 id="profileName" style="color: var(--text-heading); font-size: 1.2rem; font-weight: 700;">Dani Toffin</h3>
                <p id="profileEmail" style="color: var(--text-muted); font-size: 0.88rem; margin-bottom: 1.2rem;">demo@neurodocs.ai</p>
                <div style="background: rgba(56, 189, 248, 0.08); border: 1px solid var(--border-glow); padding: 0.75rem; border-radius: 12px; font-size: 0.85rem; color: var(--primary-cyan); margin-bottom: 1.5rem;">
                    ⚡ <strong>Pro Developer Account Active</strong> &bull; Free Generation Enabled
                </div>
                <button class="btn btn-outline" style="width: 100%; border-color: rgba(239,68,68,0.4); color: #f87171; justify-content: center;" onclick="logoutUser()">
                    <i class="fa-solid fa-right-from-bracket"></i> Log Out Account
                </button>
            </div>

            <!-- Logged Out Auth View -->
            <div id="loggedOutView">
                <div class="auth-tab-bar">
                    <button class="auth-tab-btn active" id="authTabLogin" onclick="switchAuthTab('login')">
                        <i class="fa-solid fa-right-to-bracket"></i> Log In
                    </button>
                    <button class="auth-tab-btn" id="authTabSignup" onclick="switchAuthTab('signup')">
                        <i class="fa-solid fa-user-plus"></i> Create Account
                    </button>
                </div>

                <div id="authAlert" style="display: none; font-size: 0.85rem; padding: 0.65rem 0.9rem; border-radius: 10px; margin-bottom: 1rem;"></div>

                <!-- Login Container -->
                <div id="loginFormContainer">
                    <div class="key-input-box">
                        <i class="fa-solid fa-envelope" style="color: var(--text-muted);"></i>
                        <input type="email" id="loginEmail" placeholder="Email address (e.g. demo@neurodocs.ai)" onkeypress="if(event.key==='Enter') submitLogin()">
                    </div>
                    <div class="key-input-box">
                        <i class="fa-solid fa-lock" style="color: var(--text-muted);"></i>
                        <input type="password" id="loginPassword" placeholder="Password (e.g. password123)" onkeypress="if(event.key==='Enter') submitLogin()">
                        <i id="loginEyeToggle" class="fa-solid fa-eye eye-toggle" onclick="toggleLoginPasswordVisibility()"></i>
                    </div>
                    <button type="button" class="btn btn-accent" style="width: 100%; margin-top: 0.75rem; justify-content: center;" onclick="submitLogin()">
                        <i class="fa-solid fa-right-to-bracket"></i> Log In to NeuroDocs
                    </button>
                </div>

                <!-- Signup Container -->
                <div id="signupFormContainer" style="display: none;">
                    <div class="key-input-box">
                        <i class="fa-solid fa-user" style="color: var(--text-muted);"></i>
                        <input type="text" id="signupName" placeholder="Full Name" onkeypress="if(event.key==='Enter') submitSignup()">
                    </div>
                    <div class="key-input-box">
                        <i class="fa-solid fa-envelope" style="color: var(--text-muted);"></i>
                        <input type="email" id="signupEmail" placeholder="Email address" onkeypress="if(event.key==='Enter') submitSignup()">
                    </div>
                    <div class="key-input-box">
                        <i class="fa-solid fa-lock" style="color: var(--text-muted);"></i>
                        <input type="password" id="signupPassword" placeholder="Password (min 6 characters)" onkeypress="if(event.key==='Enter') submitSignup()">
                        <i id="signupEyeToggle" class="fa-solid fa-eye eye-toggle" onclick="toggleSignupPasswordVisibility()"></i>
                    </div>
                    <button type="button" class="btn btn-accent" style="width: 100%; margin-top: 0.75rem; justify-content: center;" onclick="submitSignup()">
                        <i class="fa-solid fa-user-plus"></i> Create Free Account
                    </button>
                </div>

                <div style="position: relative; margin: 1.25rem 0; text-align: center;">
                    <hr style="border: none; border-top: 1px solid var(--border-subtle);">
                    <span style="position: absolute; top: -10px; left: 50%; transform: translateX(-50%); background: #0f172a; padding: 0 10px; font-size: 0.75rem; color: var(--text-muted);">QUICK DEMO</span>
                </div>

                <button class="btn btn-outline" type="button" style="width: 100%; justify-content: center; border-color: rgba(56, 189, 248, 0.4); color: var(--primary-cyan);" onclick="quickDemoLogin()">
                    <i class="fa-solid fa-wand-magic-sparkles"></i> ⚡ 1-Click Quick Demo Login
                </button>
            </div>
        </div>
    </div>

    <main>
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

            <!-- Generator Column Configurator -->
            <div class="generator-config-box">
                <div class="config-title">
                    <i class="fa-solid fa-list-check" style="color: var(--primary-cyan);"></i>
                    Generator Column &mdash; Select Sections to Generate:
                </div>
                <div class="section-chips" id="sectionChipsContainer">
                    <label class="chip-item"><input type="checkbox" value="Project Overview & Purpose" checked> <span>Overview & Purpose</span></label>
                    <label class="chip-item"><input type="checkbox" value="Key Features & Capabilities" checked> <span>Features & Capabilities</span></label>
                    <label class="chip-item"><input type="checkbox" value="Architecture & Code Structure" checked> <span>Architecture & Structure</span></label>
                    <label class="chip-item"><input type="checkbox" value="Function & Method Descriptions" checked> <span>Functions & Methods</span></label>
                    <label class="chip-item"><input type="checkbox" value="Class & Data Models" checked> <span>Classes & Models</span></label>
                    <label class="chip-item"><input type="checkbox" value="Inputs, Outputs & API Schema" checked> <span>Inputs, Outputs & Schema</span></label>
                    <label class="chip-item"><input type="checkbox" value="Installation & Setup Guide" checked> <span>Installation Guide</span></label>
                    <label class="chip-item"><input type="checkbox" value="Code Examples & Usage" checked> <span>Usage Examples</span></label>
                </div>
                <div class="preset-row">
                    <span style="font-size: 0.75rem; color: var(--text-muted); font-weight: 600;">PRESETS:</span>
                    <button type="button" class="preset-btn" onclick="selectSectionPreset('all')"><i class="fa-solid fa-check-double"></i> Select All</button>
                    <button type="button" class="preset-btn" onclick="selectSectionPreset('quickstart')"><i class="fa-solid fa-bolt"></i> Quickstart Spec</button>
                    <button type="button" class="preset-btn" onclick="selectSectionPreset('api')"><i class="fa-solid fa-code"></i> API Reference</button>
                </div>
            </div>

            <button id="generateBtn" class="btn-generate" onclick="generateDocs()">
                <i class="fa-solid fa-bolt"></i> Generate AI Documentation
            </button>
        </section>

        <section class="workspace-panel">
            <div class="panel-top">
                <div class="panel-heading"><i class="fa-solid fa-file-invoice"></i> Documentation Hub</div>
                <div class="tab-group">
                    <button class="tab-btn active" id="tabRendered" onclick="switchTab('rendered')">Preview</button>
                    <button class="tab-btn" id="tabRaw" onclick="switchTab('raw')">Raw Markdown</button>
                </div>
            </div>

            <!-- Voice Explanation Control Bar -->
            <div id="voiceControlBar" class="voice-control-bar" style="display: none;">
                <button id="voicePlayBtn" class="btn-voice" onclick="toggleVoiceExplanation()">
                    <i class="fa-solid fa-volume-high"></i> <span id="voiceBtnText">Voice Explanation</span>
                </button>
                <button class="btn-voice-stop" title="Stop Audio" onclick="stopVoiceExplanation()"><i class="fa-solid fa-stop"></i></button>
                <div class="voice-rate-wrap">
                    <i class="fa-solid fa-gauge-high" style="font-size: 0.75rem; color: var(--accent-pink);"></i>
                    <select id="voiceRate" onchange="updateVoiceSettings()">
                        <option value="0.85">0.85x</option>
                        <option value="1.0" selected>1.0x</option>
                        <option value="1.25">1.25x</option>
                        <option value="1.5">1.5x</option>
                    </select>
                </div>
                <div id="voiceEqualizer" class="voice-equalizer" style="display: none;">
                    <span class="eq-bar"></span><span class="eq-bar"></span><span class="eq-bar"></span><span class="eq-bar"></span>
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
    </main>

    <!-- Analytics Dashboard View -->
    <section id="dashboardView" class="dashboard-panel" style="display: none;">
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-icon" style="background: rgba(56, 189, 248, 0.12); color: var(--primary-cyan);"><i class="fa-solid fa-file-invoice"></i></div>
                <div class="stat-info">
                    <small>TOTAL DOCS GENERATED</small>
                    <h2 id="statTotalDocs">1,284</h2>
                    <span class="stat-trend" style="color: var(--emerald-green);"><i class="fa-solid fa-arrow-trend-up"></i> +18.4% this week</span>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background: rgba(129, 140, 248, 0.12); color: var(--accent-purple);"><i class="fa-solid fa-code"></i></div>
                <div class="stat-info">
                    <small>CODE LINES ANALYZED</small>
                    <h2 id="statCodeLines">84,520</h2>
                    <span class="stat-trend" style="color: var(--primary-cyan);"><i class="fa-solid fa-layer-group"></i> 12 Languages</span>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background: rgba(52, 211, 153, 0.12); color: var(--emerald-green);"><i class="fa-solid fa-microchip"></i></div>
                <div class="stat-info">
                    <small>AI ENGINE STATUS</small>
                    <h2>Gemini 2.0</h2>
                    <span class="stat-trend" style="color: var(--emerald-green);"><i class="fa-solid fa-circle-check"></i> 100% Operational</span>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-icon" style="background: rgba(244, 114, 182, 0.12); color: var(--accent-pink);"><i class="fa-solid fa-stopwatch"></i></div>
                <div class="stat-info">
                    <small>AVG GENERATION SPEED</small>
                    <h2>1.18s</h2>
                    <span class="stat-trend" style="color: var(--amber-gold);"><i class="fa-solid fa-bolt"></i> Ultra Fast</span>
                </div>
            </div>
        </div>

        <div class="dashboard-grid">
            <div class="dash-card">
                <div class="dash-card-header"><i class="fa-solid fa-chart-simple" style="color: var(--primary-cyan);"></i> Source Code Language Distribution</div>
                <div class="lang-bar-container">
                    <div class="lang-bar-segment" style="width: 45%; background: var(--primary-cyan);" title="Python 45%"></div>
                    <div class="lang-bar-segment" style="width: 28%; background: var(--accent-purple);" title="JavaScript 28%"></div>
                    <div class="lang-bar-segment" style="width: 15%; background: var(--emerald-green);" title="C++ / C 15%"></div>
                    <div class="lang-bar-segment" style="width: 12%; background: var(--accent-pink);" title="Java / Other 12%"></div>
                </div>
                <div class="lang-legend">
                    <span><i class="fa-solid fa-circle" style="color: var(--primary-cyan);"></i> Python (45%)</span>
                    <span><i class="fa-solid fa-circle" style="color: var(--accent-purple);"></i> JavaScript (28%)</span>
                    <span><i class="fa-solid fa-circle" style="color: var(--emerald-green);"></i> C++ / C (15%)</span>
                    <span><i class="fa-solid fa-circle" style="color: var(--accent-pink);"></i> Java / Other (12%)</span>
                </div>
            </div>

            <div class="dash-card">
                <div class="dash-card-header"><i class="fa-solid fa-server" style="color: var(--accent-purple);"></i> Telemetry & Capacity</div>
                <p style="font-size: 0.85rem; color: var(--text-body); margin-bottom: 0.85rem;">All visitor requests are balanced across active Gemini API server pools.</p>
                <div style="display: flex; flex-direction: column; gap: 0.5rem; font-size: 0.8rem;">
                    <div style="display: flex; justify-content: space-between;"><span>Server Load:</span><strong style="color: var(--emerald-green);">22% (Optimal)</strong></div>
                    <div style="display: flex; justify-content: space-between;"><span>Free Visitor Quota:</span><strong style="color: var(--primary-cyan);">Unlimited</strong></div>
                    <div style="display: flex; justify-content: space-between;"><span>Active Export Engines:</span><strong style="color: var(--accent-pink);">PDF, DOCX, MD, ZIP</strong></div>
                </div>
            </div>
        </div>
    </section>

    <!-- History Library View -->
    <section id="historyView" class="dashboard-panel" style="display: none;">
        <div class="dash-card">
            <div class="dash-card-header" style="display: flex; justify-content: space-between; align-items: center;">
                <div><i class="fa-solid fa-clock-rotate-left" style="color: var(--primary-cyan);"></i> Generated Documentation Library</div>
                <button class="btn btn-outline" style="font-size: 0.78rem; padding: 0.35rem 0.75rem;" onclick="clearHistoryLibrary()"><i class="fa-solid fa-trash"></i> Clear History</button>
            </div>

            <div id="historyListContainer">
                <!-- History Items Rendered Here -->
            </div>
        </div>
    </section>

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

        const authModal = document.getElementById('authModal');
        const userAuthBtn = document.getElementById('userAuthBtn');
        const userAuthText = document.getElementById('userAuthText');
        const loggedInView = document.getElementById('loggedInView');
        const loggedOutView = document.getElementById('loggedOutView');
        const authAlert = document.getElementById('authAlert');
        const profileName = document.getElementById('profileName');
        const profileEmail = document.getElementById('profileEmail');

        const codeEditor = document.getElementById('codeEditor');
        const fileInput = document.getElementById('fileInput');
        const dropArea = document.getElementById('dropArea');
        const generateBtn = document.getElementById('generateBtn');
        const outputRendered = document.getElementById('outputRendered');
        const outputRaw = document.getElementById('outputRaw');
        const voiceControlBar = document.getElementById('voiceControlBar');
        let currentToken = null;
        let isSpeaking = false;
        let currentUtterance = null;

        const savedKey = localStorage.getItem('neurodocs_gemini_key');
        if (savedKey) apiKeyInput.value = savedKey;

        function openKeyModal() { keyModal.classList.add('open'); }
        function closeKeyModal() { keyModal.classList.remove('open'); }

        function openAuthModal() { authModal.classList.add('open'); }
        function closeAuthModal() { authModal.classList.remove('open'); }

        function switchAuthTab(tab) {
            authAlert.style.display = 'none';
            if (tab === 'login') {
                document.getElementById('authTabLogin').classList.add('active');
                document.getElementById('authTabSignup').classList.remove('active');
                document.getElementById('loginFormContainer').style.display = 'block';
                document.getElementById('signupFormContainer').style.display = 'none';
            } else {
                document.getElementById('authTabSignup').classList.add('active');
                document.getElementById('authTabLogin').classList.remove('active');
                document.getElementById('signupFormContainer').style.display = 'block';
                document.getElementById('loginFormContainer').style.display = 'none';
            }
        }

        function toggleLoginPasswordVisibility() {
            const pwdInput = document.getElementById('loginPassword');
            const toggleIcon = document.getElementById('loginEyeToggle');
            if (pwdInput.type === 'password') {
                pwdInput.type = 'text';
                toggleIcon.className = 'fa-solid fa-eye-slash eye-toggle';
            } else {
                pwdInput.type = 'password';
                toggleIcon.className = 'fa-solid fa-eye eye-toggle';
            }
        }

        function toggleSignupPasswordVisibility() {
            const pwdInput = document.getElementById('signupPassword');
            const toggleIcon = document.getElementById('signupEyeToggle');
            if (pwdInput.type === 'password') {
                pwdInput.type = 'text';
                toggleIcon.className = 'fa-solid fa-eye-slash eye-toggle';
            } else {
                pwdInput.type = 'password';
                toggleIcon.className = 'fa-solid fa-eye eye-toggle';
            }
        }

        async function checkAuthSession() {
            const token = localStorage.getItem('neurodocs_auth_token');
            if (!token) {
                renderLoggedOutState();
                return;
            }
            try {
                const res = await fetch('/api/auth/me', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ token: token })
                });
                const data = await res.json();
                if (data.authenticated && data.user) {
                    renderLoggedInState(data.user);
                } else {
                    renderLoggedOutState();
                }
            } catch (err) {
                renderLoggedOutState();
            }
        }

        function renderLoggedInState(user) {
            userAuthText.textContent = user.name;
            userAuthBtn.className = 'key-badge-btn custom';
            profileName.textContent = user.name;
            profileEmail.textContent = user.email;
            loggedInView.style.display = 'block';
            loggedOutView.style.display = 'none';
        }

        function renderLoggedOutState() {
            userAuthText.textContent = 'Sign In';
            userAuthBtn.className = 'key-badge-btn missing';
            loggedInView.style.display = 'none';
            loggedOutView.style.display = 'block';
        }

        async function submitLogin() {
            const email = document.getElementById('loginEmail').value.trim();
            const password = document.getElementById('loginPassword').value;

            if (!email || !password) {
                authAlert.style.display = 'block';
                authAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                authAlert.style.color = '#f87171';
                authAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> Please enter both email and password.`;
                return;
            }

            authAlert.style.display = 'block';
            authAlert.style.background = 'rgba(56, 189, 248, 0.15)';
            authAlert.style.color = '#38bdf8';
            authAlert.innerHTML = `<i class="fa-solid fa-spinner spin"></i> Authenticating...`;

            try {
                const res = await fetch('/api/auth/login', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ email: email, password: password })
                });
                const data = await res.json();
                if (data.success) {
                    localStorage.setItem('neurodocs_auth_token', data.token);
                    authAlert.style.background = 'rgba(52, 211, 153, 0.15)';
                    authAlert.style.color = '#34d399';
                    authAlert.innerHTML = `<i class="fa-solid fa-circle-check"></i> ${data.message}`;
                    setTimeout(() => {
                        renderLoggedInState(data.user);
                        closeAuthModal();
                    }, 500);
                } else {
                    authAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                    authAlert.style.color = '#f87171';
                    authAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> ${data.message}`;
                }
            } catch (err) {
                authAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                authAlert.style.color = '#f87171';
                authAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> Login request failed.`;
            }
        }

        async function submitSignup() {
            const name = document.getElementById('signupName').value.trim();
            const email = document.getElementById('signupEmail').value.trim();
            const password = document.getElementById('signupPassword').value;

            if (!email || !password) {
                authAlert.style.display = 'block';
                authAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                authAlert.style.color = '#f87171';
                authAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> Please fill in all required fields.`;
                return;
            }

            authAlert.style.display = 'block';
            authAlert.style.background = 'rgba(56, 189, 248, 0.15)';
            authAlert.style.color = '#38bdf8';
            authAlert.innerHTML = `<i class="fa-solid fa-spinner spin"></i> Creating account...`;

            try {
                const res = await fetch('/api/auth/signup', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ name: name, email: email, password: password })
                });
                const data = await res.json();
                if (data.success) {
                    localStorage.setItem('neurodocs_auth_token', data.token);
                    authAlert.style.background = 'rgba(52, 211, 153, 0.15)';
                    authAlert.style.color = '#34d399';
                    authAlert.innerHTML = `<i class="fa-solid fa-circle-check"></i> ${data.message}`;
                    setTimeout(() => {
                        renderLoggedInState(data.user);
                        closeAuthModal();
                    }, 500);
                } else {
                    authAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                    authAlert.style.color = '#f87171';
                    authAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> ${data.message}`;
                }
            } catch (err) {
                authAlert.style.background = 'rgba(239, 68, 68, 0.15)';
                authAlert.style.color = '#f87171';
                authAlert.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> Signup request failed.`;
            }
        }

        async function quickDemoLogin() {
            switchAuthTab('login');
            document.getElementById('loginEmail').value = 'demo@neurodocs.ai';
            document.getElementById('loginPassword').value = 'password123';
            submitLogin();
        }

        async function logoutUser() {
            const token = localStorage.getItem('neurodocs_auth_token');
            if (token) {
                fetch('/api/auth/logout', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ token: token })
                });
            }
            localStorage.removeItem('neurodocs_auth_token');
            renderLoggedOutState();
            closeAuthModal();
        }

        checkAuthSession();
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
                        keyBadgeText.textContent = 'Custom Key Active';
                    } else {
                        keyBadgeBtn.className = 'key-badge-btn server';
                        keyBadgeText.textContent = 'Free Access Active';
                    }
                } else {
                    keyBadgeBtn.className = 'key-badge-btn server';
                    keyBadgeText.textContent = 'Free Access Active';
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

        // Voice Explanation Logic
        function cleanTextForSpeech(text) {
            return text
                .replace(/```[\\s\\S]*?```/g, ' Code snippet omitted. ')
                .replace(/`([^`]+)`/g, '$1')
                .replace(/#{1,6}\\s+/g, '')
                .replace(/[*_~]/g, '')
                .replace(/\\[([^\\]]+)\\]\\([^)]+\\)/g, '$1')
                .replace(/\\n+/g, '. ');
        }

        function toggleVoiceExplanation() {
            if (!('speechSynthesis' in window)) {
                alert('Voice synthesis is not supported in your browser.');
                return;
            }

            if (window.speechSynthesis.speaking && !window.speechSynthesis.paused) {
                window.speechSynthesis.pause();
                isSpeaking = false;
                document.getElementById('voiceBtnText').textContent = 'Resume Voice';
                document.getElementById('voiceEqualizer').style.display = 'none';
                return;
            }

            if (window.speechSynthesis.paused) {
                window.speechSynthesis.resume();
                isSpeaking = true;
                document.getElementById('voiceBtnText').textContent = 'Pause Voice';
                document.getElementById('voiceEqualizer').style.display = 'flex';
                return;
            }

            const docText = outputRaw.value;
            if (!docText) return;

            const speechText = cleanTextForSpeech(docText);
            currentUtterance = new SpeechSynthesisUtterance(speechText);
            currentUtterance.rate = parseFloat(document.getElementById('voiceRate').value || 1.0);

            const voices = window.speechSynthesis.getVoices();
            const englishVoice = voices.find(v => v.lang.startsWith('en') && (v.name.includes('Google') || v.name.includes('Natural') || v.name.includes('Online')));
            if (englishVoice) currentUtterance.voice = englishVoice;

            currentUtterance.onstart = () => {
                isSpeaking = true;
                document.getElementById('voiceBtnText').textContent = 'Pause Voice';
                document.getElementById('voiceEqualizer').style.display = 'flex';
            };

            currentUtterance.onend = () => {
                isSpeaking = false;
                document.getElementById('voiceBtnText').textContent = 'Voice Explanation';
                document.getElementById('voiceEqualizer').style.display = 'none';
            };

            currentUtterance.onerror = () => {
                isSpeaking = false;
                document.getElementById('voiceBtnText').textContent = 'Voice Explanation';
                document.getElementById('voiceEqualizer').style.display = 'none';
            };

            window.speechSynthesis.cancel();
            window.speechSynthesis.speak(currentUtterance);
        }

        function stopVoiceExplanation() {
            if ('speechSynthesis' in window) {
                window.speechSynthesis.cancel();
                isSpeaking = false;
                document.getElementById('voiceBtnText').textContent = 'Voice Explanation';
                document.getElementById('voiceEqualizer').style.display = 'none';
            }
        }

        function updateVoiceSettings() {
            if (currentUtterance && window.speechSynthesis.speaking) {
                stopVoiceExplanation();
                toggleVoiceExplanation();
            }
        }

        function selectSectionPreset(preset) {
            const checkboxes = document.querySelectorAll('#sectionChipsContainer input[type="checkbox"]');
            checkboxes.forEach(cb => {
                if (preset === 'all') {
                    cb.checked = true;
                } else if (preset === 'quickstart') {
                    cb.checked = ['Project Overview & Purpose', 'Installation & Setup Guide', 'Code Examples & Usage'].includes(cb.value);
                } else if (preset === 'api') {
                    cb.checked = ['Function & Method Descriptions', 'Class & Data Models', 'Inputs, Outputs & API Schema'].includes(cb.value);
                }
            });
        }

        authModal.addEventListener('click', (e) => { if (e.target === authModal) closeAuthModal(); });
        keyModal.addEventListener('click', (e) => { if (e.target === keyModal) closeKeyModal(); });

        async function generateDocs() {
            const code = codeEditor.value.trim();
            if (!code) {
                alert('Please enter or upload source code first.');
                return;
            }

            const selectedSections = Array.from(document.querySelectorAll('#sectionChipsContainer input[type="checkbox"]:checked')).map(cb => cb.value);

            stopVoiceExplanation();
            generateBtn.disabled = true;
            generateBtn.innerHTML = `<i class="fa-solid fa-spinner spin"></i> Analyzing & Generating AI Documentation...`;

            try {
                const res = await fetch('/api/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        code: code,
                        model: modelSelect.value,
                        api_key: apiKeyInput.value.trim(),
                        sections: selectedSections
                    })
                });

                const data = await res.json();
                if (!res.ok) throw new Error(data.error || 'Generation failed');

                const docText = data.documentation;
                currentToken = data.token;

                outputRendered.innerHTML = marked.parse(docText);
                outputRaw.value = docText;
                voiceControlBar.style.display = 'flex';

                ['Docx', 'Pdf', 'Md', 'Zip'].forEach(fmt => {
                    const card = document.getElementById('btn' + fmt);
                    card.href = `/api/download/${fmt.toLowerCase()}?token=${currentToken}`;
                    card.classList.remove('disabled');
                });

                // Register into History Library
                const titleLine = docText.split('\n').find(l => l.trim().startsWith('#')) || '# Software Documentation';
                addToHistory({
                    id: currentToken,
                    token: currentToken,
                    title: titleLine.replace(/^#+\\s*/, '').trim() || 'Software Documentation',
                    model: modelSelect.value,
                    sectionsCount: selectedSections.length,
                    timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
                });

            } catch (err) {
                alert('Error: ' + err.message);
            } finally {
                generateBtn.disabled = false;
                generateBtn.innerHTML = `<i class="fa-solid fa-bolt"></i> Generate AI Documentation`;
            }
        }

        function switchViewMode(mode) {
            const studioView = document.querySelector('main');
            const dashboardView = document.getElementById('dashboardView');
            const historyView = document.getElementById('historyView');

            document.getElementById('navStudio').className = 'nav-mode-btn' + (mode === 'studio' ? ' active' : '');
            document.getElementById('navDashboard').className = 'nav-mode-btn' + (mode === 'dashboard' ? ' active' : '');
            document.getElementById('navHistory').className = 'nav-mode-btn' + (mode === 'history' ? ' active' : '');

            if (mode === 'studio') {
                studioView.style.display = 'flex';
                dashboardView.style.display = 'none';
                historyView.style.display = 'none';
            } else if (mode === 'dashboard') {
                studioView.style.display = 'none';
                dashboardView.style.display = 'block';
                historyView.style.display = 'none';
            } else if (mode === 'history') {
                studioView.style.display = 'none';
                dashboardView.style.display = 'none';
                historyView.style.display = 'block';
                renderHistoryLibrary();
            }
        }

        function getHistoryList() {
            try {
                return JSON.parse(localStorage.getItem('neurodocs_history_list')) || [];
            } catch (e) { return []; }
        }

        function saveHistoryList(list) {
            localStorage.setItem('neurodocs_history_list', JSON.stringify(list));
            const badge = document.getElementById('historyBadgeCount');
            if (badge) badge.textContent = list.length;
        }

        function addToHistory(item) {
            const list = getHistoryList();
            list.unshift(item);
            saveHistoryList(list.slice(0, 30));
        }

        function clearHistoryLibrary() {
            if (confirm('Clear all documentation history?')) {
                saveHistoryList([]);
                renderHistoryLibrary();
            }
        }

        function renderHistoryLibrary() {
            const container = document.getElementById('historyListContainer');
            if (!container) return;
            const list = getHistoryList();
            const badge = document.getElementById('historyBadgeCount');
            if (badge) badge.textContent = list.length;

            if (list.length === 0) {
                container.innerHTML = `
                    <div style="text-align: center; padding: 3rem 1rem; color: var(--text-muted);">
                        <i class="fa-solid fa-folder-open" style="font-size: 2.5rem; margin-bottom: 0.75rem; color: var(--primary-cyan); opacity: 0.7;"></i>
                        <p style="font-weight: 600; font-size: 1rem; color: var(--text-heading);">No documentation generated yet.</p>
                        <p style="font-size: 0.82rem; margin-top: 0.3rem;">Switch to AI Generator Studio to create your first documentation.</p>
                    </div>`;
                return;
            }

            container.innerHTML = list.map(item => `
                <div class="history-item">
                    <div>
                        <h4 style="color: var(--text-heading); font-size: 0.95rem; font-weight: 700;"><i class="fa-solid fa-file-code" style="color: var(--primary-cyan);"></i> ${escapeHtml(item.title)}</h4>
                        <div style="font-size: 0.78rem; color: var(--text-muted); margin-top: 0.3rem; display: flex; gap: 0.85rem; flex-wrap: wrap;">
                            <span><i class="fa-solid fa-microchip"></i> ${escapeHtml(item.model)}</span>
                            <span><i class="fa-solid fa-clock"></i> ${escapeHtml(item.timestamp)}</span>
                            <span><i class="fa-solid fa-layer-group"></i> ${item.sectionsCount} Sections</span>
                        </div>
                    </div>
                    <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
                        <a href="/api/download/docx?token=${item.token}" class="btn btn-outline" style="font-size: 0.75rem; padding: 0.3rem 0.65rem;"><i class="fa-solid fa-file-word" style="color: #3b82f6;"></i> DOCX</a>
                        <a href="/api/download/pdf?token=${item.token}" class="btn btn-outline" style="font-size: 0.75rem; padding: 0.3rem 0.65rem;"><i class="fa-solid fa-file-pdf" style="color: #ef4444;"></i> PDF</a>
                        <a href="/api/download/md?token=${item.token}" class="btn btn-outline" style="font-size: 0.75rem; padding: 0.3rem 0.65rem;"><i class="fa-solid fa-file-code" style="color: #10b981;"></i> MD</a>
                        <a href="/api/download/zip?token=${item.token}" class="btn btn-outline" style="font-size: 0.75rem; padding: 0.3rem 0.65rem;"><i class="fa-solid fa-file-zipper" style="color: #f59e0b;"></i> ZIP</a>
                    </div>
                </div>
            `).join('');
        }

        function escapeHtml(str) {
            return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        }

        // Initialize History Badge Count
        const initialHistoryList = getHistoryList();
        const initialBadge = document.getElementById('historyBadgeCount');
        if (initialBadge) initialBadge.textContent = initialHistoryList.length;
    </script>
</body>
</html>"""
    return render_template_string(html_template)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
