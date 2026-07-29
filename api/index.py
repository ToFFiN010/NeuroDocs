from http.server import BaseHTTPRequestHandler

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.end_headers()
        html_content = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI Documentation Generator</title>
    <style>
        body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
            background: #0f172a;
            color: #f8fafc;
        }
        .card {
            background: #1e293b;
            padding: 2.5rem;
            border-radius: 16px;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.5), 0 8px 10px -6px rgba(0, 0, 0, 0.5);
            text-align: center;
            max-width: 540px;
            border: 1px solid #334155;
            margin: 1rem;
        }
        h1 { margin-top: 0; color: #38bdf8; font-size: 1.8rem; }
        p { line-height: 1.6; color: #94a3b8; margin-bottom: 1.2rem; }
        .badge {
            background: #0284c7;
            color: #f0f9ff;
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 0.85rem;
            font-weight: 600;
            display: inline-block;
            margin-bottom: 1rem;
        }
        .info-box {
            background: #0f172a;
            border-left: 4px solid #38bdf8;
            padding: 1rem;
            text-align: left;
            border-radius: 0 8px 8px 0;
            margin: 1.5rem 0;
            font-size: 0.95rem;
            color: #cbd5e1;
        }
        a.btn {
            display: inline-block;
            background: #2563eb;
            color: white;
            padding: 0.75rem 1.5rem;
            border-radius: 8px;
            text-decoration: none;
            font-weight: 600;
            transition: background 0.2s;
        }
        a.btn:hover { background: #1d4ed8; }
    </style>
</head>
<body>
    <div class="card">
        <h1>📄 AI Documentation Generator</h1>
        <div><span class="badge">Vercel Serverless Status: Active</span></div>
        <p>This project is built with <strong>Streamlit</strong>, which requires a continuous WebSocket server.</p>
        <div class="info-box">
            👉 For the full interactive UI, deploy this GitHub repository to <strong>Streamlit Community Cloud</strong> (share.streamlit.io) or run locally using <code>streamlit run app.py</code>.
        </div>
        <a class="btn" href="https://share.streamlit.io" target="_blank" rel="noopener">Go to Streamlit Community Cloud</a>
    </div>
</body>
</html>"""
        self.wfile.write(html_content.encode('utf-8'))
