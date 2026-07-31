import os
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

key = os.getenv("GEMINI_API_KEY", "").strip()
if not key:
    print("Warning: GEMINI_API_KEY not set in environment.")

genai.configure(api_key=key)

models_to_try = ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-flash-latest", "gemini-pro-latest"]
code = "print('Hello World')"

prompt = f"""
Analyze the following source code and generate professional documentation.

Include:
1. Project Overview
2. Features
3. Function Descriptions
4. Class Descriptions
5. Installation Guide
6. Usage Instructions

Source Code:

{code}
"""

documentation = None
for m in models_to_try:
    try:
        model = genai.GenerativeModel(m)
        res = model.generate_content(prompt)
        if res.text:
            documentation = res.text
            print(f"--- Generated using {m} ---")
            break
    except Exception as e:
        continue

if documentation:
    print(documentation)
else:
    print("Could not generate documentation with available models.")
