import os
import google.generativeai as genai

genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
model = genai.GenerativeModel("gemini-2.5-flash")

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
response = model.generate_content(prompt)
print(response.text)

