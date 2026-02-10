from flask import Flask, request, jsonify, send_from_directory

from flask_cors import CORS

import os

from openai import OpenAI

import requests

import json

from datetime import datetime

import PyPDF2

from docx import Document

import re

import io



def markdown_to_html(text: str) -> str:

    if not text:

        return ""



    text = re.sub(r"\*\*(.*?)\*\*", r"<strong>\1</strong>", text)

    text = re.sub(r"\*(.*?)\*", r"<em>\1</em>", text)

    text = re.sub(r"`([^`]*)`", r"\1", text)

    text = re.sub(r"^#{1,6}\s*(.*)$", r"<h3>\1</h3>", text, flags=re.MULTILINE)

    text = re.sub(r"\n{2,}", "<br><br>", text)



    return text





from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime

from sqlalchemy.orm import declarative_base, sessionmaker

from datetime import datetime as dt

import os

os.makedirs("generated", exist_ok=True)

from openpyxl import Workbook



def generate_excel(text: str, filename: str):

    wb = Workbook()

    ws = wb.active

    ws.title = "Report"



    def split_row(line: str):

        line = (line or "").strip()

        if not line:

            return [""]



        # Skip markdown-like separators

        if re.match(r"^[\s\-\|]+$", line):

            return []



        if "\t" in line:

            return [p.strip() for p in line.split("\t")]



        if "|" in line:

            parts = [p.strip() for p in re.split(r"\s*\|\s*", line) if p.strip()]

            if parts:

                return parts



        if " | " in line:

            return [p.strip() for p in line.split(" | ")]



        if "; " in line:

            parts = [p.strip() for p in line.split("; ")]

            if len(parts) >= 2:

                return parts



        if ": " in line:

            left, right = line.split(": ", 1)

            if len(left) <= 40:

                return [left.strip(), right.strip()]



        line = re.sub(r"^[-•]\s*", "", line)

        return [line]



    row_idx = 1

    for line in text.split("\n"):

        parts = split_row(line)

        if not parts:

            continue

        for j, part in enumerate(parts, start=1):

            ws.cell(row=row_idx, column=j, value=part)

        row_idx += 1



    path = f"generated/{filename}.xlsx"

    wb.save(path)

    return path

from docx import Document



def generate_word(text: str, filename: str):

    doc = Document()

    for line in text.split("\n"):

        doc.add_paragraph(line)



    path = f"generated/{filename}.docx"

    doc.save(path)

    return path

from reportlab.lib.pagesizes import A4

from reportlab.pdfgen import canvas



def generate_pdf(text: str, filename: str):

    path = f"generated/{filename}.pdf"

    c = canvas.Canvas(path, pagesize=A4)



    y = 800

    for line in text.split("\n"):

        c.drawString(40, y, line[:100])

        y -= 14

        if y < 40:

            c.showPage()

            y = 800



    c.save()

    return path



def strip_html(text: str) -> str:

    if not text:

        return ""

    text = re.sub(r"<br\s*\/?>", "\n", text)

    text = re.sub(r"<[^>]*>", "", text)

    return text.strip()



def clean_document_content(text: str) -> str:

    """Remove boilerplate lines like 'Here is ... in PDF format' and 'Download ...'."""

    if not text:

        return ""

    lines = [l.strip() for l in text.splitlines()]

    cleaned = []

    for line in lines:

        if not line:

            cleaned.append(line)

            continue

        l = line.lower()

        if l.startswith("download ") or "download the pdf" in l or "download the docx" in l or "download the word" in l:

            continue

        if l.startswith("here is") and ("format" in l or "document" in l or "pdf" in l or "docx" in l or "word" in l):

            continue

        if l.startswith("here is the") and ("format" in l or "document" in l or "pdf" in l or "docx" in l or "word" in l):

            continue

        cleaned.append(line)

    # Collapse multiple empty lines

    out = "\n".join(cleaned)

    out = re.sub(r"\n{3,}", "\n\n", out).strip()

    return out



def detect_file_type(message: str) -> str:

    m = (message or "").lower()

    if "xlsx" in m or "xls" in m or "excel" in m or "эксель" in m or "ексель" in m:

        return "excel"

    if "docx" in m or "doc" in m or "word" in m or "ворд" in m or "док" in m:

        return "word"

    if "pdf" in m or "пдф" in m:

        return "pdf"

    return "pdf"



DATABASE_URL = os.getenv(

    "DATABASE_URL",

    "postgresql+psycopg2://postgres:alm@localhost:5432/stratex"

)



engine = create_engine(DATABASE_URL, echo=False, future=True)

SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)

Base = declarative_base()



class ChatMessage(Base):

    __tablename__ = "chat_messages"



    id = Column(Integer, primary_key=True)

    user_id = Column(String, index=True)

    role = Column(String)

    content = Column(Text)

    created_at = Column(DateTime, default=dt.utcnow)



Base.metadata.create_all(engine)



def db_add_message(user_id: str, role: str, content: str):

    db = SessionLocal()

    try:

        db.add(ChatMessage(

            user_id=user_id,

            role=role,

            content=content

        ))

        db.commit()

    finally:

        db.close()



def db_get_history(user_id: str, limit: int = 20):

    db = SessionLocal()

    try:

        rows = (

            db.query(ChatMessage)

            .filter(ChatMessage.user_id == user_id)

            .order_by(ChatMessage.created_at.asc())

            .all()

        )

        rows = rows[-limit:]

        return [{"role": r.role, "content": r.content} for r in rows]

    finally:

        db.close()

    



app = Flask(__name__)

CORS(app)

@app.route("/api/generate-file", methods=["POST"])

def generate_file():

    data = request.get_json(silent=True) or {}

    content = data.get("content", "")

    file_type = data.get("file_type")  # excel | word | pdf



    if (not content) or (file_type not in ["excel", "word", "pdf"]):

        return jsonify({"success": False, "error": "Invalid request"}), 400



    name = f"report_{int(datetime.now().timestamp())}"



    if file_type == "excel":

        path = generate_excel(content, name)

    elif file_type == "word":

        path = generate_word(content, name)

    else:

        path = generate_pdf(content, name)



    return jsonify({

        "success": True,

        "download_url": f"/api/download/{os.path.basename(path)}"

    })





@app.route("/api/download/<filename>", methods=["GET"])

def download_file(filename):

    return send_from_directory("generated", filename, as_attachment=True)



# API keys (set via environment variables)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")

TAVILY_API_KEY = "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"



# OpenAI initialization

openai_client = OpenAI(api_key=OPENAI_API_KEY)



# Simple users database (for MVP)

USERS = {

    "demo": "demo123",

    "business": "business2024"

}



# Chat history for each user

chat_histories = {}



# Web search via Tavily

def web_search(query):

    """Search for information on the internet via Tavily API"""

    if not TAVILY_API_KEY:

        return {"answer": "", "sources": [], "error": "TAVILY_API_KEY is not set"}

    try:

        url = "https://api.tavily.com/search"

        payload = {

            "api_key": TAVILY_API_KEY,

            "query": query,

            "search_depth": "advanced",

            "include_answer": True,

            "max_results": 5

        }

        

        response = requests.post(url, json=payload)

        data = response.json()

        

        # Build results

        results = {

            "answer": data.get("answer", ""),

            "sources": []

        }

        

        for result in data.get("results", []):

            results["sources"].append({

                "title": result.get("title", ""),

                "url": result.get("url", ""),

                "content": result.get("content", "")[:300]  # first 300 characters

            })

        

        return results

    except Exception as e:

        print(f"Search error: {e}")

        return {"answer": "", "sources": []}



# Read PDF

def extract_pdf_text(file_bytes):

    """Extract text from a PDF"""

    try:

        pdf_file = io.BytesIO(file_bytes)

        pdf_reader = PyPDF2.PdfReader(pdf_file)

        text = ""

        for page in pdf_reader.pages:

            text += page.extract_text()

        return text

    except Exception as e:

        return f"PDF reading error: {str(e)}"



# Read DOCX

def extract_docx_text(file_bytes):

    """Extract text from a DOCX"""

    try:

        doc_file = io.BytesIO(file_bytes)

        doc = Document(doc_file)

        text = "\n".join([paragraph.text for paragraph in doc.paragraphs])

        return text

    except Exception as e:

        return f"DOCX reading error: {str(e)}"



# Chat with AI

def chat_with_ai(message, system_prompt, conversation_history=None, search_context=None):

    """Chat with OpenAI"""

    try:

        if not OPENAI_API_KEY:

            return "OpenAI API key is not configured. Please set OPENAI_API_KEY."

        messages = []

        

        # System prompt

        full_system_prompt = (

            system_prompt + "\n\n"

            f"Current date: {datetime.now().strftime('%d.%m.%Y')}\n\n"

            "Formatting rules:\n"

            "- DO NOT use Markdown (** ## `).\n"

            "- Use ONLY simple HTML tags.\n"

            "- Allowed tags: <strong>, <em>, <u>, <br>, <p>, <ul>, <li>, <h3>\n"

            "- Do NOT wrap everything in tags.\n"

        )





        

        if search_context:

            full_system_prompt += f"\n\nContext from the web:\n{search_context}"

        

        messages.append({"role": "system", "content": full_system_prompt})

        

        # Conversation history

        if conversation_history:

            messages.extend(conversation_history[-10:])  # last 10 messages

        

        # Current message

        messages.append({"role": "user", "content": message})

        

        # Request to Groq

        response = openai_client.chat.completions.create(

            model=OPENAI_MODEL,

            messages=messages,

            temperature=0.7,

            max_tokens=2048,

            top_p=0.9,

        )

        

        return response.choices[0].message.content

    

    except Exception as e:

        print(f"AI Error: {e}")

        return f"An error occurred while processing the request: {str(e)}"



# Detect user intent

def detect_intent(message):

    """Detect what the user wants"""

    message_lower = (message or "").lower()

    

    # Does it need web search?

    search_keywords = [

        'market', 'trend', 'news', 'relevant', 'price', 'rate', 'latest', 'today', 'now', 'current',

        'search', 'browse', 'web', 'internet', 'source', 'sources',

        'рынок', 'тренд', 'новост', 'цена', 'курс', 'сегодня', 'сейчас', 'актуальн',

        'поиск', 'найди', 'интернет', 'источник', 'ссылка'

    ]

    needs_search = any(keyword in message_lower for keyword in search_keywords)

    

    # Does it need document generation?

    doc_keywords = [

        "report", "plan", "presentation", "proposal", "document",

        "email", "e-mail", "letter",

        "отчет", "план", "презентац", "предложен", "документ",

        "письмо", "емейл", "e-mail"

    ]

    file_keywords = [

        "docx", "doc", "word",

        "pdf",

        "xlsx", "xls", "excel",

        "ворд", "док", "пдф", "эксел", "excel"

    ]

    needs_document = (

        any(keyword in message_lower for keyword in doc_keywords) or

        any(keyword in message_lower for keyword in file_keywords)

    )

    

    # Document analysis?

    analyze_keywords = ['analyze', 'check', 'evaluate', 'review', 'study', 'анализ', 'проверь', 'оцен', 'разбор']

    needs_analysis = any(keyword in message_lower for keyword in analyze_keywords)

    

    return {

        "needs_search": needs_search,

        "needs_document": needs_document,

        "needs_analysis": needs_analysis

    }



# === API ENDPOINTS ===



@app.route('/api/auth', methods=['POST'])

def authenticate():

    """Simple authentication"""

    data = request.json

    username = data.get('username', '')

    password = data.get('password', '')

    

    if username in USERS and USERS[username] == password:

        return jsonify({

            "success": True,

            "user": {

                "username": username,

                "tier": "premium" if username == "business" else "free"

            }

        })

    else:

        return jsonify({

            "success": False,

            "error": "Invalid username or password"

        }), 401



@app.route('/api/chat', methods=['POST'])

def chat():

    """Main AI chat"""

    try:

        data = request.json

        message = data.get('message', '')

        user_id = data.get('user_id', 'guest')

        is_guest = data.get('is_guest', True)

        

        intent = detect_intent(message)



        # Guest mode check

        if is_guest:

            if intent['needs_search'] or intent['needs_document'] or intent['needs_analysis']:

                return jsonify({

                    "success": True,

                    "message": "Эта функция доступна только после входа. Нажмите «Login» и войдите в аккаунт.",

                    "require_auth": True

                })

            # Limited functionality for guests

            if len(message) > 500:

                return jsonify({

                    "success": False,

                    "message": "Guest mode is limited to 500 characters. Please log in for full access.",

                    "require_auth": True

                })

            # Simple reply without search

            system_prompt = """You are the demo version of Stratex AI. Reply briefly and suggest logging in to unlock:

- Market and trend analysis

- Full report generation

- Working with documents

- Advanced features



Use context from previous messages when relevant.



Use context from previous messages when relevant."""

            

            conversation_history = db_get_history(user_id, limit=10)

            response = chat_with_ai(message, system_prompt, conversation_history)



            db_add_message(user_id, "user", message)

            db_add_message(user_id, "assistant", response)

            

            return jsonify({

                "success": True,

                "message": response,

                "is_demo": True,

                "upgrade_prompt": "💎 Log in for full access"

            })



        # Full functionality for authenticated users

        # Load history

        if user_id not in chat_histories:

            chat_histories[user_id] = []

        

        conversation_history = db_get_history(user_id, limit=20)

        

        # Web search if needed

        search_context = None

        search_results = None

        

        if intent['needs_search']:

            search_data = web_search(message)

            if search_data.get("error"):

                search_context = None

            elif search_data['answer']:

                search_context = f"Up-to-date information:\n{search_data['answer']}\n\n"

                search_context += "Sources:\n"

                for source in search_data['sources']:

                    search_context += f"- {source['title']}: {source['content']}\n"

                search_results = search_data

        

        # System prompt

        system_prompt = """You are Stratex AI, a professional business assistant.



Your specializations:

📊 Market and trend analysis (with web search)

📄 Report, plan, and document generation

💼 Business communication (emails, proposals)

📈 Strategic planning



Working principles:

1. Structured output — use headings, lists, and tables

2. Measurable — provide concrete KPIs, timelines, numbers

3. Professional — business tone

4. Practical — focus on execution

5. Use context from previous messages when relevant



Table formatting:

For tabular data, use a simple pipe-separated table (plain text, no markdown styling):

Header 1 | Header 2

--- | ---

Row 1 | Row 2



Document formats:

If the user asks for a report/plan — FIRST ask for the format:

- Brief (1 page)

- Standard (2–3 pages)

- Detailed (5+ pages)

- Presentation (slides)



After answering — generate in the chosen format."""



        if intent['needs_search'] and not search_context:

            system_prompt += "\n\nImportant: If the user asks for up-to-date web info and no web context is provided, explicitly say that live web search is unavailable right now and ask to enable TAVILY_API_KEY."

        

        # Generate response

        ai_response = chat_with_ai(

            message, 

            system_prompt, 

            conversation_history,

            search_context

        )

        ai_response = markdown_to_html(ai_response)



        # Auto-generate file for document requests

        if intent.get("needs_document"):

            try:

                file_type = detect_file_type(message)

                name = f"report_{int(datetime.now().timestamp())}"

                content = clean_document_content(strip_html(ai_response))



                if file_type == "excel":

                    path = generate_excel(content, name)

                elif file_type == "word":

                    path = generate_word(content, name)

                else:

                    path = generate_pdf(content, name)



                base_url = request.host_url.rstrip("/")

                download_url = f"{base_url}/api/download/{os.path.basename(path)}"



                doc_label = {"word": "DOCX", "excel": "XLSX", "pdf": "PDF"}.get(file_type, "DOC")

                link_label = {"word": "WORD", "excel": "EXCEL", "pdf": "PDF"}.get(file_type, "DOC")

                ai_response = (

                    f"Here is your {doc_label} document:<br>"

                    f"<a href=\"{download_url}\" target=\"_blank\" "

                    f"rel=\"noopener\" style=\"color:#6f47eb;font-weight:600;\">"

                    f"Download {link_label}</a>"

                )

            except Exception as e:

                print(f"File generation error: {e}")



        

        # Save to history

        db_add_message(user_id, "user", message)

        db_add_message(user_id, "assistant", ai_response)

        

        response_data = {

            "success": True,

            "message": ai_response,

            "is_demo": False

        }

        

        # Add sources if search was used

        if search_results:

            response_data["sources"] = search_results["sources"]

        

        return jsonify(response_data)

    

    except Exception as e:

        print(f"Chat error: {e}")

        return jsonify({

            "success": False,

            "error": str(e)

        }), 500



@app.route('/api/analyze-document', methods=['POST'])

def analyze_document():

    """Analyze an uploaded document"""

    try:

        # Auth check

        user_id = request.form.get('user_id', 'guest')

        is_guest = request.form.get('is_guest', 'true') == 'true'

        

        if is_guest:

            return jsonify({

                "success": False,

                "message": "Document analysis is available only for authenticated users",

                "require_auth": True

            }), 403

        

        # Get file

        if 'file' not in request.files:

            return jsonify({"success": False, "error": "No file uploaded"}), 400

        

        file = request.files['file']

        filename = file.filename

        

        # Read content

        file_bytes = file.read()

        

        if filename.endswith('.pdf'):

            text = extract_pdf_text(file_bytes)

        elif filename.endswith('.docx'):

            text = extract_docx_text(file_bytes)

        elif filename.endswith('.txt'):

            text = file_bytes.decode('utf-8')

        else:

            return jsonify({"success": False, "error": "Unsupported format"}), 400

        

        # AI analysis

        system_prompt = """You are an expert in business document analysis.



Analyze the document using these criteria:

1. **Document type** — what is it?

2. **Main content** — short summary

3. **Key points** — 3–5 main points

4. **Strengths** — what is done well

5. **Areas for improvement** — what can be improved

6. **Recommendations** — concrete next steps



Output format: structured with headings."""

        

        analysis = chat_with_ai(

            f"Analyze the document '{filename}':\n\n{text[:4000]}",

            system_prompt

        )

        

        return jsonify({

            "success": True,

            "filename": filename,

            "analysis": analysis,

            "document_length": len(text)

        })

    

    except Exception as e:

        print(f"Document analysis error: {e}")

        return jsonify({

            "success": False,

            "error": str(e)

        }), 500



@app.route('/api/market-analysis', methods=['POST'])

def market_analysis():

    """Dedicated endpoint for market analysis"""

    try:

        data = request.json

        query = data.get('query', '')

        is_guest = data.get('is_guest', True)

        

        if is_guest:

            return jsonify({

                "success": False,

                "message": "Market analysis is available only for authenticated users",

                "require_auth": True

            }), 403

        

        # Web search

        search_results = web_search(query)

        

        # AI analysis based on sources

        system_prompt = """You are a market analyst. Based on the information found, create:



1. **Current situation** — what is happening now

2. **Key trends** — 3–5 key trends

3. **Opportunities** — for businesses

4. **Risks** — potential threats

5. **Recommendations** — concrete actions



Use data from the sources and include numbers when possible."""

        

        context = f"Search results for '{query}':\n\n"

        for source in search_results.get('sources', []):

            context += f"- {source['title']}: {source['content']}\n"

        

        analysis = chat_with_ai(

            f"Analyze the market: {query}",

            system_prompt,

            search_context=context

        )

        

        return jsonify({

            "success": True,

            "query": query,

            "analysis": analysis,

            "sources": search_results.get('sources', [])

        })

    

    except Exception as e:

        print(f"Market analysis error: {e}")

        return jsonify({

            "success": False,

            "error": str(e)

        }), 500



@app.route('/api/generate-report', methods=['POST'])

def generate_report():

    """Generate a report/plan"""

    try:

        data = request.json

        report_type = data.get('type', '')

        parameters = data.get('parameters', {})

        format_type = data.get('format', 'standard')

        is_guest = data.get('is_guest', True)

        

        if is_guest:

            return jsonify({

                "success": False,

                "message": "Report generation is available only for authenticated users",

                "require_auth": True

            }), 403

        

        format_instructions = {

            'brief': "1 page, key points as a list",

            'standard': "2–3 pages, structured with sections",

            'detailed': "5+ pages, detailed analysis with examples",

            'presentation': "presentation slide format"

        }

        

        system_prompt = f"""You are an expert in business documentation.



Create a {report_type} report in the following format: {format_instructions.get(format_type, 'standard')}



Required elements:

- Executive summary

- Structured headings

- Concrete numbers and KPIs

- Timelines

- Recommendations



Parameters: {json.dumps(parameters, ensure_ascii=False)}"""

        

        report = chat_with_ai(

            f"Create a {report_type} report",

            system_prompt

        )

        

        return jsonify({

            "success": True,

            "report": report,

            "type": report_type,

            "format": format_type

        })

    

    except Exception as e:

        print(f"Report generation error: {e}")

        return jsonify({

            "success": False,

            "error": str(e)

        }), 500



@app.route('/api/status', methods=['GET'])

def status():

    """API status check"""

    return jsonify({

        "success": True,

        "status": "online",

        "openai": "configured" if OPENAI_API_KEY else "missing_key",

        "tavily": "configured" if TAVILY_API_KEY else "missing_key",

        "model": OPENAI_MODEL

    })



if __name__ == '__main__':

    print("🚀 Stratex AI Backend started!")

    print("📡 Open AI API: connected")

    print("🔍 Tavily Search: connected")

    print("🌐 Server: http://localhost:5000")

    print("\n📝 Test accounts:")

    print("   demo / demo123 (basic)")

    print("   business / business2024 (full)")

    

    app.run(debug=True, port=5000, host='0.0.0.0')

