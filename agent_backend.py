import os
import json
import re
import sys
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

# Define dynamic paths
USER_PROFILE = os.environ.get("USERPROFILE") or os.path.expanduser("~")
GEMINI_DIR = os.path.join(USER_PROFILE, ".gemini", "antigravity")

def generate_self_signed_cert(cert_path="localhost.crt", key_path="localhost.key"):
    from cryptography import x509
    from cryptography.x509.oid import NameOID, ExtendedKeyUsageOID
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.hazmat.primitives import serialization
    import datetime

    # Generate private key
    key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )
    
    # Generate certificate
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "VN"),
        x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, "Hanoi"),
        x509.NameAttribute(NameOID.LOCALITY_NAME, "Hanoi"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "JP-VI Translator"),
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
    ])
    
    utc_now = datetime.datetime.now(datetime.timezone.utc) if hasattr(datetime, "timezone") else datetime.datetime.utcnow()
    
    cert = x509.CertificateBuilder().subject_name(
        subject
    ).issuer_name(
        issuer
    ).public_key(
        key.public_key()
    ).serial_number(
        x509.random_serial_number()
    ).not_valid_before(
        utc_now - datetime.timedelta(days=1)
    ).not_valid_after(
        # 10 years validity
        utc_now + datetime.timedelta(days=3650)
    ).add_extension(
        x509.BasicConstraints(ca=True, path_length=None),
        critical=True,
    ).add_extension(
        x509.KeyUsage(
            digital_signature=True,
            content_commitment=False,
            key_encipherment=True,
            data_encipherment=False,
            key_agreement=False,
            key_cert_sign=True,
            crl_sign=True,
            encipher_only=False,
            decipher_only=False,
        ),
        critical=True,
    ).add_extension(
        x509.ExtendedKeyUsage([
            ExtendedKeyUsageOID.SERVER_AUTH,
            ExtendedKeyUsageOID.CLIENT_AUTH,
        ]),
        critical=False,
    ).add_extension(
        x509.SubjectAlternativeName([
            x509.DNSName("localhost"),
            x509.DNSName("127.0.0.1"),
        ]),
        critical=False,
    ).sign(key, hashes.SHA256())
    
    # Write certificate to file
    with open(cert_path, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))
        
    # Write private key to file
    with open(key_path, "wb") as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption()
        ))


# Initialize FastAPI App
app = FastAPI(
    title="Antigravity Translation Agent Backend",
    description="Backend API for JP-VI Office Add-in translation using Google Antigravity / Gemini Agent",
    version="1.0.0"
)

# Configure CORS Middleware
# Essential because the Office Add-in runs on a different port (e.g., https://localhost:3000)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins for local development
    allow_credentials=True,
    allow_methods=["*"],  # Allow all HTTP methods (GET, POST, etc.)
    allow_headers=["*"],  # Allow all headers
)

# Define request/response structures
class TranslationRequest(BaseModel):
    text: str
    source_lang: str
    target_lang: str
    api_key: str = None
    model: str = "flash_lite"

class TranslationResponse(BaseModel):
    translation: str
    hiragana: str = ""

@app.get("/")
async def health_check():
    """
    Health check endpoint. The Office Add-in pings this endpoint to update
    the connection status LED (🟢 Connected) in the settings panel.
    """
    return {
        "status": "ok",
        "message": "Antigravity SDK Agent backend is running successfully.",
        "version": "1.0.0"
    }

def cleanup_conversation(conversation_id: str):
    import shutil
    import time
    
    # Wait briefly for files to close handles
    time.sleep(0.5)
    
    gemini_dir = GEMINI_DIR
    
    # 1. Clean conversations folder
    conv_dir = os.path.join(gemini_dir, "conversations")
    if os.path.exists(conv_dir):
        for ext in [".db", ".db-wal", ".db-shm", ".pb"]:
            file_path = os.path.join(conv_dir, f"{conversation_id}{ext}")
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Error removing conversation file {file_path}: {e}")

    # 2. Clean annotations folder
    ann_dir = os.path.join(gemini_dir, "annotations")
    if os.path.exists(ann_dir):
        file_path = os.path.join(ann_dir, f"{conversation_id}.pbtxt")
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Error removing annotation file {file_path}: {e}")

    # 3. Clean brain folder
    brain_dir = os.path.join(gemini_dir, "brain", conversation_id)
    if os.path.exists(brain_dir):
        try:
            shutil.rmtree(brain_dir)
        except Exception as e:
            print(f"Error removing brain folder {conversation_id}: {e}")

def call_antigravity_agent_api(text: str, src: str, tgt: str, model: str = "flash_lite") -> dict:
    import subprocess
    import time
    
    agentapi_path = os.path.join(
        USER_PROFILE,
        "AppData", "Local", "Programs",
        "Antigravity IDE", "resources", "app", "extensions", "antigravity", "bin",
        "language_server_windows_x64.exe"
    )
    if not os.path.exists(agentapi_path):
        raise FileNotFoundError(f"Antigravity Agent binary not found at: {agentapi_path}")

    prompt = f"""You are a professional Japanese-Vietnamese translation agent.
Translate the following text from {src.upper()} to {tgt.upper()}:
"{text}"

Requirements:
1. Translate accurately, keeping the original tone and layout.
2. If translating from Japanese to Vietnamese, also provide the Hiragana phonetics for the original Japanese text.
3. If translating from Vietnamese to Japanese, provide the Hiragana phonetics for the translated Japanese text.

Return your response in the following strict JSON format:
{{
  "translation": "your translation here",
  "hiragana": "hiragana representation of the Japanese text (e.g., にほんご / nihongo)"
}}
Ensure your output contains ONLY the JSON block. Do not wrap it in markdown or add explanations.
"""

    cmd = [agentapi_path, "agentapi", "new-conversation", f"--model={model}", prompt]
    result = subprocess.run(cmd, capture_output=True)
    stdout_str = result.stdout.decode('utf-8', errors='replace')
    stderr_str = result.stderr.decode('utf-8', errors='replace')
    if result.returncode != 0:
        raise Exception(f"Failed to start Antigravity agent: {stderr_str or stdout_str}")
        
    data = json.loads(stdout_str)
    conversation_id = data["response"]["newConversation"]["conversationId"]
    
    transcript_path = os.path.join(
        GEMINI_DIR, "brain",
        conversation_id,
        ".system_generated",
        "logs",
        "transcript.jsonl"
    )
    
    response_data = None
    try:
        # Poll for response
        for _ in range(15):  # Wait up to 15 seconds
            time.sleep(1)
            if not os.path.exists(transcript_path):
                continue
                
            with open(transcript_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
                
            for line in reversed(lines):
                try:
                    step = json.loads(line)
                    if step.get("source") == "MODEL" and step.get("type") == "PLANNER_RESPONSE":
                        content = step.get("content", "")
                        # Ensure it's a final response and not a tool call response
                        if content and not step.get("tool_calls"):
                            cleaned_content = re.sub(r"^```json\s*", "", content.strip())
                            cleaned_content = re.sub(r"\s*```$", "", cleaned_content)
                            try:
                                parsed_data = json.loads(cleaned_content)
                                response_data = {
                                    "translation": parsed_data.get("translation", ""),
                                    "hiragana": parsed_data.get("hiragana", "")
                                }
                            except Exception:
                                response_data = {"translation": content, "hiragana": ""}
                            break
                except Exception:
                    pass
            if response_data:
                break
    finally:
        # Clean up temporary conversation files
        cleanup_conversation(conversation_id)
        
    if response_data:
        return response_data
    raise Exception("Timeout waiting for Antigravity Agent response.")

@app.post("/translate", response_model=TranslationResponse)
async def translate(request: TranslationRequest):
    """
    Translation endpoint called by the Office Add-in.
    Accepts text and language codes, returns translation and phonetic hiragana.
    """
    text = request.text.strip()
    src = request.source_lang
    tgt = request.target_lang

    if not text:
        raise HTTPException(status_code=400, detail="Văn bản cần dịch không được để trống.")

    try:
        # Check if api_key is provided in the request payload, otherwise fall back to environment variable
        api_key = request.api_key or os.getenv("GEMINI_API_KEY")

        if api_key:
            # Mode 1: Translation using custom API Key (Gemini API)
            print("Using custom API Key for translation...")
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            prompt = f"""You are a professional Japanese-Vietnamese translation agent.
Translate the following text from {src.upper()} to {tgt.upper()}:
"{text}"

Requirements:
1. Translate accurately, keeping the original tone and layout.
2. If translating from Japanese to Vietnamese, also provide the Hiragana phonetics for the original Japanese text.
3. If translating from Vietnamese to Japanese, provide the Hiragana phonetics for the translated Japanese text.

Return your response in the following strict JSON format:
{{
  "translation": "your translation here",
  "hiragana": "hiragana representation of the Japanese text (e.g., にほんご / nihongo)"
}}
Ensure your output contains ONLY the JSON block. Do not wrap it in markdown or add explanations.
"""
            response = model.generate_content(prompt)
            raw_text = response.text.strip()

            cleaned_text = re.sub(r"^```json\s*", "", raw_text)
            cleaned_text = re.sub(r"\s*```$", "", cleaned_text)
            
            try:
                data = json.loads(cleaned_text)
                return TranslationResponse(
                    translation=data.get("translation", ""),
                    hiragana=data.get("hiragana", "")
                )
            except Exception as json_err:
                print("Failed to parse JSON response, falling back to raw output:", json_err)
                return TranslationResponse(translation=raw_text, hiragana="")
        else:
            # Mode 2: Call internal Antigravity Agent (without local API Key)
            try:
                print("No API Key configured. Attempting to call internal Antigravity Agent API...")
                agent_res = call_antigravity_agent_api(text, src, tgt, model=request.model)
                return TranslationResponse(
                    translation=agent_res.get("translation", ""),
                    hiragana=agent_res.get("hiragana", "")
                )
            except Exception as agent_err:
                print(f"Internal Antigravity Agent API call failed: {agent_err}")
                raise HTTPException(
                    status_code=500,
                    detail=f"Lỗi kết nối hoặc xử lý của Antigravity Agent: {str(agent_err)}"
                )

    except HTTPException:
        # Re-raise HTTPExceptions
        raise
    except Exception as e:
        print("Translation endpoint error:", e)
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/cleanup-history")
async def cleanup_history():
    """
    Scans and deletes all temporary translation conversations, preserving 
    the active developer conversations.
    """
    import shutil
    
    # IDs of important active development conversations to keep
    CONVERSATIONS_TO_KEEP = {
        "5aae4e8f-2ca2-4329-9402-17d64317e638", # Current conversation
        "63406780-76b6-4b11-99ce-498efcf02b07"  # Integrating Antigravity SDK Agent
    }

    gemini_dir = GEMINI_DIR
    deleted_count = 0
    
    try:
        # 1. Clean conversations folder
        conv_dir = os.path.join(gemini_dir, "conversations")
        if os.path.exists(conv_dir):
            for filename in os.listdir(conv_dir):
                conv_id = os.path.splitext(filename)[0]
                if conv_id.endswith(".db"):
                    conv_id = os.path.splitext(conv_id)[0]
                    
                if conv_id not in CONVERSATIONS_TO_KEEP:
                    file_path = os.path.join(conv_dir, filename)
                    try:
                        os.remove(file_path)
                        deleted_count += 1
                    except Exception as e:
                        print(f"Failed to delete {filename}: {e}")

        # 2. Clean annotations folder
        ann_dir = os.path.join(gemini_dir, "annotations")
        if os.path.exists(ann_dir):
            for filename in os.listdir(ann_dir):
                conv_id = os.path.splitext(filename)[0]
                if conv_id not in CONVERSATIONS_TO_KEEP:
                    file_path = os.path.join(ann_dir, filename)
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        print(f"Failed to delete {filename}: {e}")

        # 3. Clean brain folder
        brain_dir = os.path.join(gemini_dir, "brain")
        if os.path.exists(brain_dir):
            for foldername in os.listdir(brain_dir):
                if foldername not in CONVERSATIONS_TO_KEEP and foldername != "tempmediaStorage":
                    folder_path = os.path.join(brain_dir, foldername)
                    try:
                        shutil.rmtree(folder_path)
                    except Exception as e:
                        print(f"Failed to delete brain folder {foldername}: {e}")
                        
        return {"status": "ok", "message": f"Dọn dẹp hoàn tất. Đã xóa {deleted_count} tệp tin hội thoại."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi dọn dẹp lịch sử: {str(e)}")
# Serve static files from 'dist' directory if present
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

dist_dir = os.path.join(base_dir, "dist")
if os.path.exists(dist_dir):
    print(f"Mounting static files from: {dist_dir}")
    from fastapi.staticfiles import StaticFiles
    app.mount("/", StaticFiles(directory=dist_dir, html=True), name="static")
else:
    print(f"Warning: Static files directory 'dist' not found at {dist_dir}. Frontend will not be served.")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Antigravity Translation Agent Backend")
    parser.add_argument("--prod", action="store_true", help="Run in production HTTPS mode on port 3000")
    parser.add_argument("--host", default="127.0.0.1", help="Host address to bind to")
    parser.add_argument("--port", type=int, help="Port to run the server on")
    args = parser.parse_args()

    # Determine paths for certs
    cert_file = os.path.join(base_dir, "localhost.crt")
    key_file = os.path.join(base_dir, "localhost.key")
    
    # Generate certs if they don't exist
    if not os.path.exists(cert_file) or not os.path.exists(key_file):
        try:
            print("SSL certificates not found. Generating self-signed certificate for localhost...")
            generate_self_signed_cert(cert_file, key_file)
            print(f"Successfully generated certificate: {cert_file}")
        except Exception as e:
            print(f"Error generating self-signed certificate: {e}")
            
    is_prod = args.prod or (os.path.exists(cert_file) and os.path.exists(key_file))
    port = args.port or (3000 if is_prod else 8000)
    is_frozen = getattr(sys, 'frozen', False)
    
    if is_prod and os.path.exists(cert_file) and os.path.exists(key_file):
        print(f"Starting Antigravity SDK Agent translation server on HTTPS: https://localhost:{port}")
        if is_frozen:
            uvicorn.run(app, host=args.host, port=port, ssl_keyfile=key_file, ssl_certfile=cert_file)
        else:
            uvicorn.run("agent_backend:app", host=args.host, port=port, ssl_keyfile=key_file, ssl_certfile=cert_file, reload=True)
    else:
        print(f"Starting Antigravity SDK Agent translation server on HTTP: http://localhost:{port}")
        if is_frozen:
            uvicorn.run(app, host=args.host, port=port)
        else:
            uvicorn.run("agent_backend:app", host=args.host, port=port, reload=True)
