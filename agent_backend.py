import os
import json
import re
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

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

def call_antigravity_agent_api(text: str, src: str, tgt: str) -> dict:
    import subprocess
    import time
    
    agentapi_path = r"c:\Users\fssv-vu-thien\AppData\Local\Programs\Antigravity IDE\resources\app\extensions\antigravity\bin\language_server_windows_x64.exe"
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

    cmd = [agentapi_path, "agentapi", "new-conversation", "--model=flash_lite", prompt]
    result = subprocess.run(cmd, capture_output=True)
    stdout_str = result.stdout.decode('utf-8', errors='replace')
    stderr_str = result.stderr.decode('utf-8', errors='replace')
    if result.returncode != 0:
        raise Exception(f"Failed to start Antigravity agent: {stderr_str or stdout_str}")
        
    data = json.loads(stdout_str)
    conversation_id = data["response"]["newConversation"]["conversationId"]
    
    transcript_path = os.path.join(
        r"C:\Users\fssv-vu-thien\.gemini\antigravity\brain",
        conversation_id,
        ".system_generated",
        "logs",
        "transcript.jsonl"
    )
    
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
                            return {
                                "translation": parsed_data.get("translation", ""),
                                "hiragana": parsed_data.get("hiragana", "")
                            }
                        except Exception:
                            return {"translation": content, "hiragana": ""}
            except Exception:
                pass
                
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
                agent_res = call_antigravity_agent_api(text, src, tgt)
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

if __name__ == "__main__":
    print("Starting Antigravity SDK Agent translation server on http://localhost:8000")
    uvicorn.run("agent_backend:app", host="127.0.0.1", port=8000, reload=True)
