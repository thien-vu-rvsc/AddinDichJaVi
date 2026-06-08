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
        # 1. Antigravity Agent logic using Gemini API (Dual Mode: Live API / Mock Demo)
        api_key = os.getenv("GEMINI_API_KEY")

        if api_key:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            
            # Use gemini-1.5-flash for fast translations
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            # Agent Prompt to return JSON containing translation and phonetics
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
            # Request translation from LLM
            response = model.generate_content(prompt)
            raw_text = response.text.strip()

            # Clean markdown JSON formatting if the model wraps it
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
            # 2. Mock / Demo Response (if GEMINI_API_KEY is not set)
            # Provides a quick way to test the frontend integration offline
            print("GEMINI_API_KEY is not set. Running in demo mode.")
            
            if src == "ja":
                translation = f"[Demo Agent] Đây là bản dịch mẫu tiếng Việt cho: '{text}'"
                hiragana = "これはにほんごのふりがなです / kore wa nihongo no furigana desu"
            else:
                translation = f"[Demo Agent] Đây là bản dịch mẫu tiếng Nhật cho: '{text}'"
                hiragana = "はい、これはベトナムごからのほんやくです / hai, kore wa betonamugo kara no hon'yaku desu"
                
            return TranslationResponse(translation=translation, hiragana=hiragana)

    except Exception as e:
        print("Translation endpoint error:", e)
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    print("Starting Antigravity SDK Agent translation server on http://localhost:8000")
    uvicorn.run("agent_backend:app", host="127.0.0.1", port=8000, reload=True)
