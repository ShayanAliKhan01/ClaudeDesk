from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from anthropic import Anthropic
import os
from dotenv import load_dotenv

load_dotenv()

client = Anthropic(
    api_key=os.getenv("ANTHROPIC_API_KEY"),
    base_url=os.getenv("ANTHROPIC_BASE_URL"),
    default_headers={
        "anthropic-workspace-id": os.getenv("ANTHROPIC_WORKSPACE_ID")
    }
)


app = FastAPI()

history = []

@app.get("/", response_class=HTMLResponse)
async def home():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/chat")
async def chat(request: Request):
    data = await request.json()
    user_msg = data["message"]
    history.append({"role": "user", "content": user_msg})

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=history
    )

    reply = response.content[0].text
    history.append({"role": "assistant", "content": reply})
    return {"reply": reply}

@app.post("/reset")
async def reset():
    history.clear()
    return {"status": "cleared"}