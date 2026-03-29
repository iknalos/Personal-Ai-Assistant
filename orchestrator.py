"""
AI Orchestrator — Gemini Edition
=================================
Natural language interface using Gemini as the brain.
Run: python orchestrator.py
Open: http://localhost:5500
"""

from flask import Flask, request, jsonify, send_from_directory
import requests
import re
import time
import datetime
import os
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()  # loads keys from .env file

app = Flask(__name__)

@app.after_request
def add_cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    return response

# ============================================================
# CONFIGURATION — fill these in
# ============================================================
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY", "")
ANYTHINGLLM_URL = "http://localhost:3001/api/v1"
ANYTHINGLLM_KEY = os.getenv("ANYTHINGLLM_KEY", "")
WORKSPACE_SLUG  = os.getenv("WORKSPACE_SLUG", "my-workspace")
# ============================================================

# Model rotation — best to worst, auto-fallback on rate limits
MODELS = [
    "gemini-2.5-pro",      # best — 2 RPM, 50 req/day free
    "gemini-2.5-flash",    # great — fast, generous free tier
    "gemini-2.0-flash",    # good  — 15 RPM free
    "gemini-1.5-flash",    # fallback — most generous free limits
]

current_model_index = 0
rate_limit_until    = [0.0] * len(MODELS)  # per-model cooldown timestamps


def get_model():
    """Return highest available model based on per-model cooldowns."""
    global current_model_index
    now = time.time()
    for i, model in enumerate(MODELS):
        if now >= rate_limit_until[i]:
            if i != current_model_index:
                current_model_index = i
                print(f"  Switched to {model}")
            return model
    return MODELS[-1]


def on_rate_limit(model_index, daily=False):
    """Set cooldown for a specific model — daily or per-minute."""
    if daily:
        midnight = (datetime.datetime.now() + datetime.timedelta(days=1)).replace(
            hour=0, minute=5, second=0, microsecond=0)
        rate_limit_until[model_index] = midnight.timestamp()
        print(f"  {MODELS[model_index]} daily limit hit — unavailable until midnight")
    else:
        rate_limit_until[model_index] = time.time() + 65
        print(f"  {MODELS[model_index]} rate limited — retrying in 65s")


SYSTEM_PROMPT = """You are a smart assistant that helps users manage files and data on their Windows PC.
The user's Windows username is: test
Default paths:
- Desktop:   C:\\Users\\test\\Desktop
- Documents: C:\\Users\\test\\Documents
- Downloads: C:\\Users\\test\\Downloads

You have two modes:

MODE 1 — TASK MODE
When the user asks to do something with files, folders, or Excel/data, respond with EXACTLY:
TASK: <the @agent command to send>
CONFIRM: <friendly one-sentence confirmation asking user to confirm>

Example:
User: "create a folder called Work on my desktop"
TASK: @agent use the create_folder tool with path "C:\\Users\\test\\Desktop\\Work"
CONFIRM: I'll create a folder called "Work" on your Desktop. Go ahead?

MODE 2 — CHAT MODE
When the user is chatting, asking questions, or following up — respond naturally.
Do NOT use TASK:/CONFIRM: format here.

MODE 3 — CLARIFICATION MODE
When the request is too vague, ask one specific question before acting.

RULES:
- Never invent file paths — always use C:\\Users\\test\\ as base.
- Always prefix agent commands with @agent.
- Be concise — no filler words like "Great!" or "Sure!".
- If user says yes/go ahead/do it/correct/ok → they are confirming.
- If user says no/stop/cancel/wrong → cancel and ask what they want instead.
- For destructive actions (delete) always confirm first."""


def chat_with_gemini(history, user_message):
    """Send message to Gemini with automatic model fallback on rate limit."""
    genai.configure(api_key=GEMINI_API_KEY)

    for attempt in range(len(MODELS) * 2):
        model_name  = get_model()
        model_index = MODELS.index(model_name)
        try:
            model = genai.GenerativeModel(
                model_name=model_name,
                system_instruction=SYSTEM_PROMPT
            )
            gemini_history = []
            for msg in history:
                if msg["role"] == "user":
                    gemini_history.append({"role": "user",  "parts": [msg["content"]]})
                elif msg["role"] == "assistant":
                    gemini_history.append({"role": "model", "parts": [msg["content"]]})

            chat = model.start_chat(history=gemini_history)
            response = chat.send_message(user_message)
            return response.text

        except Exception as e:
            err = str(e)
            if "429" in err or "quota" in err.lower() or "rate" in err.lower():
                daily = "daily" in err.lower() or "per_day" in err.lower() or "resource_exhausted" in err.lower()
                on_rate_limit(model_index, daily=daily)
                continue
            return f"Error connecting to Gemini: {err}"

    return "All models are currently rate limited. Try again in a minute."


def send_to_anythingllm(message):
    """Send an @agent command to AnythingLLM and get result."""
    try:
        headers = {
            "Authorization": f"Bearer {ANYTHINGLLM_KEY}",
            "Content-Type": "application/json"
        }
        response = requests.post(
            f"{ANYTHINGLLM_URL}/workspace/{WORKSPACE_SLUG}/chat",
            headers=headers,
            json={"message": message, "mode": "chat"},
            timeout=120
        )
        data = response.json()
        return data.get("textResponse", "No response from AnythingLLM")
    except Exception as e:
        return f"Error connecting to AnythingLLM: {str(e)}"


def summarize_with_gemini(original_request, agent_response, history):
    """Ask Gemini to summarize the agent result in plain English."""
    summary_prompt = f"""The user asked: "{original_request}"
The file agent responded with this EXACT result: "{agent_response}"

Instructions:
- If the result contains a list of files or folders, show ALL of them clearly to the user.
- If the result confirms a successful action, confirm it in one sentence.
- If the result contains an error, explain it simply.
- NEVER say "you can see the full list above" — always include the actual content in your response.
- Keep it concise but complete."""
    return chat_with_gemini(history, summary_prompt)


# In-memory state
conversation_history = []
pending_task = None


@app.route("/")
def index():
    import os
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), "index.html")


@app.route("/chat", methods=["POST"])
def chat():
    global conversation_history, pending_task

    user_message = request.json.get("message", "").strip()
    if not user_message:
        return jsonify({"response": "Please type a message.", "type": "error"})

    confirm_words = ["yes", "go ahead", "do it", "sure", "correct", "proceed", "ok", "okay", "yep", "yeah"]
    cancel_words  = ["no", "stop", "cancel", "don't", "mistake", "wrong", "nope"]

    # Handle pending confirmation
    if pending_task:
        if any(w in user_message.lower() for w in confirm_words):
            task_command = pending_task
            pending_task = None
            agent_response = send_to_anythingllm(task_command)
            summary = summarize_with_gemini(user_message, agent_response, conversation_history)
            conversation_history.append({"role": "assistant", "content": summary})
            return jsonify({"response": summary, "type": "success"})
        elif any(w in user_message.lower() for w in cancel_words):
            pending_task = None
            reply = "Cancelled. What would you like to do instead?"
            conversation_history.append({"role": "assistant", "content": reply})
            return jsonify({"response": reply, "type": "chat"})

    # Get Gemini response
    gemini_response = chat_with_gemini(conversation_history, user_message)
    conversation_history.append({"role": "user",      "content": user_message})
    conversation_history.append({"role": "assistant", "content": gemini_response})

    # Check if Gemini wants to execute a task
    if "TASK:" in gemini_response and "CONFIRM:" in gemini_response:
        task_match    = re.search(r"TASK:\s*(.+?)(?:\n|CONFIRM:)", gemini_response, re.DOTALL)
        confirm_match = re.search(r"CONFIRM:\s*(.+)", gemini_response, re.DOTALL)
        if task_match and confirm_match:
            pending_task = task_match.group(1).strip()
            confirm_msg  = confirm_match.group(1).strip()
            # Auto-confirm read operations — no need to ask user
            if any(x in pending_task.lower() for x in ["read_excel", "list_files", "search_files", "debug_excel"]):
                agent_response = send_to_anythingllm(pending_task)
                # Feed result back to Gemini to decide next step
                followup = chat_with_gemini(
                    conversation_history,
                    f"The file agent returned this data: {agent_response}\n\nNow proceed with the original task using this information. If you need to generate a chart or do further analysis, do it now using TASK:/CONFIRM: format."
                )
                conversation_history.append({"role": "assistant", "content": followup})
                if "TASK:" in followup and "CONFIRM:" in followup:
                    t = re.search(r"TASK:\s*(.+?)(?:\n|CONFIRM:)", followup, re.DOTALL)
                    c = re.search(r"CONFIRM:\s*(.+)", followup, re.DOTALL)
                    if t and c:
                        pending_task = t.group(1).strip()
                        return jsonify({
                            "response": c.group(1).strip() + "\n\n(Reply yes to confirm or no to cancel)",
                            "type": "confirm"
                        })
                return jsonify({"response": followup, "type": "chat"})
            return jsonify({
                "response": confirm_msg + "\n\n(Reply yes to confirm or no to cancel)",
                "type": "confirm"
            })

    return jsonify({"response": gemini_response, "type": "chat"})


@app.route("/reset", methods=["POST"])
def reset():
    global conversation_history, pending_task
    conversation_history = []
    pending_task = None
    return jsonify({"status": "ok"})


@app.route("/status", methods=["GET"])
def status():
    gemini_ok      = bool(GEMINI_API_KEY)
    anythingllm_ok = False
    try:
        headers = {"Authorization": f"Bearer {ANYTHINGLLM_KEY}"}
        r = requests.get(f"{ANYTHINGLLM_URL}/auth", headers=headers, timeout=5)
        anythingllm_ok = r.status_code == 200
    except:
        pass
    return jsonify({"gemini": gemini_ok, "anythingllm": anythingllm_ok})


if __name__ == "__main__":
    print("\n" + "="*50)
    print("  AI Orchestrator (Gemini Edition)")
    print("="*50)
    print(f"\n  Models: {' → '.join(MODELS)}")
    print("\n  CHECKLIST:")
    print("  1. AnythingLLM running?")
    print("  2. GEMINI_API_KEY filled in?")
    print("  3. ANYTHINGLLM_KEY filled in?")
    print("\n  Open browser at: http://localhost:5500")
    print("="*50 + "\n")
    app.run(host="0.0.0.0", port=5500, debug=False)