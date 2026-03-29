# Personal AI Assistant

A fully private, local AI assistant for Windows that lets you manage files, analyze Excel data, generate charts, and write documents, all through natural language. No data ever leaves your PC.

## What it does

- **File management** — create, move, copy, organize folders and files using plain English
- **Excel analysis** — read any spreadsheet and generate bar, line, pie, or scatter charts
- **Document writing** — write essays, cover letters, reports and save them as `.docx` files
- **Grammar correction** — paste text and get it corrected or rewritten in different styles
- **Natural language** — no commands needed, just talk to it like a person

## Architecture

```
You (browser at localhost:5500)
        ↓
Gemini API (understands your intent)
        ↓
orchestrator.py (translates to agent commands)
        ↓
AnythingLLM API (executes via MCP)
        ↓
ai_server.py (touches your actual files — 100% local)
```

Your files never leave your PC. Only your chat messages go to Gemini.

## Tech stack

- **[Ollama](https://ollama.com)** — runs local AI models (qwen2.5-coder:7b)
- **[AnythingLLM](https://anythingllm.com)** — local agent interface with MCP support
- **[Gemini API](https://aistudio.google.com)** — natural language understanding (free tier)
- **Python** — MCP server + Flask orchestrator
- **HTML/CSS/JS** — chat UI

## Setup

### Prerequisites
- Windows 10/11
- Python 3.11+ (via Miniconda recommended)
- [Ollama](https://ollama.com) installed
- [AnythingLLM Desktop](https://anythingllm.com/desktop) installed
- A free [Gemini API key](https://aistudio.google.com)

### Installation

1. **Clone the repo**
```bash
git clone https://github.com/yourusername/personal-ai-assistant
cd personal-ai-assistant
```

2. **Create conda environment**
```bash
conda create -n ai311 python=3.11
conda activate ai311
pip install -r requirements.txt
```

3. **Set up environment variables**

Copy `.env.example` to `.env` and fill in your keys:
```
GEMINI_API_KEY=your-gemini-key-here
ANYTHINGLLM_KEY=your-anythingllm-key-here
WORKSPACE_SLUG=my-workspace
```

4. **Pull the AI model**
```bash
ollama pull qwen2.5-coder:7b
```

5. **Configure AnythingLLM MCP**

In AnythingLLM → Agent Skills → MCP Servers, add:
```json
{
  "mcpServers": {
    "windows-file-server": {
      "command": "C:\\ProgramData\\Miniconda3\\envs\\ai311\\python.exe",
      "args": ["C:\\path\\to\\ai_server.py"]
    }
  }
}
```

### Running

1. Open **AnythingLLM**
2. Start the orchestrator:
```bash
conda activate ai311
python orchestrator.py
```
3. Open browser at **http://localhost:5500**

## Usage examples

```
create a folder called Projects on my desktop
move all PDFs from Downloads to Documents
organize my Desktop by file type
plot my sales data from test.xlsx
write a cover letter for a software engineer role
correct the grammar in this paragraph: [paste text]
```

## Privacy

- Your **files never leave your PC** — all file operations run locally via `ai_server.py`
- Only your **chat messages** are sent to Gemini for intent understanding
- The local AI model (qwen2.5-coder:7b) runs entirely on your hardware via Ollama
- No telemetry, no cloud storage, no data sharing

## License

MIT License — free to use, modify and distribute.
