# ResumeFlow AI — Local Setup

## Folder Structure
```
resume_converter/
├── app.py
├── requirements.txt
├── templates/
│   └── index.html
├── static/
│   ├── css/style.css
│   └── js/main.js
├── uploads/     (auto-created)
└── outputs/     (auto-created)
```

## Setup (one time)

```bash
# 1. Create virtual environment
python -m venv venv

# 2. Activate it
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

## Add your OpenRouter API Key

Open `app.py` and replace line:
```python
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "your-api-key-here")
```
Either set as environment variable OR paste key directly:
```python
OPENROUTER_API_KEY = "sk-or-v1-xxxxxxxxxxxx"
```

## Run

```bash
python app.py
```

Then open: http://localhost:5000

## How it works
1. Upload candidate PDF resume
2. Select company preset (Google / Amazon / Microsoft etc.) or write custom format
3. Click Convert — AI reads PDF, reformats, builds DOCX
4. Preview in browser, download DOCX

## Get OpenRouter API Key
→ https://openrouter.ai/keys (free tier available)
