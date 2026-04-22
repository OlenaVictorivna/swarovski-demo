# 💎 Swarovski · AI-Readiness Scoring Demo

A Streamlit app that accepts any Swarovski Desktop Procedure PDF,
sends it to LLM (Claude or Gemini) with the full scoring prompt, and returns a
formatted AI-Readiness Scorecard as a `.docx` file.

---

## Quick Start

### 1. Prerequisites
- Python 3.10+
- An Anthropic API key (`sk-ant-...`) or Gemini API key

### 2. Install dependencies
```bash
cd swarovski_demo
pip install -r requirements.txt
```

### 3. Place the scoring prompt file
Place `scoring_prompt.docx` in the same folder as `app.py`:

```
swarovski_demo/
├── app.py
├── requirements.txt
├── README.md
└── scoring_prompt.docx   ← rename and place here
```

### 4. Set your API key (optional — the app will prompt you if not set)
```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

### 5. Run the app
```bash
streamlit run app.py
```
The browser opens at **http://localhost:8501**

---

On the Web UI: 
1. **Upload to the app** — drag the PDF into the uploader.

2. **Click Generate Scorecard**:
   - *"The app extracts the full PDF text"*
   - *"It sends it to agent with the scoring prompt as the system instruction"*
   - *"Agent scores 14 dimensions and returns structured JSON"*

3. **Download the .docx** — open it in Word and compare side-by-side
   with `EgenciaInvoice_Scorecard.docx` to show output fidelity.

---

## How it works (technical)

```
PDF upload
    │
    ▼
PyMuPDF text extraction
    │
    ▼
LLM API call
  • model: claude-opus-4-6 or gemini-3.1
  • system: full scoring_prompt.docx content
  • user: extracted procedure text
    │
    ▼
JSON response (14 dimension scores + rationales + quadrant)
    │
    ▼
python-docx scorecard builder
    │
    ▼
.docx download
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `scoring_prompt.docx` not found | Rename the prompt file and place it next to `app.py` |
| API error / timeout | Check your API key; procedure PDFs can be large — try with a shorter one first |
| JSON parse error | Claude occasionally adds commentary; this is handled automatically by stripping markdown fences |
| Blank scores | The procedure may be too short or not follow the standard template |
