# invoice-generator
# 🧾 AI Invoice Generator

A modern, responsive invoice generator built with **Streamlit**. Fill in the form, watch the invoice update live, then download it as a **PDF** or **Excel** file. An optional **Gemini Flash** assistant can build line items from plain text, write notes/terms and polish item descriptions.

**No paid APIs needed.** Everything works offline except the optional AI tab, which uses Google's free Gemini API key.

## ✨ Features

| Area | What you get |
|---|---|
| Business | Name, logo upload (PNG/JPG/WEBP), email, phone, address, tax ID |
| Customer | Name, email, phone, billing address |
| Invoice | Invoice number, dates, 10 currencies, brand colour, notes, terms |
| Items | Add / edit / delete unlimited products in an editable table |
| Maths | Subtotal, discount (% or fixed), tax %, shipping, grand total (exact `Decimal` arithmetic, rounded half-up) |
| Preview | Live, printable invoice that updates as you type |
| Export | PDF (multi-page, ReportLab) and Excel (live formulas, editable) |
| Reset | One click to clear the form, or load sample data |
| AI (optional) | Smart-fill from text, write notes & terms, polish descriptions |

## 📁 Files

```
app.py            # the whole application
requirements.txt  # Python dependencies
README.md         # this file
```

## 🚀 Run locally

Requires **Python 3.10+**.

```bash
# 1. (recommended) create a virtual environment
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate

# 2. install dependencies
pip install -r requirements.txt

# 3. start the app
streamlit run app.py
```

The app opens at http://localhost:8501.

## 🤖 Enable the AI assistant (optional, free)

1. Go to https://aistudio.google.com/apikey and create a free API key.
2. Paste it into **"Gemini API key"** in the app sidebar, **or** store it as a secret (recommended):

   * Locally: create `.streamlit/secrets.toml`
     ```toml
     GEMINI_API_KEY = "your-key-here"
     ```
   * On Streamlit Cloud: see the deploy section below.

The default model is `gemini-flash-latest` (a Google-maintained alias that always points to the newest Flash model). If it is unavailable or out of free quota, the app automatically tries `gemini-flash-lite-latest` and `gemini-2.5-flash`. You can type any other model name in the sidebar.

> Free-tier limits and model names change from time to time. If you see a quota error, wait a minute or change the model name in the sidebar.

## 🐙 Push to GitHub

1. Create an account at https://github.com and click **New repository**. Name it e.g. `ai-invoice-generator`, keep it **Public** (needed for free Streamlit Cloud), and do **not** tick "Add a README".
2. In your project folder (containing the 3 files), run:

```bash
# protect secrets: never commit your API key
printf "venv/\n__pycache__/\n*.pyc\n.streamlit/secrets.toml\n.env\n" > .gitignore

git init
git add .
git commit -m "Initial commit: AI Invoice Generator"
git branch -M main
git remote add origin https://github.com/<your-username>/ai-invoice-generator.git
git push -u origin main
```

3. When Git asks for a password, use a **Personal Access Token** (GitHub → Settings → Developer settings → Personal access tokens → Fine-grained → give it "Contents: Read and write" for this repo), or sign in via the GitHub CLI: `gh auth login`.

Later updates:

```bash
git add .
git commit -m "Describe your change"
git push
```

## ☁️ Deploy on Streamlit Community Cloud (free)

1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click **Create app** → **Deploy a public app from GitHub**.
3. Choose your repository, branch `main`, and main file path `app.py`.
4. Open **Advanced settings**:
   * Choose Python **3.11** or **3.12**.
   * In **Secrets**, paste (optional, for AI features):
     ```toml
     GEMINI_API_KEY = "your-key-here"
     ```
5. Click **Deploy**. After a minute or two you get a public `https://<name>.streamlit.app` URL. Every `git push` redeploys automatically.

## 🛠️ Troubleshooting

* **Print button does nothing** – some browsers block printing from inside an embedded frame. Use **Download PDF** instead (it is the same invoice).
* **Non-Latin text (e.g. Urdu, Arabic, Chinese) is missing in the PDF** – the PDF uses built-in Latin fonts. The live preview and Excel file display all languages correctly.
* **Currency symbols** – PKR/INR/AED/SAR use text prefixes (`Rs`, `AED`, `SAR`) so they render correctly in every PDF viewer.
* **AI says "API key not valid"** – re-copy the key from Google AI Studio; make sure there are no spaces.
* **Excel shows empty totals in a phone preview** – totals are real formulas. Open the file in Excel / LibreOffice / Google Sheets and they calculate instantly.
* **Streamlit error about `width="stretch"` or `st.iframe`** – upgrade Streamlit: `pip install -U streamlit`.

## 🔒 Privacy

All invoice data stays in your browser session and the app server's memory. Nothing is saved to disk. Text is sent to Google only when you press one of the AI buttons.
