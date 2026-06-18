# VO Generator

Reads a Word doc table with **VO draft** and **Audio file** columns, generates speech via the Eleven Labs API using the voice **Mikko**, and saves `.wav` files (44.1 kHz, 16-bit mono) into a folder named after the doc with `_audio` appended.

---

## Requirements

- Python 3
- An [Eleven Labs](https://elevenlabs.io) account with API access
- A voice named **Mikko** in your Eleven Labs account
- Eleven Labs API key with **Text to Speech**, **Voice Generation**, and **Voices** permissions

---

## Setup (one-time)

**1. Install dependencies**
```bash
pip3 install -r requirements.txt
```

**2. Configure your API key**

Copy `.env.example` to `.env` and fill in your key:
```bash
cp .env.example .env
```
Then edit `.env`:
```
ELEVENLABS_API_KEY=your-key-here
```

**3. Add the shell alias** (optional but recommended)

Add this to your `~/.zshrc` (or `~/.bashrc`):
```bash
alias vogen="python3 /path/to/VO-Generator/vo_generator.py"
```
Then reload:
```bash
source ~/.zshrc
```

---

## Usage

**Dry run** — parses the doc, previews rows, no API calls made:
```bash
vogen "/path/to/your/doc.docx" --dry-run
```

**Full run** — generates all `.wav` files:
```bash
vogen "/path/to/your/doc.docx"
```

Output lands in a folder alongside the doc, named `<DocName>_audio/`.

---

## Word Doc Requirements

- First two rows are label/header rows and are skipped automatically
- Column headers must contain **"VO draft"** and **"Audio file"** (case-insensitive)
- Audio file column values can use any extension — output is always `.wav`
