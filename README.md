# VO Generator

Reads a Word doc table with **VO draft** and **Audio file** columns, generates speech via the Eleven Labs API using the voice **Mikko** and the newest flagship model (currently **Eleven v4**), and returns a downloadable `.zip` of `.wav` files (44.1 kHz, 16-bit mono).

---

## Option 1 — Docker (recommended for sharing)

No Python or pip required. Just Docker Desktop.

**1. Install Docker Desktop**
Download from [docker.com/products/docker-desktop](https://www.docker.com/products/docker-desktop) and launch it.

**2. Pull and run the image**
```bash
docker pull wallyspoon/vo-generator
docker run -p 5000:5000 wallyspoon/vo-generator
```

**3. Open your browser**
Go to [http://localhost:5000](http://localhost:5000)

On first launch you'll be prompted to enter your Eleven Labs API key. It's saved locally and you won't see that screen again.

---

## Option 2 — Run locally (for developers)

**1. Clone the repo**
```bash
git clone https://github.com/WallySpoon/VO-Generator.git
cd VO-Generator
```

**2. Install dependencies**
```bash
pip3 install -r requirements.txt
```

**3. Add the shell alias**

Add this to your `~/.zshrc` (or `~/.bashrc`):
```bash
alias vogen="python3 /path/to/VO-Generator/app.py"
```
Then reload:
```bash
source ~/.zshrc
```

**4. Run it**
```bash
vogen
```

Your browser opens automatically. On first launch you'll be prompted for your Eleven Labs API key.

---

## Eleven Labs API Key

You'll need a free or paid [Eleven Labs](https://elevenlabs.io) account.

1. Go to **Profile → API Keys**
2. Create a new key with these permissions:
   - Text to Speech → **Access**
   - Voice Generation → **Access**
   - Voices → **Read**
   - Models → **Read** (lets the app find the newest model and fill the model dropdown; without it, the app falls back to `eleven_v4`)
3. Paste the key into the setup screen on first launch

---

## Word Doc Requirements

- First two rows are label/header rows and are skipped automatically
- Column headers must contain **"VO draft"** and **"Audio file"** (case-insensitive)
- Audio file column values can use any extension — output is always `.wav`
- Upload the `.docx` via the browser — output is downloaded as a `.zip`