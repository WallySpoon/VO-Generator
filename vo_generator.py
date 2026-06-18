"""
VO Generator — Batch Text-to-Speech from a Word Doc
=====================================================
Reads a .docx table with "VO draft" and "Audio file" columns,
generates speech via Eleven Labs API using voice "Mikko",
and saves .wav files (44.1 kHz, 16-bit mono) to the same
folder as the source document.

Usage:
    python vo_generator.py <path_to_docx>

Environment:
    ELEVENLABS_API_KEY — your Eleven Labs API key
    (or pass via --api-key flag)

Requirements:
    pip install python-docx requests
"""

import argparse
import io
import os
import struct
import sys
import time
import wave

import requests
from docx import Document
from dotenv import load_dotenv

load_dotenv()

# ─── Configuration ────────────────────────────────────────────

ELEVENLABS_BASE = "https://api.elevenlabs.io/v1"
VOICE_NAME = "Mikko"
MODEL_ID = "eleven_multilingual_v2"
OUTPUT_FORMAT = "pcm_44100"  # Raw PCM; requires Pro tier for 44.1 kHz
SAMPLE_RATE = 44100
SAMPLE_WIDTH = 2  # 16-bit
CHANNELS = 1      # Mono


# ─── Helpers ──────────────────────────────────────────────────

def get_headers(api_key: str) -> dict:
    return {
        "xi-api-key": api_key,
        "Content-Type": "application/json",
    }


def find_voice_id(api_key: str, voice_name: str) -> str:
    """Look up a voice by name and return its ID."""
    url = f"{ELEVENLABS_BASE}/voices"
    resp = requests.get(url, headers={"xi-api-key": api_key})
    resp.raise_for_status()
    voices = resp.json().get("voices", [])
    for v in voices:
        if v["name"].lower() == voice_name.lower():
            return v["voice_id"]
    available = [v["name"] for v in voices]
    raise ValueError(
        f'Voice "{voice_name}" not found. Available voices: {available}'
    )


def pcm_to_wav(pcm_data: bytes, out_path: str):
    """Wrap raw PCM bytes in a WAV container and write to disk."""
    with wave.open(out_path, "wb") as wf:
        wf.setnchannels(CHANNELS)
        wf.setsampwidth(SAMPLE_WIDTH)
        wf.setframerate(SAMPLE_RATE)
        wf.writeframes(pcm_data)


def generate_speech(api_key: str, voice_id: str, text: str) -> bytes:
    """Call the Eleven Labs TTS endpoint and return raw PCM bytes."""
    url = f"{ELEVENLABS_BASE}/text-to-speech/{voice_id}"
    params = {"output_format": OUTPUT_FORMAT}
    payload = {
        "text": text,
        "model_id": MODEL_ID,
    }
    resp = requests.post(
        url, headers=get_headers(api_key), params=params, json=payload
    )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Eleven Labs API error {resp.status_code}: {resp.text[:300]}"
        )
    return resp.content


def read_vo_table(docx_path: str) -> list[dict]:
    """
    Open a .docx, find the first table with 'VO draft' and 'Audio file'
    columns, and return a list of row dicts.
    """
    doc = Document(docx_path)
    for table in doc.tables:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        headers_lower = [h.lower() for h in headers]

        # Flexible matching — look for columns containing the key phrases
        vo_idx = next(
            (i for i, h in enumerate(headers_lower) if "vo draft" in h), None
        )
        audio_idx = next(
            (i for i, h in enumerate(headers_lower) if "audio file" in h), None
        )

        if vo_idx is not None and audio_idx is not None:
            rows = []
            for row in table.rows[2:]:  # skip 2 label rows
                vo_text = row.cells[vo_idx].text.strip()
                audio_name = row.cells[audio_idx].text.strip()
                if vo_text and audio_name:
                    rows.append({"vo_text": vo_text, "audio_file": audio_name})
            return rows

    raise ValueError(
        "No table found with both 'VO draft' and 'Audio file' columns. "
        "Check column headers in the document."
    )


def ensure_wav_extension(filename: str) -> str:
    """Strip any existing extension (e.g. .mp3) and ensure .wav."""
    base, _ = os.path.splitext(filename)
    return base + ".wav"


# ─── Main ─────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate VO .wav files from a Word doc table via Eleven Labs."
    )
    parser.add_argument("docx_path", help="Path to the .docx file with the VO table")
    parser.add_argument(
        "--api-key",
        default=os.environ.get("ELEVENLABS_API_KEY"),
        help="Eleven Labs API key (or set ELEVENLABS_API_KEY env var)",
    )
    parser.add_argument(
        "--voice",
        default=VOICE_NAME,
        help=f'Voice name to use (default: "{VOICE_NAME}")',
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse the doc and list rows without generating audio",
    )
    args = parser.parse_args()

    if not args.api_key:
        print("ERROR: No API key provided.")
        print("  Set ELEVENLABS_API_KEY or pass --api-key <key>")
        sys.exit(1)

    docx_path = os.path.abspath(args.docx_path)
    if not os.path.isfile(docx_path):
        print(f"ERROR: File not found: {docx_path}")
        sys.exit(1)

    doc_stem = os.path.splitext(os.path.basename(docx_path))[0]
    output_dir = os.path.join(os.path.dirname(docx_path), f"{doc_stem}_audio")
    os.makedirs(output_dir, exist_ok=True)

    # ── Parse the document ──
    print(f"\n📄 Reading: {docx_path}")
    rows = read_vo_table(docx_path)
    print(f"   Found {len(rows)} VO row(s)\n")

    for i, row in enumerate(rows, 1):
        fname = ensure_wav_extension(row["audio_file"])
        text_preview = row["vo_text"][:60] + ("…" if len(row["vo_text"]) > 60 else "")
        print(f"   [{i}] {fname}  ←  \"{text_preview}\"")

    if args.dry_run:
        print("\n✅ Dry run complete — no audio generated.")
        return

    # ── Resolve voice ──
    print(f'\n🔍 Looking up voice "{args.voice}"...')
    voice_id = find_voice_id(args.api_key, args.voice)
    print(f"   Found: {voice_id}\n")

    # ── Generate audio ──
    successes = 0
    failures = 0

    for i, row in enumerate(rows, 1):
        fname = ensure_wav_extension(row["audio_file"])
        out_path = os.path.join(output_dir, fname)
        text_preview = row["vo_text"][:50] + ("…" if len(row["vo_text"]) > 50 else "")

        print(f"   [{i}/{len(rows)}] Generating: {fname}")
        print(f"            Text: \"{text_preview}\"")

        try:
            pcm_data = generate_speech(args.api_key, voice_id, row["vo_text"])
            pcm_to_wav(pcm_data, out_path)
            duration = len(pcm_data) / (SAMPLE_RATE * SAMPLE_WIDTH * CHANNELS)
            print(f"            ✅ Saved ({duration:.1f}s audio)\n")
            successes += 1
        except Exception as e:
            print(f"            ❌ Failed: {e}\n")
            failures += 1

        # Respect rate limits — brief pause between requests
        if i < len(rows):
            time.sleep(0.5)

    # ── Summary ──
    print("─" * 50)
    print(f"Done.  ✅ {successes} succeeded  |  ❌ {failures} failed")
    print(f"Files saved to: {output_dir}")


if __name__ == "__main__":
    main()
