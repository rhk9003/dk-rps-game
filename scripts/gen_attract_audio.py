"""用 ElevenLabs API 產生 idle attract MP3。

執行前：
    pip install requests python-dotenv
    cp .env.example .env  # 填入 ELEVENLABS_API_KEY 與 ELEVENLABS_VOICE_ID
    python scripts/gen_attract_audio.py

想改廣告詞：編輯下方 PHRASES、重跑即可。
"""
import os
import pathlib
import requests
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.environ["ELEVENLABS_API_KEY"]
VOICE_ID = os.environ["ELEVENLABS_VOICE_ID"]
MODEL = "eleven_multilingual_v2"

PHRASES = [
    ("attract_1.mp3", "來玩猜拳～通通有獎喔～"),
    ("attract_2.mp3", "贏家抽 500 元折價券！"),
    ("attract_3.mp3", "免費的喔～來玩看看～"),
]

out_dir = pathlib.Path(__file__).parent.parent / "assets" / "audio"
out_dir.mkdir(parents=True, exist_ok=True)

for filename, text in PHRASES:
    r = requests.post(
        f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE_ID}",
        headers={"xi-api-key": API_KEY, "Content-Type": "application/json"},
        json={
            "text": text,
            "model_id": MODEL,
            "voice_settings": {
                "stability": 0.5,
                "similarity_boost": 0.75,
                "style": 0.3,
            },
        },
    )
    r.raise_for_status()
    (out_dir / filename).write_bytes(r.content)
    print(f"✓ {filename}  ({len(r.content)/1024:.1f} KB)  — {text}")
