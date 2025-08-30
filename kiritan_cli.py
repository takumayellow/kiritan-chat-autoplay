# kiritan_cli.py  (Windows / PowerShell用)
import os, re, subprocess, argparse, locale

# ---- Windowsの文字コード（CP932）に合わせる。環境変数で上書きも可。
ENC = os.getenv("SEIKA_ENCODING") or ("cp932" if os.name == "nt" else locale.getpreferredencoding(False))

SEIKA = os.environ.get("SEIKA_CLI")
if not SEIKA:
    raise SystemExit("環境変数 SEIKA_CLI が未設定です。")

def has_play_flag() -> bool:
    # 出力がSJIS系でも落ちないように errors='ignore'
    try:
        r = subprocess.run([SEIKA, "-h"], capture_output=True, text=True, encoding=ENC, errors="ignore", timeout=3)
        s = (r.stdout or "") + (r.stderr or "")
        return bool(re.search(r"(?i)\b-play\b", s))
    except Exception:
        return False

def speak(text: str, cid: int, speed: float, use_play: bool):
    text = (text or "").strip()
    if not text: return
    base = [SEIKA, "-cid", str(cid), "-speed", str(speed)]
    # 1回目：-play 付き（対応していなくても落ちないようにerrors='ignore'）
    cmd = base + (["-play"] if use_play else []) + ["-nc", "-t", text]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding=ENC, errors="ignore")
    # 失敗 or 未対応なら '-play' なしで再実行。ここは出力を読まない（=デコードしない）
    if r.returncode != 0 or "invalid option" in (r.stderr or "").lower():
        subprocess.run(base + ["-nc", "-t", text], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)

# ---- OpenAI（chat用）
def chat_once(prompt: str, model: str) -> str:
    from openai import OpenAI
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise SystemExit("OPENAI_API_KEY が未設定です。")
    client = OpenAI(api_key=key)
    sysmsg = "あなたは「東北きりたんEX」。親しみやすく自然に返答して。"
    r = client.chat.completions.create(
        model=model,
        messages=[{"role":"system","content":sysmsg},{"role":"user","content":prompt}],
    )
    return (r.choices[0].message.content or "").strip()

def main():
    p = argparse.ArgumentParser(prog="kiritan")
    p.add_argument("--cid", type=int, default=int(os.getenv("KIRITAN_CID","1707")))
    p.add_argument("--speed", type=float, default=float(os.getenv("KIRITAN_SPEED","1.0")))
    sub = p.add_subparsers(dest="cmd", required=True)

    s1 = sub.add_parser("say");  s1.add_argument("text", nargs="+")
    s2 = sub.add_parser("chat"); s2.add_argument("-t","--text", required=True)
    s2.add_argument("--model", default=os.getenv("OPENAI_MODEL","gpt-4o-mini"))

    args = p.parse_args()
    use_play = has_play_flag()

    if args.cmd == "say":
        speak(" ".join(args.text), args.cid, args.speed, use_play); return
    if args.cmd == "chat":
        reply = chat_once(args.text, args.model)
        print(f"[assistant] {reply}")
        speak(reply, args.cid, args.speed, use_play); return

if __name__ == "__main__":
    main()
