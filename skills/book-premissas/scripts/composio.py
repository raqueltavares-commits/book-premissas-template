"""Chama ferramentas do Composio por REST (Google Slides / Drive) como a conta da Raquel.

Uso CLI:
  python composio.py SLUG args.json [saida.json]

A API key e descoberta nos configs locais (~/.claude.json etc.) e nunca e impressa.
Conta: variavel de ambiente BOOK_USER (e-mail da conta Google conectada no Composio; obrigatoria).
"""
import json
import os
import re
import sys
import urllib.error
import urllib.request

BASE = "https://backend.composio.dev"
UID = os.environ.get("BOOK_USER") or sys.exit("ERRO: defina BOOK_USER com o e-mail da conta Google no Composio.")
_KEY = None


def _home():
    return os.environ.get("USERPROFILE") or os.path.expanduser("~")


def _key_works(key):
    try:
        req = urllib.request.Request(BASE + "/api/v3/mcp/servers?limit=1", method="GET")
        req.add_header("x-api-key", key)
        with urllib.request.urlopen(req, timeout=30) as r:
            json.loads(r.read().decode())
        return True
    except Exception:
        return False


def key():
    global _KEY
    if _KEY:
        return _KEY
    seen = []
    env = (os.environ.get("COMPOSIO_API_KEY") or "").strip()
    if env:
        seen.append(env)
    # Legacy fallback only: Claude Code configs. Never the desktop app config or memory notes.
    paths = [os.path.join(_home(), ".claude.json"),
             os.path.join(_home(), "Claude", ".claude", ".claude.json")]
    for p in paths:
        try:
            txt = open(p, encoding="utf-8").read()
        except OSError:
            continue
        for m in re.findall(r"ak_[A-Za-z0-9_-]{10,}", txt):
            if m not in seen:
                seen.append(m)
    for k in seen:
        if _key_works(k):
            _KEY = k
            return k
    sys.exit(f"ERRO: nenhuma API key do Composio valida (candidatas: {len(seen)}).")


def post(path, body):
    req = urllib.request.Request(BASE + path, data=json.dumps(body, ensure_ascii=False).encode(), method="POST")
    req.add_header("x-api-key", key())
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=180) as r:
        return json.load(r)


def run(slug, args):
    try:
        return post(f"/api/v3/tools/execute/{slug}", {"user_id": UID, "arguments": args})
    except urllib.error.HTTPError as e:
        return {"successful": False, "error": f"HTTP {e.code}: {e.read().decode()[:1500]}"}


def get_deck(pid):
    r = run("GOOGLESLIDES_PRESENTATIONS_GET", {"presentationId": pid})
    if not r.get("successful"):
        sys.exit(f"ERRO lendo {pid}: {r.get('error')}")
    return r["data"]


def batch(pid, requests, revision):
    r = run("GOOGLESLIDES_PRESENTATIONS_BATCH_UPDATE",
            {"presentationId": pid, "requests": requests, "writeControl": {"requiredRevisionId": revision}})
    if not r.get("successful"):
        sys.exit(f"ERRO no batch: {str(r.get('error'))[:1500]}")
    d = r.get("data") or {}
    return d.get("replies") or (d.get("response_data") or {}).get("replies") or []


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    out = run(sys.argv[1], json.load(open(sys.argv[2], encoding="utf-8")) if len(sys.argv) > 2 else {})
    s = json.dumps(out, ensure_ascii=False)
    if len(sys.argv) > 3:
        open(sys.argv[3], "w", encoding="utf-8").write(s)
        print("salvo", sys.argv[3], len(s))
    else:
        print(s[:4000])
