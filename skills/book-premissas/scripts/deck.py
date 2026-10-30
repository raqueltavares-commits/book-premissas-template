"""Leitura e conferencia de um deck Google Slides.

  python deck.py get    PRES_ID saida.json          # baixa o deck inteiro (JSON)
  python deck.py text   deck.json [regex]           # lista textos (shapes e celulas) com ids
  python deck.py slides deck.json                   # ordem dos slides + titulo
  python deck.py thumbs PRES_ID pasta PAGE [PAGE..] # miniaturas PNG para conferir
"""
import json
import os
import re
import sys
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import composio  # noqa: E402


def txt(obj):
    return "".join(te["textRun"]["content"] for te in obj.get("text", {}).get("textElements", []) if "textRun" in te)


def iter_text(deck):
    """(page, objectId, row, col, text) para shapes e celulas de tabela."""
    def walk(pid, els):
        for e in els:
            if "shape" in e:
                yield pid, e["objectId"], None, None, txt(e["shape"])
            elif "table" in e:
                for r, row in enumerate(e["table"]["tableRows"]):
                    for c, cell in enumerate(row["tableCells"]):
                        yield pid, e["objectId"], r, c, txt(cell)
            elif "elementGroup" in e:
                yield from walk(pid, e["elementGroup"]["children"])
    for s in deck["slides"]:
        yield from walk(s["objectId"], s.get("pageElements", []))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__); return
    cmd = sys.argv[1]
    if cmd == "get":
        d = composio.get_deck(sys.argv[2])
        json.dump(d, open(sys.argv[3], "w", encoding="utf-8"), ensure_ascii=False)
        print(d.get("title"), "| slides:", len(d["slides"]), "| rev:", d.get("revisionId"))
    elif cmd == "text":
        d = json.load(open(sys.argv[2], encoding="utf-8"))
        d = d.get("data", d)
        pat = re.compile(sys.argv[3], re.I) if len(sys.argv) > 3 else None
        for pid, oid, r, c, t in iter_text(d):
            if t.strip() and (not pat or pat.search(t)):
                loc = f" r{r}c{c}" if r is not None else ""
                print(f"[{pid}] {oid}{loc}: {t!r}")
    elif cmd == "slides":
        d = json.load(open(sys.argv[2], encoding="utf-8"))
        d = d.get("data", d)
        for n, s in enumerate(d["slides"], 1):
            ts = [txt(e["shape"]).strip() for e in s.get("pageElements", []) if "shape" in e and txt(e["shape"]).strip()]
            print(n, s["objectId"], " | ".join(t.replace("\n", " ")[:60] for t in ts[:2]))
    elif cmd == "thumbs":
        pres, out = sys.argv[2], sys.argv[3]
        os.makedirs(out, exist_ok=True)
        for pg in sys.argv[4:]:
            r = composio.run("GOOGLESLIDES_PRESENTATIONS_PAGES_GET_THUMBNAIL", {"presentationId": pres, "pageObjectId": pg})
            d = r.get("data") or {}
            url = d.get("contentUrl") or (d.get("response_data") or {}).get("contentUrl")
            if not url:
                print(pg, "ERRO", str(r)[:200]); continue
            path = os.path.join(out, pg + ".png")
            open(path, "wb").write(urllib.request.urlopen(url).read())
            print(path)
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
