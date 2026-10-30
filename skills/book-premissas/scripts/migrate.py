"""Aplica as regras do v3 numa copia v2 do book (layout atual da Rachel), a partir de um spec JSON.

  python migrate.py fase1 spec.json [--dry]   # textos, status, linhas/slide de unidades, duplica slide de materiais
  python migrate.py fase2 spec.json [--dry]   # slide de MDFs, amostras, rodape, numeracao, negrito, capa, divisor

Sempre: fase1 -> fase2 -> conferir miniaturas (deck.py thumbs). --dry so imprime o que faria.

spec.json:
{
 "presentation_id": "ID da copia v2",
 "spot": "BATEL SPOT",                 # prefixo do rodape, igual ao book
 "estilo": "Bruma",                    # entra na capa e no slide 2
 "trocas": [["p5", "texto antigo exato", "texto novo"], [null, "global", "..."]],
 "remover_linhas": {"p21_i11": [3, 2, 1]},      # linhas de unidades (indices do deck ORIGINAL)
 "remover_slides": ["p18"],                      # slide "Unidades: acabamentos e escopo"
 "slide_materiais": "p20",                       # slide de 6 cards a duplicar para o de MDFs
 "slide_mdf": {"titulo": "Materiais · MDFs e metais (3/3)",
               "subtitulo": "Padrões Berneck ... (rev. DD/MM/AAAA). Referência visual, não substitui amostra física.",
               "cards": [{"tag": "T25", "titulo": "MDF Crômio", "aplicacao": "Aplicação: ...",
                          "especificacao": "Especificação: ...", "obra": "Obra: ...", "imagem": "DRIVE_FILE_ID"}]},
 "trocar_imagens": {"p20_i26": "DRIVE_FILE_ID"}, # amostras circulares/retangulares de cards existentes
 "status_definido": {"p21_i11": [4]}             # opcional: linhas cuja decisao JA esta escrita no deck
}
As trocas de nome ("Interiores das áreas comuns", estilo na capa) sao automaticas. Status -> DEFINIDO verde so
nas linhas com decisao: alguma `troca` muda texto de outra celula da linha, ou a linha esta em `status_definido`.
As demais ficam como estao e saem listadas como "SEM DECISAO" (decidir e rodar de novo).

Detalhes do spec:
- trocas: [pagina, texto_exato, novo]. pagina = objectId do slide (ex. "p5") ou null = deck inteiro. Funciona em
  shapes E celulas de tabela. Copie o texto exato de `deck.py text` (sem o \\n final). Trocas rodam na ordem.
- remover_linhas / status_definido: indices 0-based contando o cabecalho (linha 0 = cabecalho), do deck ANTES da fase1.
- remover_slides / slide_materiais: objectId do slide (ver `deck.py slides`).
- O slide duplicado ganha o id "<slide_materiais>mdf" e seus elementos "<slide_materiais>mdf_iN".
- slide_mdf: 1 a 6 cards (o slide base tem 6; sobras sao apagadas). Tags continuam a numeracao do book (T25...).
  Requer o padrao de ids dos books da Rachel: card k = imagem _i(12+7k), tag +1, titulo +2, aplicacao +3,
  especificacao +4, obra +5; titulo _i6, subtitulo _i10. Se o deck tiver outro padrao (ids "g3..."), a fase2
  aborta: preencha o slide de MDFs com `trocas` + `trocar_imagens` usando os ids reais.
- Deteccao automatica: capa = 1o slide, slide 2 = contextualizacao, encerramento = ultimo slide, divisor 02 =
  shape "Premissas de interiores" (subtitulo = id seguinte).
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import composio  # noqa: E402
from deck import txt  # noqa: E402

A = "DAS ÁREAS COMUNS"
a = "das áreas comuns"
GREEN = {"opaqueColor": {"rgbColor": {"red": 0x1E / 255, "green": 0x7F / 255, "blue": 0x5C / 255}}}
LABELS = ("Aplicação: ", "Especificação: ", "Obra: ")
EMU_DESCE_SUBTITULO = 520000


def _u16(s):
    """Tamanho em unidades UTF-16 (indices da API do Slides)."""
    return len(s.encode("utf-16-le")) // 2


def rat(page, old, new):
    r = {"replaceAllText": {"containsText": {"text": old, "matchCase": True}, "replaceText": new}}
    if page:
        r["replaceAllText"]["pageObjectIds"] = [page]
    return r


def shapes(deck):
    for s in deck["slides"]:
        for e in s.get("pageElements", []):
            if "shape" in e:
                yield s["objectId"], e, txt(e["shape"]).rstrip("\n")


def tables(deck):
    for s in deck["slides"]:
        for e in s.get("pageElements", []):
            if "table" in e:
                yield s["objectId"], e


def auto_naming(deck, estilo):
    """Trocas automaticas: 'Interiores das áreas comuns' em capa/títulos/rodapé e estilo na capa."""
    reqs = []
    first, last = deck["slides"][0]["objectId"], deck["slides"][-1]["objectId"]
    pages_02 = set()
    for pid, e, t in shapes(deck):
        if t.startswith("02 · INTERIORES") and A not in t:
            pages_02.add(pid)
    for p in sorted(pages_02):
        reqs.append(rat(p, "02 · INTERIORES", f"02 · INTERIORES {A}"))
    reqs.append(rat(None, "FACHADA E INTERIORES  ·", f"FACHADA E INTERIORES {A}  ·"))
    for pid, e, t in shapes(deck):
        if pid == first and t == "BOOK DE PREMISSAS DE FACHADA E INTERIORES":
            reqs.append(rat(pid, t, f"{t} {A}"))
        if pid == first and "·  Interiores" in t and a not in t:
            reqs.append(rat(pid, t.strip("\n"), t.strip("\n").replace("·  Interiores", f"·  Interiores {a}") + f"  ·  Estilo {estilo}"))
        if pid == last and t == "Book de Premissas de Fachada e Interiores":
            reqs.append(rat(pid, t, t + " das Áreas Comuns"))
        if t == "Premissas de interiores":
            reqs.append(rat(pid, t, f"{t} {a}"))
        if t == "Comparativo: interiores":
            reqs.append(rat(pid, t, f"{t} {a}"))
        if t == "Qual é o conceito de interiores que deve ser seguido?":
            reqs.append(rat(pid, t, f"Qual é o conceito de interiores {a} que deve ser seguido?"))
        if t == "Quais pontos ainda precisam ser validados?":
            reqs.append(rat(pid, t, "O que a obra precisa observar na execução?"))
        if t == "fachada e interiores":
            reqs.append(rat(pid, t, f"fachada e interiores {a}"))
        if t == "Interiores" and pid == deck["slides"][1]["objectId"]:
            reqs.append({"insertText": {"objectId": e["objectId"], "insertionIndex": len(t), "text": f" {a}"}})
        m = re.search(r"premissas de fachada e interiores do ", t)
        if m:
            reqs.append(rat(pid, "premissas de fachada e interiores do ", f"premissas de fachada e de interiores {a} do "))
        for old in ("onde divergem e o que validar antes da compra e da execução.", "onde divergem e o que validar antes da compra."):
            if old in t:
                reqs.append(rat(pid, old, "onde divergem e como a obra deve executar."))
    return reqs


STATUS_ABERTO = ("A VALIDAR", "PENDENTE", "A COTAR")


def _linha_decidida(pid, cells, trocas):
    """A linha tem decisao no spec se alguma troca (do slide ou global) altera o texto
    de outra celula da mesma linha."""
    textos = [txt(c) for c in cells if txt(c).strip() not in STATUS_ABERTO]
    return any((p is None or p == pid) and old and old != new and any(old in t for t in textos)
               for p, old, new in trocas)


def status_cells(deck, trocas=(), status_definido=None):
    """Status aberto -> DEFINIDO so quando ha decisao para a linha (regras.md: todo item
    vira DEFINIDO *com a decisao escrita*). Sem decisao, mantem e avisa."""
    reqs = []
    status_definido = status_definido or {}
    for pid, e in tables(deck):
        for r, row in enumerate(e["table"]["tableRows"]):
            decidida = None
            for c, cell in enumerate(row["tableCells"]):
                if txt(cell).strip() in STATUS_ABERTO:
                    if decidida is None:
                        decidida = (r in status_definido.get(e["objectId"], [])
                                    or _linha_decidida(pid, row["tableCells"], trocas))
                    if not decidida:
                        resto = " | ".join(txt(x).strip() for x in row["tableCells"] if txt(x).strip())
                        print(f"  SEM DECISAO [{e['objectId']} linha {r}]: {resto[:90]!r} -> continua {txt(cell).strip()}")
                        continue
                    loc = {"rowIndex": r, "columnIndex": c}
                    reqs += [{"deleteText": {"objectId": e["objectId"], "cellLocation": loc, "textRange": {"type": "ALL"}}},
                             {"insertText": {"objectId": e["objectId"], "cellLocation": loc, "insertionIndex": 0, "text": "DEFINIDO"}},
                             {"updateTextStyle": {"objectId": e["objectId"], "cellLocation": loc, "textRange": {"type": "ALL"},
                              "style": {"bold": True, "fontFamily": "Arial", "fontSize": {"magnitude": 5.8, "unit": "PT"},
                                        "foregroundColor": GREEN}, "fields": "bold,fontFamily,fontSize,foregroundColor"}}]
    return reqs


def fase1(spec, deck):
    reqs = [rat(p, o, n) for p, o, n in spec.get("trocas", [])]
    reqs += auto_naming(deck, spec["estilo"])
    reqs += status_cells(deck, spec.get("trocas", []), spec.get("status_definido"))
    for oid, rows in spec.get("remover_linhas", {}).items():
        for r in sorted(rows, reverse=True):
            reqs.append({"deleteTableRow": {"tableObjectId": oid, "cellLocation": {"rowIndex": r, "columnIndex": 0}}})
    for sid in spec.get("remover_slides", []):
        reqs.append({"deleteObject": {"objectId": sid}})
    src = spec.get("slide_materiais")
    if src:
        ids = {src: src + "mdf"}
        for s in deck["slides"]:
            if s["objectId"] == src:
                for e in s["pageElements"]:
                    ids[e["objectId"]] = e["objectId"].replace(src + "_", src + "mdf_")
        reqs.append({"duplicateObject": {"objectId": src, "objectIds": ids}})
    return reqs


def fase2(spec, deck):
    reqs, final = [], {}
    by_id = {}
    for pid, e, t in shapes(deck):
        by_id[e["objectId"]] = (pid, t)
    url = lambda fid: "https://drive.google.com/uc?export=view&id=" + fid
    src = spec.get("slide_materiais")
    if src and spec.get("slide_mdf"):
        M = src + "mdf"
        if M not in [s["objectId"] for s in deck["slides"]]:
            sys.exit(f"ERRO: {M} não existe; rode a fase1 antes.")
        md = spec["slide_mdf"]

        def put(oid, new):
            """Troca o texto so desta shape (por objectId): replaceAllText no slide
            atingiria tambem outro card com o mesmo texto antigo. Insere o novo no
            inicio (herda o estilo do texto atual) e apaga o antigo logo depois."""
            pid, old = by_id[oid]
            if old != new:
                reqs.append({"insertText": {"objectId": oid, "insertionIndex": 0, "text": new}})
                if old:
                    ini = _u16(new)
                    reqs.append({"deleteText": {"objectId": oid, "textRange": {
                        "type": "FIXED_RANGE", "startIndex": ini, "endIndex": ini + _u16(old)}}})
            final[oid] = new
        put(f"{M}_i6", md["titulo"])
        if md.get("subtitulo"):
            put(f"{M}_i10", md["subtitulo"])
        cards = md["cards"]
        for k, cd in enumerate(cards):
            b = 12 + 7 * k
            tag = by_id.get(f"{M}_i{b+1}", ("", ""))[1]
            if not re.fullmatch(r"T\d+", tag):
                sys.exit(f"ERRO: layout do card {k+1} inesperado ({M}_i{b+1} = {tag!r}); ajuste à mão.")
            for off, key in zip((1, 2, 3, 4, 5), ("tag", "titulo", "aplicacao", "especificacao", "obra")):
                put(f"{M}_i{b+off}", cd[key])
            reqs.append({"replaceImage": {"imageObjectId": f"{M}_i{b}", "url": url(cd["imagem"]), "imageReplaceMethod": "CENTER_CROP"}})
        on_m = [e["objectId"] for s in deck["slides"] if s["objectId"] == M for e in s["pageElements"]]
        deleted = {o for o in on_m if int(o.rsplit("_i", 1)[1]) >= 11 + 7 * len(cards)}
        reqs += [{"deleteObject": {"objectId": o}} for o in sorted(deleted)]
    else:
        deleted = set()
    for oid, fid in spec.get("trocar_imagens", {}).items():
        reqs.append({"replaceImage": {"imageObjectId": oid, "url": url(fid), "imageReplaceMethod": "CENTER_CROP"}})
    # rodape
    pref = spec["spot"] + "  ·  BOOK DE PREMISSAS"
    for n, s in enumerate(deck["slides"], 1):
        for e in s.get("pageElements", []):
            if "shape" in e:
                t = txt(e["shape"]).rstrip("\n")
                if t.startswith(pref):
                    new = t.rsplit("·", 1)[0] + f"·  {n:02d}"
                    if new != t:
                        reqs.append(rat(s["objectId"], t, new))
    # numeracao de tabelas com coluna '#'
    for pid, e in tables(deck):
        rows = e["table"]["tableRows"]
        if txt(rows[0]["tableCells"][0]).strip() != "#":
            continue
        for r in range(1, len(rows)):
            old, new = txt(rows[r]["tableCells"][0]).rstrip("\n"), f"{r:02d}"
            if old != new and re.fullmatch(r"\d+", old):
                loc = {"rowIndex": r, "columnIndex": 0}
                reqs += [{"insertText": {"objectId": e["objectId"], "cellLocation": loc, "insertionIndex": 0, "text": new}},
                         {"deleteText": {"objectId": e["objectId"], "cellLocation": loc,
                                         "textRange": {"type": "FIXED_RANGE", "startIndex": len(new), "endIndex": len(new) + len(old)}}}]
    # rotulo em negrito, texto normal (replaceAllText herda o negrito do rotulo)
    for pid, e, t in shapes(deck):
        if e["objectId"] in deleted:
            continue
        t = final.get(e["objectId"], t)
        for lab in LABELS:
            if t.startswith(lab):
                reqs += [{"updateTextStyle": {"objectId": e["objectId"], "textRange": {"type": "FIXED_RANGE", "startIndex": 0, "endIndex": len(lab)},
                                              "style": {"bold": True}, "fields": "bold"}},
                         {"updateTextStyle": {"objectId": e["objectId"], "textRange": {"type": "FIXED_RANGE", "startIndex": len(lab), "endIndex": len(t)},
                                              "style": {"bold": False}, "fields": "bold"}}]
    # capa: rotulo cabe em 1 linha; divisor 02: subtitulo desce; slide 2: negrito no trecho do objetivo
    first = deck["slides"][0]["objectId"]
    for pid, e, t in shapes(deck):
        if pid == first and t.startswith("BOOK DE PREMISSAS DE FACHADA E INTERIORES"):
            reqs.append({"updateTextStyle": {"objectId": e["objectId"], "textRange": {"type": "ALL"},
                                             "style": {"fontSize": {"magnitude": 9, "unit": "PT"}}, "fields": "fontSize"}})
        if t == f"Premissas de interiores {a}":
            sub = re.sub(r"_i(\d+)$", lambda m: f"_i{int(m.group(1)) + 1}", e["objectId"])
            for s in deck["slides"]:
                for e2 in s["pageElements"]:
                    if e2["objectId"] == sub:
                        tr = dict(e2["transform"]); tr["translateY"] = tr.get("translateY", 0) + EMU_DESCE_SUBTITULO
                        reqs.append({"updatePageElementTransform": {"objectId": sub, "applyMode": "ABSOLUTE", "transform": tr}})
        seg = f"fachada e de interiores {a}"
        if seg in t and t.startswith("Consolidar"):
            i = t.index(seg)
            reqs.append({"updateTextStyle": {"objectId": e["objectId"], "textRange": {"type": "FIXED_RANGE", "startIndex": i, "endIndex": i + len(seg)},
                                             "style": {"bold": True}, "fields": "bold"}})
    return reqs


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 3:
        print(__doc__); return
    fase, spec = sys.argv[1], json.load(open(sys.argv[2], encoding="utf-8"))
    dry = "--dry" in sys.argv
    deck = composio.get_deck(spec["presentation_id"])
    reqs = fase1(spec, deck) if fase == "fase1" else fase2(spec, deck)
    print(f"{fase}: {len(reqs)} requests")
    if dry:
        for r in reqs:
            if "replaceAllText" in r:
                x = r["replaceAllText"]
                print(f"  [{(x.get('pageObjectIds') or ['*'])[0]}] {x['containsText']['text'][:70]!r} -> {x['replaceText'][:70]!r}")
            elif "insertText" in r and "cellLocation" not in r["insertText"]:
                x = r["insertText"]
                print(f"  [{x['objectId']}] -> {x['text'][:70]!r}")
        return
    replies = composio.batch(spec["presentation_id"], reqs, deck["revisionId"])
    zero = [r["replaceAllText"]["containsText"]["text"][:80] for r, rep in zip(reqs, replies)
            if "replaceAllText" in r and not (rep.get("replaceAllText") or {}).get("occurrencesChanged")]
    print("ok" if not zero else "ATENÇÃO, trocas sem ocorrência (texto não bateu):")
    for z in zero:
        print("  -", z)


if __name__ == "__main__":
    main()
