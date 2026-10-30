"""Amostras de MDF/metal: recorte dos moodboards REV.02, painel comparativo e upload no Drive.

  python swatches.py fetch   PASTA                       # baixa o slide "Acabamentos de mdf" dos 4 moodboards
  python swatches.py compare spec.json PASTA saida.png   # painel lado a lado para a Raquel aprovar
  python swatches.py swatch  ESTILO "NOME" PASTA saida.png   # amostra 680x500 de um MDF do catalogo
  python swatches.py color   R,G,B saida.png             # amostra de cor chapada (metal RAL)
  python swatches.py upload  PASTA_DRIVE arq.png [...]   # sobe, compartilha (leitor) e imprime os IDs

spec.json do compare:
{"titulo": "Batel Spot · Bruma · MDFs do book × catálogo",
 "linhas": [{"hoje": "Frapê", "uso": "liso · marcenaria do ático", "img": "C:/.../frape.jpg",
             "opcoes": [{"estilo": "bruma", "nome": "Nude", "sugestao": true},
                        {"cor": [234,230,202], "nome": "Metal bege RAL 1013"}]}]}
"img" da linha pode faltar (fica "sem amostra"). Opcao: {"estilo","nome"} recorta do moodboard,
{"cor"} gera cor chapada, {"img"} usa arquivo local.
"""
import hashlib
import json
import os
import sys
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import composio  # noqa: E402

MOODBOARDS = {  # estilo: (presentationId, objectId do slide "Acabamentos de mdf")
    "bruma": ("1dEvH_cjlHUFMOzQaRgrh212D5WAW6Q1bfy8hPbTyeGU", "g312d5800cd3_0_62"),
    "biofilico": ("1Ti480HgXwdH3piK3vEZ6qmSSCrS0lvP9NwdqvXtustU", "g3759da45cb1_1_823"),
    "clean": ("1l5nn89FSNder2IFYC835e9aZvVPm6k6WWx9kltbIGjo", "g312d5800cd3_0_62"),
    "industrial": ("1l_S4Yyq5teNGXgxDMwGcwE2dq8g9IpRTGfZy3Fo0bsI", "g312d5800cd3_0_62"),
}
# caixas (x0,y0,x1,y1) na miniatura 1600x900 de cada slide de MDF (medidas em 13/10/2026)
CROPS = {
    "bruma": {"Nude": (53, 116, 240, 427), "Ópera": (250, 117, 442, 431), "Argento": (454, 117, 644, 428),
              "Ruggine": (653, 117, 839, 428), "Nero": (155, 508, 342, 816), "Cacau": (353, 508, 539, 816),
              "Louro Preto": (560, 509, 744, 813)},
    "biofilico": {"Lana": (93, 127, 281, 437), "Cinza Argila": (307, 124, 487, 434), "Desert": (505, 127, 685, 439),
                  "Cinza Cobalto": (703, 128, 882, 438), "Nogal Sevilha": (93, 486, 278, 791),
                  "Carvalho Treviso": (301, 488, 483, 789), "Roble Catedral": (501, 488, 685, 791),
                  "Verti": (707, 488, 891, 795)},
    "clean": {"Cinza Cobalto": (71, 176, 219, 431), "Crômio": (236, 176, 393, 435), "Gelo": (404, 178, 558, 432),
              "Latte": (571, 178, 727, 435), "Desert": (741, 179, 900, 437), "Falésia": (68, 477, 218, 735),
              "Lana": (235, 476, 393, 738), "Carvalho Treviso": (408, 478, 561, 735),
              "Roble Catedral": (578, 480, 729, 733)},
    "industrial": {"Cinza Cristal": (51, 215, 198, 458), "Cinza Cobalto": (210, 214, 350, 458),
                   "Crômio": (361, 214, 509, 458), "Argento": (519, 215, 668, 458),
                   "Linen Grigio": (678, 215, 826, 458), "Chumbo": (832, 214, 980, 457), "Ópera": (50, 507, 198, 750),
                   "Nero": (207, 506, 352, 745), "Solanum": (361, 509, 504, 745), "Parquet": (513, 509, 655, 745),
                   "Carvalho Treviso": (665, 509, 809, 745)},
}


def _pil():
    from PIL import Image, ImageDraw, ImageFont
    return Image, ImageDraw, ImageFont


def crop(estilo, nome, pasta):
    Image, _, _ = _pil()
    im = Image.open(os.path.join(pasta, f"{estilo}_mdf.png")).convert("RGB")
    if im.size != (1600, 900):
        sys.exit(f"ERRO: miniatura {estilo} com tamanho {im.size}; refaça o fetch e confira as caixas.")
    a, b, c, d = CROPS[estilo][nome]
    m = 10
    return im.crop((a + m, b + m, c - m, d - m))


def font(size, bold=False):
    _, _, ImageFont = _pil()
    try:
        return ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf", size)
    except OSError:
        return ImageFont.load_default()


def opt_image(o, pasta):
    Image, _, _ = _pil()
    if "estilo" in o:
        return crop(o["estilo"], o["nome"], pasta)
    if "cor" in o:
        return Image.new("RGB", (200, 200), tuple(o["cor"]))
    return Image.open(o["img"]).convert("RGB")


def compare(spec, pasta, out):
    Image, ImageDraw, _ = _pil()
    T, G = 200, 20

    def tile(im):
        w, h = im.size
        s = min(w, h)
        return im.crop(((w - s) // 2, (h - s) // 2, (w - s) // 2 + s, (h - s) // 2 + s)).resize((T, T))

    linhas = spec["linhas"]
    maxo = max(len(l["opcoes"]) for l in linhas)
    W, RH = 40 + T + 70 + maxo * (T + G) + 40, T + 110
    img = Image.new("RGB", (W, 130 + RH * len(linhas)), "white")
    dr = ImageDraw.Draw(img)
    dr.text((40, 22), spec["titulo"], font=font(30, True), fill="#111")
    dr.text((40, 64), "Borda preta = hoje no book · borda verde = sugestão · demais = outras opções do catálogo",
            font=font(18), fill="#555")
    for i, l in enumerate(linhas):
        y = 115 + i * RH
        dr.text((40, y), f"{l['hoje']}  —  {l.get('uso', '')}", font=font(22, True), fill="#111")
        y += 34
        x = 40
        if l.get("img"):
            img.paste(tile(Image.open(l["img"]).convert("RGB")), (x, y))
        else:
            dr.rectangle((x, y, x + T, y + T), fill="#ebebeb")
            dr.text((x + 14, y + 80), "sem amostra\n(só render)", font=font(20, True), fill="#666")
        dr.rectangle((x - 3, y - 3, x + T + 2, y + T + 2), outline="#111", width=3)
        dr.text((x, y + T + 6), "HOJE: " + l["hoje"], font=font(16, True), fill="#111")
        x += T + 70
        dr.text((x - 55, y + T // 2 - 30), "→", font=font(48, True), fill="#0a7")
        for o in l["opcoes"]:
            img.paste(tile(opt_image(o, pasta)), (x, y))
            dr.rectangle((x - 1, y - 1, x + T, y + T), outline="#ccc", width=1)
            if o.get("sugestao"):
                dr.rectangle((x - 3, y - 3, x + T + 2, y + T + 2), outline="#0a7", width=5)
            dr.text((x, y + T + 6), o["nome"], font=font(16, True), fill="#0a7" if o.get("sugestao") else "#333")
            x += T + G
    img.save(out)
    print(out, img.size)


def to_swatch(im, out):
    w, h = im.size
    tw = h * 1.36
    if tw <= w:
        im = im.crop((int((w - tw) / 2), 0, int((w + tw) / 2), h))
    else:
        th = w / 1.36
        im = im.crop((0, int((h - th) / 2), w, int((h + th) / 2)))
    im.resize((680, 500)).save(out)
    print(out)


def upload(folder, files):
    for f in files:
        data = open(f, "rb").read()
        name = os.path.basename(f)
        r = composio.post("/api/v3/files/upload/request",
                          {"toolkit_slug": "googledrive", "tool_slug": "GOOGLEDRIVE_UPLOAD_FILE", "filename": name,
                           "mimetype": "image/png", "md5": hashlib.md5(data).hexdigest()})
        url = r.get("new_presigned_url") or r.get("newPresignedUrl")
        if url:
            req = urllib.request.Request(url, data=data, method="PUT")
            req.add_header("Content-Type", "image/png")
            urllib.request.urlopen(req, timeout=120).read()
        res = composio.run("GOOGLEDRIVE_UPLOAD_FILE", {"file_to_upload": {"name": name, "s3key": r["key"], "mimetype": "image/png"},
                                                       "folder_to_upload_to": folder})
        fid = (res.get("data") or {}).get("id")
        if not fid:
            print(name, "ERRO", str(res)[:300]); continue
        sh = composio.run("GOOGLEDRIVE_ADD_FILE_SHARING_PREFERENCE", {"file_id": fid, "role": "reader", "type": "anyone"})
        if not sh.get("successful", True) or sh.get("error"):
            # without public read the Slides batch can't fetch the image: don't hand out the ID
            print(name, "ERRO compartilhamento", fid, str(sh.get("error") or sh)[:300]); continue
        print(name, fid)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__); return
    cmd = sys.argv[1]
    if cmd == "fetch":
        os.makedirs(sys.argv[2], exist_ok=True)
        for estilo, (pid, page) in MOODBOARDS.items():
            r = composio.run("GOOGLESLIDES_PRESENTATIONS_PAGES_GET_THUMBNAIL", {"presentationId": pid, "pageObjectId": page})
            d = r.get("data") or {}
            url = d.get("contentUrl") or (d.get("response_data") or {}).get("contentUrl")
            path = os.path.join(sys.argv[2], f"{estilo}_mdf.png")
            open(path, "wb").write(urllib.request.urlopen(url).read())
            print(path)
    elif cmd == "compare":
        compare(json.load(open(sys.argv[2], encoding="utf-8")), sys.argv[3], sys.argv[4])
    elif cmd == "swatch":
        to_swatch(crop(sys.argv[2], sys.argv[3], sys.argv[4]), sys.argv[5])
    elif cmd == "color":
        Image, _, _ = _pil()
        to_swatch(Image.new("RGB", (680, 500), tuple(int(v) for v in sys.argv[2].split(","))), sys.argv[3])
    elif cmd == "upload":
        upload(sys.argv[2], sys.argv[3:])
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
