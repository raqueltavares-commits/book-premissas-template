"""Baixa arquivos do Drive (via Composio) e reduz imagens a 2400 px.

  python dl.py PASTA_SAIDA ID[=nome] [ID[=nome] ...]
  python dl.py PASTA_SAIDA @lista.txt          # uma linha por arquivo: ID=nome (ou so ID)

Imagens (.jpg/.png/.jpeg/.webp) sao salvas como JPEG com lado maior <= 2400 px.
"""
import io
import os
import sys
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import composio  # noqa: E402

MAX = 2400


def fetch(fid):
    r = composio.run("GOOGLEDRIVE_DOWNLOAD_FILE", {"file_id": fid})
    d = r.get("data", r)
    c = d.get("downloaded_file_content") or d.get("file") or {}
    url = c.get("s3url") or c.get("url")
    if not url:
        raise RuntimeError(f"sem url para {fid}: {str(r)[:300]}")
    return c.get("name") or fid, urllib.request.urlopen(url, timeout=600).read()


_IMG_EXTS = (".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp")


def safe_path(out_dir, name):
    """Keep only the file name and refuse anything that would land outside out_dir."""
    base = os.path.basename(name.replace("\\", "/"))
    if base in ("", ".", ".."):
        raise ValueError(f"nome de arquivo invalido: {name!r}")
    root = os.path.realpath(out_dir)
    path = os.path.realpath(os.path.join(root, base))
    if os.path.dirname(path) != root:
        raise ValueError(f"caminho fora da pasta de saida: {name!r}")
    return path


def final_path(path, name):
    """Path the file will actually have on disk (images become .jpg)."""
    if os.path.splitext(name)[1].lower() in _IMG_EXTS:
        return os.path.splitext(path)[0] + ".jpg"
    return path


def save(out_dir, name, data, fid=None, taken=None):
    """Save `data` as `name` inside out_dir. Never overwrites: if the final path already
    exists (this run or a previous one), add the suffix _<fid>. Returns None when that
    suffixed file also exists (same file already downloaded)."""
    ext = os.path.splitext(name)[1].lower()
    path = safe_path(out_dir, name)
    taken = taken if taken is not None else set()
    fp = final_path(path, name)
    if fp in taken or os.path.exists(fp):
        if not fid:
            return None
        stem, e = os.path.splitext(os.path.basename(path))
        path = safe_path(out_dir, f"{stem}_{fid}{e}")
        if os.path.exists(final_path(path, name)):
            return None
    if ext in _IMG_EXTS:
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        im = im.convert("RGB")
        w, h = im.size
        s = MAX / max(w, h)
        if s < 1:
            im = im.resize((int(w * s), int(h * s)), Image.LANCZOS)
        path = os.path.splitext(path)[0] + ".jpg"
        im.save(path, "JPEG", quality=90)
    else:
        with open(path, "wb") as f:
            f.write(data)
    taken.add(path)
    return path


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    items = []
    for a in sys.argv[2:]:
        if a.startswith("@"):
            items += [ln.strip() for ln in open(a[1:], encoding="utf-8") if ln.strip() and not ln.startswith("#")]
        else:
            items.append(a)
    import time
    taken = set()  # final paths written in this run (avoid silent overwrite)
    for it in items:
        fid, _, name = it.partition("=")
        if name:
            try:
                p = safe_path(out_dir, name)
            except ValueError as e:
                print("ERRO", fid, e)
                continue
            # same final name save() would use (foto.png -> foto.jpg; "foto" -> foto.jpg)
            if os.path.exists(final_path(p, name)) or os.path.exists(p + ".jpg") or os.path.exists(p):
                print("ja existe", name)
                continue
        data = None
        for tent in range(6):
            try:
                orig, data = fetch(fid)
                break
            except Exception as e:  # noqa: BLE001
                msg = str(e)
                if "Quota" in msg or "quota" in msg or "429" in msg:
                    print("quota; aguardando 30 s", fid)
                    time.sleep(30)
                    continue
                print("ERRO", fid, msg[:200])
                break
        if data is None:
            print("ERRO (desistiu)", fid)
            continue
        if name:
            ext = os.path.splitext(orig)[1] or ".jpg"
            name = name if os.path.splitext(name)[1] else name + ext
        else:
            name = orig
        try:
            saved = save(out_dir, name, data, fid=fid, taken=taken)
            print(saved if saved else f"ja existe {name}", len(data))
        except ValueError as e:
            print("ERRO", fid, e)


if __name__ == "__main__":
    main()
