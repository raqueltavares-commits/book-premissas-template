"""Sobe um .pptx para o Drive convertendo em Google Slides.
  python upload_slides.py arquivo.pptx PASTA_ID "Nome do arquivo"

ATUALIZADO (22/10/2026): o 404 "Tool not found" some passando a versao do toolkit no execute
(VERSION abaixo, a mesma que o MCP devolve em "version"). Assim este script sobe direto, sem o MCP.
ANTIGO (15/10/2026): a chamada direta a GOOGLEDRIVE_RESUMABLE_UPLOAD pela API devolve 404 "Tool not found".
Use este script so ate o stage (imprime o s3key) e chame o MCP `GOOGLEDRIVE_RESUMABLE_UPLOAD` com
file_to_upload={name, s3key, mimetype pptx}, metadata={name, parents:[PASTA], mimeType: application/vnd.google-apps.presentation},
folder_to_upload_to=PASTA, chunkSize=8388608.
ATUALIZAR um Slides existente mantendo o link (testado 15/10/2026): mesmo MCP com file_id=<ID do Slides>,
metadata={name: "<nome final>", mimeType: application/vnd.google-apps.presentation} e file_to_upload.name = nome final
(sem o name, o arquivo passa a se chamar como o pptx). O conteúdo é substituído inteiro.
"""
import hashlib, json, os, sys, urllib.request
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import composio
VERSION = os.environ.get("GDRIVE_VERSION", "20261022_00")
MT = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
f, folder, name = sys.argv[1], sys.argv[2], sys.argv[3]
data = open(f, "rb").read()
r = composio.post("/api/v3/files/upload/request", {"toolkit_slug": "googledrive", "tool_slug": "GOOGLEDRIVE_RESUMABLE_UPLOAD",
                  "filename": os.path.basename(f), "mimetype": MT, "md5": hashlib.md5(data).hexdigest()})
url = r.get("new_presigned_url") or r.get("newPresignedUrl")
if url:
    req = urllib.request.Request(url, data=data, method="PUT"); req.add_header("Content-Type", MT)
    urllib.request.urlopen(req, timeout=600).read()
res = composio.run("GOOGLEDRIVE_RESUMABLE_UPLOAD", {
    "file_to_upload": {"name": os.path.basename(f), "s3key": r["key"], "mimetype": MT},
    "metadata": {"name": name, "parents": [folder], "mimeType": "application/vnd.google-apps.presentation"},
    "folder_to_upload_to": folder, "chunkSize": 262144 * 32}, version=VERSION)
print(json.dumps(res, ensure_ascii=False)[:1500])
