"""Stage de pptx no storage do Composio para o MCP GOOGLEDRIVE_RESUMABLE_UPLOAD.
  python stage.py arq1.pptx [arq2.pptx ...]   -> imprime "arquivo s3key" por linha
"""
import hashlib, os, sys, urllib.request
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import composio
MT = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
for f in sys.argv[1:]:
    data = open(f, "rb").read()
    r = composio.post("/api/v3/files/upload/request", {"toolkit_slug": "googledrive", "tool_slug": "GOOGLEDRIVE_RESUMABLE_UPLOAD",
        "filename": os.path.basename(f), "mimetype": MT, "md5": hashlib.md5(data).hexdigest()})
    url = r.get("new_presigned_url") or r.get("newPresignedUrl")
    if url:
        req = urllib.request.Request(url, data=data, method="PUT"); req.add_header("Content-Type", MT)
        urllib.request.urlopen(req, timeout=600).read()
    print(os.path.basename(f), r["key"])
