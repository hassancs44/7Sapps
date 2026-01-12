import os
from werkzeug.utils import secure_filename

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
BASE_UPLOAD = os.path.join(BASE_DIR, "uploads")
os.makedirs(BASE_UPLOAD, exist_ok=True)

def upload(file, prefix="REQ"):
    filename = f"{prefix}_{secure_filename(file.filename)}"
    path = os.path.join(BASE_UPLOAD, filename)
    file.save(path)
    return filename
