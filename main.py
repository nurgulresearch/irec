from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
import tempfile
import shutil
import irec  # ← теперь имя файла с проверками

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # позволяет Wix обращаться к API
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/check-doc/")
async def check_doc(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    doc = Document(tmp_path)

    # Базовая проверка — можно добавить больше
    part0_results, exemption = irec.validate_part_0(doc)
    part1_results = irec.validate_part_1(doc, exemption)

    return {
        "part0": part0_results,
        "part1": part1_results
    }
