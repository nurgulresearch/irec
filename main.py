from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
import tempfile
import shutil
import validate_irec  # импортируй свой код

app = FastAPI()

# Разрешаем запросы из любой системы (например, Wix)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/check-doc/")
async def check_doc(file: UploadFile = File(...)):
    # Сохраняем временный файл
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    doc = Document(tmp_path)

    # Выполняем все проверки (ты можешь адаптировать)
    part0_results, exemption = validate_irec.validate_part_0(doc)
    part1_results = validate_irec.validate_part_1(doc, exemption)
    part2_results = validate_irec.validate_part_2(doc)

    # Вернуть результаты
    return {
        "part0": part0_results,
        "part1": part1_results,
        "part2": part2_results,
    }
