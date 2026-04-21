from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import pandas as pd
import os
import uuid

app = FastAPI()


@app.get("/")
def home():
    return {"status": "API is running"}


@app.post("/validate")
async def validate(
    file: UploadFile = File(...),
    summary: str = Form("")
):
    try:
        # создаём уникальные имена файлов
        input_filename = f"/tmp/input_{uuid.uuid4()}.xlsx"
        output_filename = f"/tmp/output_{uuid.uuid4()}.xlsx"

        # сохраняем входной файл
        with open(input_filename, "wb") as f:
            f.write(await file.read())

        # читаем Excel
        df = pd.read_excel(input_filename)

        # ===== ТВОЯ ЛОГИКА (пока простая) =====
        df["validated"] = "OK"

        # если хочешь использовать текст из Telegram:
        if summary:
            df["summary"] = summary

        # =====================================

        # сохраняем результат
        df.to_excel(output_filename, index=False)

        return FileResponse(
            output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="result.xlsx"
        )

    except Exception as e:
        return {"error": str(e)}
