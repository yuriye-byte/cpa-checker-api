from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import pandas as pd
import tempfile
import os
import re

app = FastAPI()

# --- простой парсер (потом улучшим) ---
def parse_summary(text):
    rows = []
    parts = re.split(r",\s*", text)

    for part in parts:
        match = re.search(r"([A-Z]{2})-?\$?(\d+)\$\((\d+)\)", part)
        if match:
            geo = match.group(1)
            spend = float(match.group(2))
            ftd = int(match.group(3))
            cpa = spend / ftd if ftd > 0 else 0

            rows.append({
                "geo": geo,
                "manager_ftd": ftd,
                "cpa": cpa,
                "manager_spend": spend
            })
    return pd.DataFrame(rows)

# --- основной эндпоинт ---
@app.post("/validate")
async def validate(
    file: UploadFile = File(...),
    summary: str = Form(...)
):
    # сохраняем файл
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(await file.read())
    temp_input.close()

    # читаем Excel
    df = pd.read_excel(temp_input.name)

    # парсим текст
    summary_df = parse_summary(summary)

    # простая логика (потом усложним)
    result = summary_df.copy()
    result["actual_sum"] = result["manager_ftd"] * result["cpa"]

    # сохраняем результат
    output_path = temp_input.name.replace(".xlsx", "_result.xlsx")
    result.to_excel(output_path, index=False)

    return FileResponse(output_path, filename="result.xlsx")