from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
import os
import shutil
import tempfile
import traceback

from checker_core import process_file

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
        with tempfile.TemporaryDirectory() as tmp:
            input_path = os.path.join(tmp, "input.xlsx")
            output_path = os.path.join(tmp, "result_deposits.xlsx")

            with open(input_path, "wb") as f:
                shutil.copyfileobj(file.file, f)

            result = process_file(input_path, summary, output_path)

            if not os.path.exists(output_path):
                return JSONResponse(
                    status_code=500,
                    content={
                        "error": "Output file was not created",
                        "result": result,
                    },
                )

            return FileResponse(
                output_path,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="result_deposits.xlsx"
            )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc()
            }
        )