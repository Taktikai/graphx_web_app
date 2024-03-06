from fastapi import FastAPI, UploadFile, Request
from fastapi.responses import FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from Xlsx_Creator import create_xls
from pathlib import Path
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

templates = Jinja2Templates(directory="templates")



@app.get("/")
def dashboard(request: Request):
    return templates.TemplateResponse("dashboard.html", {"request": request})



@app.post("/uploadfile/")
async def create_upload_file(file_upload: UploadFile):
    with open(file_upload.filename, "wb") as buffer:
        buffer.write(await file_upload.read())

    xls_file_path = create_xls(file_upload.filename)

    os.remove(file_upload.filename)

    return {"filename": xls_file_path}



@app.get("/downloadfile/")
async def download_file(filename: str):
    if not Path(filename).is_file():
        return {"error": "File not found"}
    return FileResponse(filename, filename=os.path.basename(filename))