from fastapi import FastAPI, Path, Response, UploadFile, File
from typing import Optional, Annotated
from Xlsx_Creator import create_xls
from pathlib import Path

UPLOAD_DIR = Path() / 'uploads'

app = FastAPI()


@app.post("/files/")
async def create_file(file: Annotated[bytes, File()]):
    return {"file_size": len(file)}


@app.post("/uploadfile/")
async def create_upload_file(file_upload: UploadFile):
    data = await file_upload.read()
    save_to = UPLOAD_DIR / file_upload.filename
    with open(save_to, 'wb') as f:
        f.write(data)
    return {"filename": file_upload.filename}











"""students = {
    1: {
        "name": "john",
        "age": 17,
        "class": "year 12"}
}

@app.get("/")
def index():
    return {"name": "First Data"}

@app.get("/get-student/{student_id}")
def get_student(student_id: int = Path(description="The ID of the student you wanna view")):
    return students[student_id]

@app.get("/get-by-name")
def get_student(*, name : Optional[str] = None, test : int):
    for student_id in students:
        if students[student_id]["name"] == name:
            return students[student_id]
        return {"Data": "Not found"}"""
