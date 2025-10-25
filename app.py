from fastapi import FastAPI, Response
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from openpyxl import Workbook
from io import BytesIO
import os

app = FastAPI(title="SheetForge Online")

# Serve /frontend and return index.html at "/"
if not os.path.exists("frontend"):
    os.makedirs("frontend", exist_ok=True)
app.mount("/frontend", StaticFiles(directory="frontend"), name="frontend")

@app.get("/", response_class=HTMLResponse)
async def home():
    with open("frontend/index.html", "r", encoding="utf-8") as f:
        return f.read()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"]
)

class Widget(BaseModel):
    id: int
    type: str
    title: Optional[str] = None
    sub: Optional[str] = None
    x: Optional[int] = 0
    y: Optional[int] = 0
    w: Optional[int] = 0
    h: Optional[int] = 0

class Sheet(BaseModel):
    name: str
    widgets: List[Widget] = []

class SmartSave(BaseModel):
    enabled: bool = False
    baseName: str = "Report_"
    dateFormat: str = "yyyy-mm-dd"
    requireCopyOnOpen: bool = True
    templateTag: str = "TEMPLATE"
    sequenceIfExists: bool = True

class Layout(BaseModel):
    project: str = "Untitled Workbook"
    sheets: List[Sheet]
    settings: SmartSave = SmartSave()

def compile_xlsx(layout: Layout) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "SheetForge Export"
    ws.append(["Project", layout.project])
    ws.append(["SmartSave Enabled", str(layout.settings.enabled)])
    ws.append([])
    for s in layout.sheets:
        ws2 = wb.create_sheet(title=s.name[:31])
        ws2.append(["Title", "Subtitle", "X", "Y", "W", "H", "Type"])
        for w in s.widgets:
            ws2.append([w.title or w.type, w.sub or "", w.x or 0, w.y or 0, w.w or 0, w.h or 0, w.type])
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

@app.post("/api/compile")
async def compile_endpoint(layout: Layout):
    data = compile_xlsx(layout)
    fname = f"{layout.project.replace(' ','_')}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'}
    )
