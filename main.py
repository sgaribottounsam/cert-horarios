from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os
from nota_horarios_cursada_gpt import generar_certificado

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def show_form(request: Request):
    return templates.TemplateResponse("formulario.html", {"request": request})

@app.post("/generar", response_class=FileResponse)
async def generar_certificado_form(
    request: Request,
    dni: str = Form(...),
    sello: str = Form(None)
):
    incluir_sello = sello == "true"
    generar_certificado(str(dni), incluir_sello=incluir_sello)
    archivo = f"Certificado_{dni}.pdf"
    return FileResponse(archivo, media_type="application/pdf", filename=archivo)

