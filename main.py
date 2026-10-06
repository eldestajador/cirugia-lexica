import io
import os
import re
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from PIL import Image
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from google import genai
from google.genai import types
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable
)
from reportlab.lib.units import cm

app = FastAPI(title="Cirugía Léxica ELE")

API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    print("ADVERTENCIA: La variable de entorno GEMINI_API_KEY no está configurada.")

cliente_gemini = genai.Client(api_key=API_KEY)


def limpiar_markdown_residual(texto: str) -> str:
    linea = texto.strip()
    if linea.startswith("```"):
        return ""
    return linea


def fabricar_word(texto_ia: str, contexto_destino: str) -> io.BytesIO:
    doc = Document()

    for seccion in doc.sections:
        seccion.top_margin = Inches(0.8)
        seccion.bottom_margin = Inches(0.8)
        seccion.left_margin = Inches(0.9)
        seccion.right_margin = Inches(0.9)

    tag = "ACTIVIDAD DIDÁCTICA ADAPTADA • CONTEXTO: " + contexto_destino.upper()
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(4)
    run_meta = p_meta.add_run(tag)
    run_meta.font.name = "Arial"
    run_meta.font.size = Pt(8.5)
    run_meta.font.bold = True
    run_meta.font.color.rgb = RGBColor(120, 120, 120)

    p_alumno = doc.add_paragraph()
    p_alumno.paragraph_format.space_after = Pt(16)
    subraya = "_" * 30
    run_alumno = p_alumno.add_run(f"Nombre del estudiante: {subraya}   Fecha: ___________")
    run_alumno.font.name = "Arial"
    run_alumno.font.size = Pt(9.5)
    run_alumno.font.color.rgb = RGBColor(90, 90, 90)

    texto_reparado = re.sub(r"\n\s*\|\s*", " | ", texto_ia)
    lineas = [l.strip() for l in texto_reparado.split("\n") if l.strip()]
    
    i = 0
    total_lineas = len(lineas)

    while i < total_lineas:
        linea_raw = lineas[i]
        linea_str = limpiar_markdown_residual(linea_raw)

        if not linea_str:
            i += 1
            continue

        if "|" in linea_raw and not linea_raw.startswith(("#", "Respuestas:", "Respuesta:")):
            filas_tabla = []
            while i < total_lineas and "|" in lineas[i] and not lineas[i].startswith(("#", "Respuestas:", "Respuesta:")):
                fila_actual = lineas[i]
                if not re.match(r"^\|?(\s*:?-+:?\s*\|?)+$", fila_actual):
                    celdas = [c.strip() for c in fila_actual.strip("|").split("|") if c.strip()]
                    if len(celdas) >= 2:
                        filas_tabla.append(celdas[:2])
                i += 1

            if filas_tabla:
                tabla = doc.add_table(rows=len(filas_tabla), cols=2)
                tabla.style = "Table Grid"

                for num_f, fila in enumerate(filas_tabla):
                    for num_c, texto_celda in enumerate(fila):
                        celda = tabla.cell(num_f, num_c)
                        celda.text = ""
                        p_celda = celda.paragraphs[0]
                        p_celda.paragraph_format.space_before = Pt(3)
                        p_celda.paragraph_format.space_after = Pt(3)

                        partes = re.split(r"(\*\*.*?\*\*)", texto_celda)
                        for parte in partes:
                            if parte.startswith("**") and parte.endswith("**"):
                                r = p_celda.add_run(parte[2:-2])
                                r.font.bold = True
                            else:
                                r = p_celda.add_run(re.sub(r"[*_]", "", parte))
                            r.font.name = "Arial"
                            r.font.size = Pt(9.5)
                            if num_f == 0:
                                r.font.bold = True

                p_espacio = doc.add_paragraph()
                p_espacio.paragraph_format.space_before = Pt(2)
                p_espacio.paragraph_format.space_after = Pt(4)
            continue

        if linea_str in ("---", "***", "___"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            i += 1
            continue

        if linea_str.lower().startswith(("foto:", "pie de foto", "imagen:")):
            i += 1
            continue

        if re.match(r"^#\s+", linea_raw):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
            texto_titulo = re.sub(r"^#+\s*", "", linea_str).strip()
            run = p.add_run(re.sub(r"[*_]", "", texto_titulo))
            run.font.name = "Arial"
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(24, 43, 73)

        elif re.match(r"^##\s+", linea_raw):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            texto_subtitulo = re.sub(r"^#+\s*", "", linea_str).strip()
            run = p.add_run(re.sub(r"[*_]", "", texto_subtitulo))
            run.font.name = "Arial"
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(41, 128, 185)

        elif re.match(r"^###\s+", linea_raw):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
            texto_h3 = re.sub(r"^#+\s*", "", linea_str).strip()
            run = p.add_run(re.sub(r"[*_]", "", texto_h3))
            run.font.name = "Arial"
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.color.rgb = RGBColor(30, 80, 120)

        elif re.search(r"Actividad\s+\d+\.\d+", linea_raw, re.IGNORECASE):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            texto_limpio = re.sub(r"^[#\s*]+", "", linea_str).strip()
            run = p.add_run(re.sub(r"[*_]", "", texto_limpio))
            run.font.name = "Arial"
            run.font.size = Pt(10.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(20, 50, 90)

        elif re.match(r"^\d+[\.\)]\s+", linea_str):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.left_indent = Inches(0.2)
            run = p.add_run(re.sub(r"[*_]", "", linea_str))
            run.font.name = "Arial"
            run.font.size = Pt(10)

        elif re.match(r"^[a-dA-D][\.\)]\s+", linea_str) or linea_raw.startswith(("- ", "• ", "* ", "– ")):
            contenido_inciso = re.sub(r"^(\*|-|•|–)\s+", "", linea_str)
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Inches(0.35)
            p.add_run("• ").font.name = "Arial"

            partes = re.split(r"(\*\*.*?\*\*)", contenido_inciso)
            for parte in partes:
                if parte.startswith("**") and parte.endswith("**"):
                    run = p.add_run(parte[2:-2])
                    run.font.bold = True
                else:
                    run = p.add_run(re.sub(r"[*_]", "", parte))
                run.font.name = "Arial"
                run.font.size = Pt(10)

        else:
            if len(linea_str.split()) == 1 and len(linea_str) < 18 and not linea_str.endswith((".", ":", "?", "!", "]")):
                i += 1
                continue

            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(3)
            
            partes = re.split(r"(\*\*.*?\*\*)", linea_str)
            for parte in partes:
                if parte.startswith("**") and parte.endswith("**"):
                    run = p.add_run(parte[2:-2])
                    run.font.bold = True
                else:
                    run = p.add_run(re.sub(r"[*_]", "", parte))
                run.font.name = "Arial"
                run.font.size = Pt(10.5)

        i += 1

    buffer_doc = io.BytesIO()
    doc.save(buffer_doc)
    buffer_doc.seek(0)
    return buffer_doc


def fabricar_pdf(texto_ia: str, contexto_destino: str) -> io.BytesIO:
    buffer_pdf = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer_pdf,
        pagesize=letter,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm
    )

    styles = getSampleStyleSheet()

    estilo_meta = ParagraphStyle(
        'MetaHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#787878'),
        spaceAfter=4
    )

    estilo_alumno = ParagraphStyle(
        'AlumnoHeader',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#5A5A5A'),
        spaceAfter=14
    )

    estilo_h1 = ParagraphStyle(
        'H1Custom',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#182B49'),
        spaceBefore=12,
        spaceAfter=6
    )

    estilo_h2 = ParagraphStyle(
        'H2Custom',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor('#2980B9'),
        spaceBefore=10,
        spaceAfter=4
    )

    estilo_h3 = ParagraphStyle(
        'H3Custom',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.HexColor('#1E5078'),
        spaceBefore=8,
        spaceAfter=3
    )

    estilo_cuerpo = ParagraphStyle(
        'BodyCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#2A2A2A'),
        spaceAfter=4
    )

    estilo_bullet = ParagraphStyle(
        'BulletCustom',
        parent=estilo_cuerpo,
        leftIndent=14,
        spaceAfter=3
    )

    estilo_celda = ParagraphStyle(
        'CellCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#333333')
    )

    historia = []

    tag = f"ACTIVIDAD DIDÁCTICA ADAPTADA • CONTEXTO: {contexto_destino.upper()}"
    historia.append(Paragraph(tag, estilo_meta))
    subraya = "_" * 28
    historia.append(Paragraph(f"Nombre del estudiante: {subraya}   Fecha: ___________", estilo_alumno))
    historia.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor('#D0D7DE'), spaceAfter=10))

    texto_reparado = re.sub(r"\n\s*\|\s*", " | ", texto_ia)
    lineas = [l.strip() for l in texto_reparado.split("\n") if l.strip()]

    i = 0
    total = len(lineas)

    while i < total:
        linea_raw = lineas[i]
        linea_str = limpiar_markdown_residual(linea_raw)

        if not linea_str or linea_str.lower().startswith(("foto:", "pie de foto", "imagen:")):
            i += 1
            continue

        if linea_str in ("---", "***", "___"):
            historia.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#E2E8F0'), spaceAfter=6, spaceBefore=6))
            i += 1
            continue

        if "|" in linea_raw and not linea_raw.startswith(("#", "Respuestas:", "Respuesta:")):
            filas_tabla = []
            while i < total and "|" in lineas[i] and not lineas[i].startswith(("#", "Respuestas:", "Respuesta:")):
                fila_actual = lineas[i]
                if not re.match(r"^\|?(\s*:?-+:?\s*\|?)+$", fila_actual):
                    celdas = [c.strip() for c in fila_actual.strip("|").split("|") if c.strip()]
                    if len(celdas) >= 2:
                        filas_tabla.append([
                            Paragraph(re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", celdas[0]), estilo_celda),
                            Paragraph(re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", celdas[1]), estilo_celda)
                        ])
                i += 1

            if filas_tabla:
                t = Table(filas_tabla, colWidths=[5 * cm, 12 * cm])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
                    ('BOX', (0, 0), (-1, -1), 0.8, colors.HexColor('#CBD5E1')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ]))
                historia.append(t)
                historia.append(Spacer(1, 0.3 * cm))
            continue

        if re.match(r"^#\s+", linea_raw):
            texto = re.sub(r"^#+\s*", "", linea_str).strip()
            historia.append(Paragraph(texto, estilo_h1))
        elif re.match(r"^##\s+", linea_raw):
            texto = re.sub(r"^#+\s*", "", linea_str).strip()
            historia.append(Paragraph(texto, estilo_h2))
        elif re.match(r"^###\s+", linea_raw):
            texto = re.sub(r"^#+\s*", "", linea_str).strip()
            historia.append(Paragraph(texto, estilo_h3))
        elif linea_raw.startswith(("- ", "• ", "* ", "– ")):
            contenido = re.sub(r"^(\*|-|•|–)\s+", "", linea_str)
            contenido_html = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", contenido)
            historia.append(Paragraph(f"• {contenido_html}", estilo_bullet))
        else:
            contenido_html = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", linea_str)
            historia.append(Paragraph(contenido_html, estilo_cuerpo))

        i += 1

    doc.build(historia)
    buffer_pdf.seek(0)
    return buffer_pdf


@app.get("/", response_class=HTMLResponse)
async def interfaz_web():
    return """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Cirugía Léxica | IA Didáctica para Docentes ELE</title>
        <link rel="preconnect" href="[https://fonts.googleapis.com](https://fonts.googleapis.com)">
        <link rel="preconnect" href="[https://fonts.gstatic.com](https://fonts.gstatic.com)" crossorigin>
        <link href="[https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap](https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap)" rel="stylesheet">
        <style>
            :root {
                --brand-green: #00e676;
                --brand-green-hover: #00c866;
                --brand-green-light: #e8fdf2;
                --brand-green-glow: rgba(0, 230, 118, 0.25);
                --dark-navy: #091a24;
                --dark-card: #0f2330;
                --bg-base: #f9fbfb;
                --surface-card: #ffffff;
                --text-main: #0a1924;
                --text-muted: #536b78;
                --border-color: #e2ece9;
                --radius-pill: 9999px;
            }

            * { box-sizing: border-box; margin: 0; padding: 0; }
            
            body {
                font-family: 'Plus Jakarta Sans', sans-serif;
                background-color: var(--bg-base);
                color: var(--text-main);
                min-height: 100vh;
                -webkit-font-smoothing: antialiased;
            }

            .top-banner {
                background-color: var(--dark-navy);
                color: #f1f7f5;
                padding: 10px 16px;
                text-align: center;
                font-size: 13px;
                font-weight: 500;
                display: flex;
                justify-content: center;
                align-items: center;
                position: relative;
                border-bottom: 1px solid rgba(0, 230, 118, 0.2);
            }

            .top-banner span.highlight {
                color: var(--brand-green);
                font-weight: 700;
                margin-right: 6px;
            }

            .banner-close {
                position: absolute;
                right: 18px;
                cursor: pointer;
                font-size: 14px;
                color: #8fa6b2;
            }

            .nav-container {
                max-width: 1200px;
                margin: 16px auto 0 auto;
                padding: 0 16px;
            }

            .floating-nav {
                background: #ffffff;
                border-radius: 18px;
                padding: 10px 24px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                box-shadow: 0 6px 24px rgba(9, 26, 36, 0.04);
                border: 1.5px solid var(--border-color);
            }

            .nav-brand {
                display: flex;
                align-items: center;
                gap: 10px;
                font-size: 18px;
                font-weight: 800;
                letter-spacing: -0.5px;
                color: var(--dark-navy);
                text-transform: uppercase;
            }

            .brand-badge {
                width: 26px;
                height: 26px;
                background: var(--brand-green);
                border-radius: 8px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 8px var(--brand-green-glow);
                color: var(--dark-navy);
                font-size: 14px;
                font-weight: 900;
            }

            .nav-links {
                display: flex;
                gap: 26px;
                list-style: none;
                font-size: 13.5px;
                font-weight: 600;
                color: var(--text-muted);
            }

            .nav-links li:hover { color: var(--dark-navy); cursor: pointer; }

            .nav-actions {
                display: flex;
                align-items: center;
                gap: 12px;
            }

            .credits-chip {
                background: var(--brand-green-light);
                color: #008744;
                border: 1.5px solid #a3f7cb;
                font-size: 12px;
                font-weight: 800;
                padding: 6px 14px;
                border-radius: var(--radius-pill);
                display: flex;
                align-items: center;
                gap: 6px;
            }

            .hero-wrapper {
                max-width: 1200px;
                margin: 40px auto;
                padding: 0 20px;
                display: grid;
                grid-template-columns: 1.15fr 1fr;
                gap: 44px;
                align-items: start;
            }

            .trust-caption {
                font-size: 13.5px;
                font-weight: 800;
                color: #00a852;
                text-transform: uppercase;
                letter-spacing: 0.8px;
                margin-bottom: 12px;
                display: flex;
                align-items: center;
                gap: 6px;
            }

            .trust-caption::before {
                content: '';
                display: inline-block;
                width: 8px;
                height: 8px;
                background: var(--brand-green);
                border-radius: 50%;
                box-shadow: 0 0 8px var(--brand-green);
            }

            .hero-title {
                font-family: 'Barlow Condensed', sans-serif;
                font-size: 68px;
                line-height: 0.93;
                font-weight: 900;
                text-transform: uppercase;
                color: var(--dark-navy);
                letter-spacing: -0.5px;
                margin-bottom: 22px;
            }

            .hero-description {
                font-size: 16.5px;
                line-height: 1.5;
                color: var(--text-muted);
                max-width: 480px;
                margin-bottom: 30px;
            }

            .value-grid {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 14px;
                max-width: 490px;
            }

            .value-card {
                background: #ffffff;
                border: 1.5px solid var(--border-color);
                border-radius: 14px;
                padding: 16px;
                box-shadow: 0 2px 8px rgba(9, 26, 36, 0.02);
                transition: all 0.2s ease;
            }

            .value-card:hover {
                transform: translateY(-2px);
                border-color: var(--brand-green);
                box-shadow: 0 8px 20px rgba(0, 230, 118, 0.12);
            }

            .v-icon-wrap {
                width: 38px;
                height: 38px;
                border-radius: 10px;
                background-color: var(--brand-green-light);
                display: flex;
                align-items: center;
                justify-content: center;
                margin-bottom: 10px;
            }

            .v-svg {
                width: 20px;
                height: 20px;
                stroke: #00a852;
            }

            .v-title {
                font-size: 13.5px;
                font-weight: 700;
                color: var(--dark-navy);
                margin-bottom: 4px;
            }

            .v-desc {
                font-size: 12px;
                color: var(--text-muted);
                line-height: 1.45;
            }

            .tool-card {
                background: #ffffff;
                border: 1.5px solid var(--border-color);
                border-radius: 24px;
                padding: 28px;
                box-shadow: 0 16px 40px rgba(9, 26, 36, 0.05);
            }

            .tool-header {
                font-size: 19px;
                font-weight: 800;
                color: var(--dark-navy);
                margin-bottom: 18px;
                letter-spacing: -0.3px;
            }

            .form-field { margin-bottom: 16px; }

            label.field-caption {
                display: block;
                font-size: 12px;
                font-weight: 800;
                text-transform: uppercase;
                letter-spacing: 0.6px;
                color: var(--text-muted);
                margin-bottom: 6px;
            }

            input[type="text"], select {
                width: 100%;
                padding: 12px 14px;
                font-size: 14px;
                border: 1.5px solid var(--border-color);
                border-radius: 12px;
                background: #fff;
                color: var(--text-main);
                font-family: inherit;
                transition: border-color 0.15s, box-shadow 0.15s;
            }

            input[type="text"]:focus, select:focus {
                outline: none;
                border-color: var(--brand-green);
                box-shadow: 0 0 0 3px var(--brand-green-glow);
            }

            .file-dropzone {
                background: #ffffff;
                border: 1.5px solid var(--border-color);
                border-radius: 14px;
                padding: 16px 20px;
                display: flex;
                align-items: center;
                gap: 16px;
                cursor: pointer;
                transition: all 0.2s ease;
                box-shadow: 0 1px 3px rgba(9, 26, 36, 0.02);
            }

            .file-dropzone:hover {
                border-color: var(--brand-green);
                background: #f7fdfa;
                transform: translateY(-1px);
            }

            .file-dropzone input { display: none; }

            .dropzone-icon-box {
                width: 44px;
                height: 44px;
                border-radius: 12px;
                background-color: var(--brand-green-light);
                display: flex;
                align-items: center;
                justify-content: center;
                flex-shrink: 0;
            }

            .dropzone-icon-box svg {
                width: 22px;
                height: 22px;
                stroke: #00a852;
            }

            .dropzone-text-group {
                display: flex;
                flex-direction: column;
                gap: 3px;
                text-align: left;
            }

            .dropzone-btn-action {
                font-size: 13.5px;
                font-weight: 700;
                color: var(--dark-navy);
            }

            .dropzone-sub {
                font-size: 12px;
                color: var(--text-muted);
            }

            .grid-selectors {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 12px;
            }

            .addons-box {
                background: #f7fbf9;
                border: 1.5px solid var(--border-color);
                border-radius: 14px;
                padding: 14px;
                margin-bottom: 20px;
            }

            .check-card {
                display: flex;
                align-items: center;
                gap: 10px;
                background: #ffffff;
                border: 1.5px solid var(--border-color);
                padding: 9px 12px;
                border-radius: 10px;
                cursor: pointer;
                transition: border-color 0.15s;
            }

            .check-card:hover { border-color: #b8d9cc; }

            .check-card input[type="checkbox"] {
                width: 17px;
                height: 17px;
                accent-color: #00c866;
            }

            .check-card b {
                font-size: 13px;
                color: var(--dark-navy);
                display: block;
            }

            .check-card span {
                font-size: 11.5px;
                color: var(--text-muted);
            }

            /* --- ESTILOS SLIDER TIPO QUILLBOT --- */
            .slider-wrapper {
                background: #ffffff;
                border: 1.5px solid var(--border-color);
                border-radius: 12px;
                padding: 14px 16px;
                position: relative;
                margin-top: 10px;
            }

            .slider-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 12px;
            }

            .slider-title {
                font-size: 13px;
                font-weight: 700;
                color: var(--dark-navy);
            }

            .slider-badge {
                font-size: 11px;
                font-weight: 800;
                background: var(--brand-green-light);
                color: #008744;
                padding: 3px 9px;
                border-radius: 20px;
                border: 1px solid #a3f7cb;
            }

            .slider-track-container {
                position: relative;
                padding: 6px 0;
            }

            .qb-slider {
                -webkit-appearance: none;
                width: 100%;
                height: 6px;
                border-radius: 4px;
                background: #e2ece9;
                outline: none;
                cursor: pointer;
                margin: 0;
            }

            .qb-slider::-webkit-slider-thumb {
                -webkit-appearance: none;
                appearance: none;
                width: 20px;
                height: 20px;
                border-radius: 50%;
                background: #ffffff;
                border: 3px solid var(--brand-green);
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15);
                cursor: pointer;
                transition: transform 0.1s ease;
            }

            .qb-slider::-webkit-slider-thumb:hover {
                transform: scale(1.2);
            }

            .slider-ticks {
                display: flex;
                justify-content: space-between;
                margin-top: 8px;
                font-size: 10.5px;
                font-weight: 600;
                color: var(--text-muted);
            }

            .qb-tooltip {
                position: absolute;
                bottom: calc(100% + 8px);
                left: 0%;
                transform: translateX(-50%);
                background: var(--dark-navy);
                color: #ffffff;
                padding: 6px 10px;
                border-radius: 8px;
                font-size: 11.5px;
                font-weight: 500;
                white-space: nowrap;
                pointer-events: none;
                opacity: 0;
                transition: opacity 0.2s;
                box-shadow: 0 4px 12px rgba(9, 26, 36, 0.2);
                z-index: 10;
            }

            .qb-tooltip::after {
                content: '';
                position: absolute;
                top: 100%;
                left: 50%;
                transform: translateX(-50%);
                border-width: 5px;
                border-style: solid;
                border-color: var(--dark-navy) transparent transparent transparent;
            }

            .slider-wrapper:hover .qb-tooltip {
                opacity: 1;
            }

            .btn-cta {
                width: 100%;
                background: var(--brand-green);
                color: var(--dark-navy);
                border: none;
                padding: 15px;
                border-radius: var(--radius-pill);
                font-size: 15.5px;
                font-weight: 800;
                cursor: pointer;
                font-family: inherit;
                box-shadow: 0 6px 20px var(--brand-green-glow);
                transition: transform 0.15s, background 0.15s, box-shadow 0.15s;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 8px;
            }

            .btn-cta:hover {
                background: var(--brand-green-hover);
                transform: translateY(-2px);
                box-shadow: 0 8px 24px rgba(0, 230, 118, 0.4);
            }

            @media (max-width: 900px) {
                .hero-wrapper {
                    grid-template-columns: 1fr;
                    margin: 18px auto;
                    gap: 24px;
                }

                .hero-title { font-size: 44px; }
                .nav-links { display: none; }
                .floating-nav { padding: 8px 16px; }
                .value-grid { grid-template-columns: 1fr; }
                .grid-selectors { grid-template-columns: 1fr; gap: 12px; }
            }

            .loader-screen {
                display: none;
                position: fixed;
                top: 0; left: 0; right: 0; bottom: 0;
                background: rgba(9, 26, 36, 0.88);
                backdrop-filter: blur(5px);
                z-index: 1000;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                color: #ffffff;
                text-align: center;
                padding: 20px;
            }

            .spinner {
                width: 48px;
                height: 48px;
                border: 4px solid rgba(255, 255, 255, 0.15);
                border-top-color: var(--brand-green);
                border-radius: 50%;
                animation: spin 0.8s linear infinite;
                margin-bottom: 18px;
            }

            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>

        <div class="top-banner">
            <span class="highlight">NUEVO:</span>
            <span>Motor de transposición cultural y adaptación didáctica con IA.</span>
            <span class="banner-close" onclick="this.parentElement.style.display='none'">✕</span>
        </div>

        <div class="nav-container">
            <nav class="floating-nav">
                <div class="nav-brand">
                    <div class="brand-badge">⚡</div>
                    <span>Cirugía Léxica</span>
                </div>

                <ul class="nav-links">
                    <li>Cómo funciona</li>
                    <li>Niveles MCER</li>
                    <li>Precios</li>
                </ul>

                <div class="nav-actions">
                    <span class="credits-chip">⚡ 5 créditos disponibles</span>
                </div>
            </nav>
        </div>

        <div class="hero-wrapper">
            
            <div class="hero-text">
                <div class="trust-caption">Para profesores de ELE que valoran su tiempo</div>
                <h1 class="hero-title">
                    Adapta tus manuales al país de tus estudiantes.
                </h1>
                <p class="hero-description">
                    Toma una foto a cualquier página de libro y transpórtala de inmediato a la cultura que necesitas, con glosario bilingüe y ejercicios calibrados en 15 segundos.
                </p>

                <div class="value-grid">
                    <div class="value-card">
                        <div class="v-icon-wrap">
                            <svg class="v-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                <polyline points="14 2 14 8 20 8"></polyline>
                                <line x1="16" y1="13" x2="8" y2="13"></line>
                                <line x1="16" y1="17" x2="8" y2="17"></line>
                                <line x1="10" y1="9" x2="8" y2="9"></line>
                            </svg>
                        </div>
                        <div class="v-title">Word (.docx) Editorial</div>
                        <div class="v-desc">Márgenes de imprenta, tipografía calibrada y sangrías de examen.</div>
                    </div>

                    <div class="value-card">
                        <div class="v-icon-wrap">
                            <svg class="v-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z"></path>
                                <path d="m9 12 2 2 4-4"></path>
                            </svg>
                        </div>
                        <div class="v-title">Solucionario Verificado</div>
                        <div class="v-desc">Cada clave respaldada por citas textuales directas del texto.</div>
                    </div>

                    <div class="value-card">
                        <div class="v-icon-wrap">
                            <svg class="v-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <circle cx="12" cy="12" r="10"></circle>
                                <polyline points="12 6 12 12 16 14"></polyline>
                                <path d="M12 2v2"></path>
                            </svg>
                        </div>
                        <div class="v-title">En 15 Segundos</div>
                        <div class="v-desc">Ahorra las 2 horas dominicales que toma transponer un texto a mano.</div>
                    </div>

                    <div class="value-card">
                        <div class="v-icon-wrap">
                            <svg class="v-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <line x1="4" y1="21" x2="4" y2="14"></line>
                                <line x1="4" y1="10" x2="4" y2="3"></line>
                                <line x1="12" y1="21" x2="12" y2="12"></line>
                                <line x1="12" y1="8" x2="12" y2="3"></line>
                                <line x1="20" y1="21" x2="20" y2="16"></line>
                                <line x1="20" y1="12" x2="20" y2="3"></line>
                                <line x1="1" y1="14" x2="7" y2="14"></line>
                                <line x1="9" y1="8" x2="15" y2="8"></line>
                                <line x1="17" y1="16" x2="23" y2="16"></line>
                            </svg>
                        </div>
                        <div class="v-title">Calibrado MCER</div>
                        <div class="v-desc">Alineación psicométrica real sin desvirtuar la dificultad original.</div>
                    </div>
                </div>
            </div>

            <div class="tool-card">
                <div class="tool-header">Subir y transformar actividad</div>

                <form id="eleForm" action="/adaptar-foto" method="post" enctype="multipart/form-data">
                    <div class="form-field">
                        <label class="field-caption">Página del manual</label>
                        <label class="file-dropzone" for="imagen">
                            <div class="dropzone-icon-box">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                    <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"></path>
                                    <circle cx="12" cy="13" r="4"></circle>
                                </svg>
                            </div>
                            <div class="dropzone-text-group">
                                <span class="dropzone-btn-action" id="dropzoneText">Elegir archivo o hacer foto</span>
                                <span class="dropzone-sub">JPG, PNG o captura directa</span>
                            </div>
                            <input type="file" id="imagen" name="imagen" accept="image/*" capture="environment" required>
                        </label>
                    </div>

                    <div class="form-field">
                        <label class="field-caption" for="destino">País o contexto meta</label>
                        <input type="text" id="destino" name="destino" placeholder="Ej: Lima (Perú), Buenos Aires, Tokio..." required>
                    </div>

                    <div class="grid-selectors form-field" style="grid-template-columns: 1fr 1fr 1fr;">
                        <div>
                            <label class="field-caption" for="modo">Modo</label>
                            <select id="modo" name="modo">
                                <option value="quirurgico" selected>Quirúrgico</option>
                                <option value="fluido">Fluido</option>
                            </select>
                        </div>
                        <div>
                            <label class="field-caption" for="nivel">Nivel MCER</label>
                            <select id="nivel" name="nivel">
                                <option value="AUTO" selected>Auto-detectar</option>
                                <option value="A1">A1</option>
                                <option value="A2">A2</option>
                                <option value="B1">B1</option>
                                <option value="B2">B2</option>
                                <option value="C1">C1</option>
                            </select>
                        </div>
                        <div>
                            <label class="field-caption" for="formato">Formato</label>
                            <select id="formato" name="formato">
                                <option value="docx" selected>Word (.docx)</option>
                                <option value="pdf">Documento PDF</option>
                            </select>
                        </div>
                    </div>

                    <div class="addons-box">
                        <label class="check-card">
                            <input type="checkbox" name="incluir_glosario" value="true">
                            <div>
                                <b>Glosario Bilingüe</b>
                                <span>Español - Inglés minimalista</span>
                            </div>
                        </label>

                        <!-- SLIDER ESTILO QUILLBOT -->
                        <div class="slider-wrapper">
                            <div class="slider-header">
                                <span class="slider-title">Batería de actividades:</span>
                                <span class="slider-badge" id="sliderBadge">Ninguna</span>
                            </div>

                            <div class="slider-track-container">
                                <input type="range" min="0" max="3" value="0" step="1" id="nivelActividades" name="profundidad_actividades" class="qb-slider">
                                <div class="qb-tooltip" id="sliderTooltip">Solo la lectura adaptada (sin ejercicios)</div>
                            </div>

                            <div class="slider-ticks">
                                <span>Sin ejercicios</span>
                                <span>Rápidas</span>
                                <span>Secuencia</span>
                                <span>Desafiantes</span>
                            </div>
                        </div>
                    </div>

                    <button type="submit" class="btn-cta">
                        <span>⚡ Generar Ficha Didáctica</span>
                    </button>
                </form>
            </div>

        </div>

        <div class="loader-screen" id="loadingOverlay">
            <div class="spinner"></div>
            <h2 style="font-size: 20px; font-weight: 800; margin-bottom: 6px;">Realizando Cirugía Léxica...</h2>
            <p style="font-size: 13.5px; color: #a4c0cd;">Analizando la imagen, transponiendo cultura y maquetando el documento.</p>
        </div>

        <script>
            const inputFoto = document.getElementById('imagen');
            const dropzoneText = document.getElementById('dropzoneText');
            const form = document.getElementById('eleForm');
            const overlay = document.getElementById('loadingOverlay');

            // Slider Reactivo
            const slider = document.getElementById('nivelActividades');
            const badge = document.getElementById('sliderBadge');
            const tooltip = document.getElementById('sliderTooltip');

            const estados = [
                {
                    badge: "Ninguna",
                    info: "Solo la lectura adaptada (sin ejercicios ni solucionario)"
                },
                {
                    badge: "Nivel 1: Rápidas",
                    info: "Verificación directa: Verdadero / Falso + Opción múltiple"
                },
                {
                    badge: "Nivel 2: Secuencia",
                    info: "Reorganización: Emparejar oraciones + Cronología de hechos"
                },
                {
                    badge: "Nivel 3: Desafiantes",
                    info: "Acción comunicativa: Estudio de caso, recomendación y debate"
                }
            ];

            function actualizarSlider() {
                const val = parseInt(slider.value);
                badge.textContent = estados[val].badge;
                tooltip.textContent = estados[val].info;
                const porcentaje = (val / 3) * 100;
                tooltip.style.left = porcentaje + "%";
            }

            slider.addEventListener('input', actualizarSlider);
            actualizarSlider();

            inputFoto.addEventListener('change', function() {
                if (this.files && this.files.length > 0) {
                    dropzoneText.textContent = "✓ " + this.files[0].name;
                }
            });

            form.addEventListener('submit', function() {
                overlay.style.display = 'flex';
                setTimeout(() => {
                    overlay.style.display = 'none';
                }, 18000);
            });
        </script>
    </body>
    </html>
    """


@app.post("/adaptar-foto")
@app.post("/adaptar-foto/")
async def adaptar_foto(
    imagen: UploadFile = File(...),
    destino: str = Form(...),
    modo: str = Form("quirurgico"),
    nivel: str = Form("AUTO"),
    formato: str = Form("docx"),
    incluir_glosario: bool = Form(False),
    profundidad_actividades: int = Form(0)
):
    try:
        contenido_bytes = await imagen.read()
        img = Image.open(io.BytesIO(contenido_bytes))

        if img.mode != "RGB":
            img = img.convert("RGB")

        max_lado = 1600
        if max(img.width, img.height) > max_lado:
            img.thumbnail((max_lado, max_lado), Image.Resampling.LANCZOS)

        buffer_jpg = io.BytesIO()
        img.save(buffer_jpg, format="JPEG", quality=85)
        bytes_optimizados = buffer_jpg.getvalue()

        # Configuración del Glosario Minimalista
        if incluir_glosario:
            instruccion_glosario = """
        6. EXCEPCIÓN DE ADICIÓN - GLOSARIO PEDAGÓGICO MINIMALISTA:
           - Inmediatamente después del texto didáctico adaptado, inserta el encabezado:
             ## GLOSARIO / VOCABULARY
           - Selecciona entre 5 y 8 términos o expresiones clave del texto adaptado.
           - Formato ESTRICTO y MINIMALISTA:
             * Cada término debe iniciar obligatoriamente con MAYÚSCULA.
             * Entrega ÚNICAMENTE la traducción directa y concisa al inglés (máximo 2 a 3 palabras en inglés).
             * PROHIBIDO incluir explicaciones, paráfrasis o definiciones en español.
             * Formato exacto por línea:
               • **Término**: Direct English translation
            """
        else:
            instruccion_glosario = ""

        # Construcción graduada según la posición del slider (0, 1, 2 o 3)
        bloques_ejercicios = []
        bloques_solucionario = []

        if profundidad_actividades >= 1:
            bloques_ejercicios.append("""
            ### NIVEL 1: VERIFICACIÓN Y COMPRENSIÓN SELECTIVA
            - Actividad 1.1 (Verdadero o Falso):
              * Redacta 4 afirmaciones directas basadas de forma literal en la lectura.
              * Al final de cada oración incluye únicamente la casilla para marcar: [ V ]  [ F ]
              * NO pidas justificación escrita al alumno ni agregues líneas de respuesta.
            - Actividad 1.2 (Opción Múltiple):
              * Redacta 3 preguntas con 3 alternativas cerradas cada una (a, b, c).
              * Cada alternativa correcta debe ser indiscutible y los distractores claramente falsos según el texto.
            """)
            bloques_solucionario.append("""
            * SOLUCIONARIO NIVEL 1:
              - Actividad 1.1: Indica claramente V o F para cada ítem acompañado de la cita textual entre comillas ("...") que lo demuestra.
              - Actividad 1.2: Indica la alternativa correcta (a, b o c) citando la justificación textual.
            """)
            
        if profundidad_actividades >= 2:
            bloques_ejercicios.append("""
            ### NIVEL 2: REORGANIZACIÓN Y COHERENCIA TEXTUAL
            - Actividad 2.1 (Emparejar mitades de oraciones):
              * Presenta 4 inicios de oración numerados (1, 2, 3, 4).
              * A continuación, presenta los 4 finales desordenados (a, b, c, d), cada uno precedido por una casilla '[   ]' para que el alumno escriba el número correspondiente:
                [   ] a) ...final de frase
                [   ] b) ...final de frase
                [   ] c) ...final de frase
                [   ] d) ...final de frase
              * PROHIBIDO añadir líneas extra de respuesta como 'Orden:', 'Solución:' o guiones al final.
            - Actividad 2.2 (Ordenar la secuencia lógica / itinerario):
              * Presenta 4 etapas clave desordenadas, cada una precedida únicamente por su casilla '[   ]' para numerar del 1 al 4.
              * PROHIBIDO añadir líneas extra al final como 'Secuencia cronológica:'.
            """)
            bloques_solucionario.append("""
            * SOLUCIONARIO NIVEL 2:
              - Actividad 2.1: Solución numérica directa en orden a, b, c, d (ejemplo: a=3, b=1, c=4, d=2).
              - Actividad 2.2: Orden cronológico correcto de las opciones (ejemplo: 4 - 2 - 1 - 3).
            """)

        if profundidad_actividades >= 3:
            bloques_ejercicios.append("""
            ### NIVEL 3: ACCIÓN SOCIAL Y APLICACIÓN COMUNICATIVA (MCER)
            - Actividad 3.1 (Estudio de caso y toma de decisiones):
              * Presenta 2 perfiles muy breves de personas con intereses o condiciones opuestas (p. ej., fechas disponibles, condición física o presupuesto).
              * Pide al estudiante que aconseje a una y desaconseje a la otra, exigiendo citar al menos 2 razones explícitas tomadas del texto.
            """)
            bloques_solucionario.append("""
            * PAUTAS PEDAGÓGICAS NIVEL 3:
              - Argumentos concretos del texto que el alumno debe mencionar en su respuesta.
            """)

        if bloques_ejercicios:
            texto_ejercicios = "\n".join(bloques_ejercicios)
            texto_solucionario = "\n".join(bloques_solucionario)
            instruccion_ejercicios = f"""
        7. ACTIVIDADES DIDÁCTICAS Y SOLUCIONARIO:
           - Tras el texto didáctico (y tras el glosario si fue solicitado), agrega el encabezado:
             ## ACTIVIDADES DE APRENDIZAJE
           {texto_ejercicios}

           - Al final absoluto de todo el documento, inserta el encabezado:
             ## CLAVE DE RESPUESTAS (SOLUCIONARIO PARA EL PROFESOR)
           - REGLAS INVIOLABLES DE RIGOR Y CERO ALUCINACIONES:
             * PRINCIPIO DE CLAUSURA: El texto adaptado es el ÚNICO universo de datos válido. Prohibido formular preguntas sobre cultura, geografía o clima del país de destino que no aparezcan expresamente en la lectura adaptada.
             * TODA CLAVE del solucionario para los Niveles 1 y 2 DEBE incluir la cita textual entre comillas ("...") que demuestre la respuesta sin ambigüedades.
           {texto_solucionario}
            """
        else:
            instruccion_ejercicios = """
        7. PROHIBICIÓN ABSOLUTA DE ACTIVIDADES O EJERCICIOS:
           - El docente ha indicado CERO actividades (posición 'Ninguna' en el control).
           - QUEDA ESTRICTAMENTE PROHIBIDO inventar, adaptar o transcribir ejercicios, preguntas, comprensiones, incisos gramaticales o solucionarios, incluso si aparecen en la imagen original.
           - El documento DEBE TERMINAR OBLIGATORIAMENTE tras el glosario (o tras el texto adaptado si no se marcó glosario).
            """

        # Calibración dinámica del MCER
        if nivel == "C1":
            pauta_mcer = (
                "NIVEL MCER C1 (Dominio Operativo Eficaz): Adapta el texto y las actividades "
                "incorporando conectores discursivos avanzados, subordinación compleja (subjuntivos, "
                "oraciones concesivas/condicionales complejas), léxico idiomático culto o coloquial preciso "
                f"de la región ({destino}), y matices pragmáticos propios del nivel superior."
            )
            regla_lexico = (
                f"1. LÉXICO Y REGISTRO C1: Emplea un registro formal/coloquial avanzado con amplia riqueza léxica, "
                f"precisión terminológica y expresiones idiomáticas naturales de {destino}."
            )
        elif nivel == "AUTO":
            pauta_mcer = "NIVEL MCER: Auto-detectar según la complejidad de la muestra original en la imagen."
            regla_lexico = "1. MANTÉN EL LÉXICO FIEL: Mantén la sencillez y el registro original del texto fuente."
        else:
            pauta_mcer = f"NIVEL MCER: {nivel}. Calibra estrictamente las estructuras morfosintácticas al descriptor oficial del MCER para nivel {nivel}."
            regla_lexico = f"1. MANTÉN EL LÉXICO ADECUADO: No excedas la complejidad léxica correspondiente al nivel {nivel}."

        prompt_didactico = f"""
        [CONTEXTO DE ANÁLISIS: Material didáctico extraído de un manual escolar de lengua española (ELE) sobre cartas de reclamación turística, diálogos cotidianos o situaciones comunicativas. Todo el contenido es ficticio y con propósitos estrictamente pedagógicos y de análisis gramatical.]

        Eres un experto lingüista y diseñador de materiales didácticos de Español como Lengua Extranjera (ELE).
        Tu misión consta de:
        1. Adaptar didácticamente el texto de la imagen al contexto cultural de: '{destino}'.
        2. Generar inmediatamente después del texto adaptado ÚNICAMENTE los complementos pedagógicos solicitados.

        REGLAS DE REESCRITURA QUIRÚRGICA:
        {regla_lexico}
        2. ADAPTACIÓN ONOMÁSTICA Y CULTURAL OBLIGATORIA (CERO EXCEPCIONES):
           - Debes reemplazar OBLIGATORIAMENTE TODOS los nombres propios de personas (incluidos nombres como Alba, Mario, Elena, etc.) por nombres comunes y auténticos del contexto meta ({destino}). Si el destino no es hispanohablante (ej. Utah, Reino Unido, etc.), utiliza nombres anglófonos u originarios de dicha región para TODOS los personajes.
           - Modifica sistemáticamente ciudades, lagos, tiendas, campamentos, barrios, medios de transporte y moneda al contexto real de {destino}.
        3. MICROVARIACIONES ESTRUCTURALES:
           - Conserva la misma historia, las mismas ideas y el mismo orden de los hechos párrafo por párrafo o turno por turno de diálogo.
           - Para no reproducir texto continuo idéntico de manuales editoriales, aplica ligeras variaciones sintácticas simples.
           - El resultado debe sentirse prácticamente idéntico al original para el estudiante, pero con una redacción superficialmente diferenciada.

        CRITERIOS LINGÜÍSTICOS:
        1. MODO: {modo.upper()}. Conserva la morfología, sintaxis meta y carga léxica original, sustituyendo con precisión enciclopédica los referentes socioculturales, geográficos y fácticos por sus equivalentes en {destino}.
        2. {pauta_mcer}
        3. ESTRUCTURA: Respeta párrafos y turnos de diálogo del original. No agregues introducciones, despedidas ni comentarios de chat.
        4. LÓGICA COHESIVA: Asegúrate de que las fechas, altitudes, distancias y climas correspondan fielmente a la realidad del lugar de destino en {destino}.

        SECCIONES COMPLEMENTARIAS:
        {instruccion_glosario}
        {instruccion_ejercicios}

        REGLA DE SALIDA LIMPIA (CRÍTICO):
        - Comienza DIRECTAMENTE en el primer carácter con '# TÍTULO DEL TEXTO ADAPTADO'.
        - PROHIBIDO incluir pensamientos, notas previas o monólogos internos.
        - Salida estricta: incluye el texto adaptado y, si fueron explícitamente activados, el glosario y las actividades. Si no se solicitaron actividades, concluye el documento inmediatamente.
        """
        
        partes_contenido = [
            types.Part.from_bytes(data=bytes_optimizados, mime_type="image/jpeg"),
            prompt_didactico
        ]

        configuracion = types.GenerateContentConfig(
            temperature=0.2,
            max_output_tokens=8192,
            thinking_config=types.ThinkingConfig(thinking_budget=0),
            safety_settings=[
                types.SafetySetting(
                    category=types.HarmCategory.HARM_CATEGORY_HARASSMENT,
                    threshold=types.HarmBlockThreshold.BLOCK_NONE,
                ),
                types.SafetySetting(
                    category=types.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
                    threshold=types.HarmBlockThreshold.BLOCK_NONE,
                ),
                types.SafetySetting(
                    category=types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
                    threshold=types.HarmBlockThreshold.BLOCK_NONE,
                ),
                types.SafetySetting(
                    category=types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
                    threshold=types.HarmBlockThreshold.BLOCK_NONE,
                ),
                types.SafetySetting(
                    category=types.HarmCategory.HARM_CATEGORY_CIVIC_INTEGRITY,
                    threshold=types.HarmBlockThreshold.BLOCK_NONE,
                ),
            ],
        )

        respuesta = cliente_gemini.models.generate_content(
            model="gemini-3.8-flash",
            contents=partes_contenido,
            config=configuracion
        )

        texto_generado = ""
        if respuesta.candidates and len(respuesta.candidates) > 0:
            candidate = respuesta.candidates[0]
            if candidate.content and candidate.content.parts:
                partes_utiles = []
                for p in candidate.content.parts:
                    if getattr(p, "thought", False):
                        continue
                    if hasattr(p, "text") and p.text:
                        partes_utiles.append(p.text)
                if partes_utiles:
                    texto_generado = "".join(partes_utiles)

        if not texto_generado and respuesta.text:
            texto_generado = respuesta.text

        if not texto_generado:
            raise HTTPException(
                status_code=500,
                detail="La API no devolvió contenido de texto para esta imagen."
            )

        # Limpiar cualquier residuo previo al título
        if "#" in texto_generado:
            texto_generado = texto_generado[texto_generado.find("#"):]

        nombre_base = f"Actividad_ELE_{destino.replace(' ', '_')}"

        if formato == "pdf":
            buffer_pdf = fabricar_pdf(texto_generado, destino)
            return StreamingResponse(
                buffer_pdf,
                media_type="application/pdf",
                headers={"Content-Disposition": f'attachment; filename="{nombre_base}.pdf"'}
            )
        else:
            buffer_docx = fabricar_word(texto_generado, destino)
            return StreamingResponse(
                buffer_docx,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": f'attachment; filename="{nombre_base}.docx"'}
            )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error durante el procesamiento didáctico: {str(e)}")