from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
from docx.shared import Pt, RGBColor as DocxRGB
import shutil
import uuid
import os
from fastapi.middleware.cors import CORSMiddleware






app = FastAPI(title="PPTX to Word Converter")


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # allow all origins; you can restrict to your frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



# Helper: Convert pptx RGB to Word RGB
def pptx_color_to_rgb(color_obj):
    try:
        if color_obj and color_obj.type == 1:  # RGB
            rgb = color_obj.rgb
            return DocxRGB(rgb[0], rgb[1], rgb[2])
    except Exception:
        pass
    return None

@app.post("/convert/")
async def convert_pptx_to_word(file: UploadFile = File(...)):
    # Save uploaded file temporarily
    temp_pptx = f"/tmp/{uuid.uuid4()}.pptx"
    with open(temp_pptx, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # Load PPTX
    prs = Presentation(temp_pptx)
    doc = Document()

    # Detect default font from PPTX
    default_font_name = "Utsaah"
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.name:
                            default_font_name = run.font.name
                            break
                    if default_font_name != "Utsaah":
                        break
            if default_font_name != "Utsaah":
                break
        if default_font_name != "Utsaah":
            break

    # Extract text and format
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX or shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    if not paragraph.text.strip():
                        continue
                    p = doc.add_paragraph()
                    for run in paragraph.runs:
                        new_run = p.add_run(run.text)
                        new_run.font.name = default_font_name
                        new_run.font.size = Pt(14)
                        new_run.bold = run.font.bold
                        new_run.italic = run.font.italic
                        new_run.underline = run.font.underline

                        rgb = pptx_color_to_rgb(run.font.color)
                        if rgb:
                            new_run.font.color.rgb = rgb

                    p.alignment = paragraph.alignment

            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table = shape.table
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if any(cells):
                        p = doc.add_paragraph(" | ".join(cells))
                        for run in p.runs:
                            run.font.name = default_font_name
                            run.font.size = Pt(14)

        # Add two new lines after each slide
        doc.add_paragraph("")
        doc.add_paragraph("")

    # Save Word file
    output_path = f"/tmp/{uuid.uuid4()}.docx"
    doc.save(output_path)

    # Return file for download
    return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="converted.docx")
