from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_SHAPE
from docx import Document
from docx.shared import Pt, RGBColor as DocxRGB, Inches
from io import BytesIO
import shutil
import uuid
import logging
from fastapi.middleware.cors import CORSMiddleware

# -----------------------------
# Logging configuration
# -----------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)
logger = logging.getLogger(__name__)

# -----------------------------
# FastAPI app
# -----------------------------
app = FastAPI(title="PPTX to Word Converter")

# CORS: allow all origins for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # or your frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------------
# Helper: Convert pptx RGB to Word RGB
# -----------------------------
def pptx_color_to_rgb(color_obj):
    try:
        if color_obj and color_obj.type == 1:  # RGB
            rgb = color_obj.rgb
            return DocxRGB(rgb[0], rgb[1], rgb[2])
    except Exception as e:
        logger.warning(f"Failed to get color: {e}")
    return None

# -----------------------------
# Endpoint: Convert PPTX to Word
# -----------------------------
@app.post("/convert/")
async def convert_pptx_to_word(file: UploadFile = File(...)):
    logger.info(f"📁 Received file: {file.filename}")

    # Save uploaded file temporarily
    temp_pptx = f"/tmp/{uuid.uuid4()}.pptx"
    try:
        with open(temp_pptx, "wb") as f:
            shutil.copyfileobj(file.file, f)
        logger.info(f"✅ Saved uploaded PPTX to {temp_pptx}")
    except Exception as e:
        logger.error(f"❌ Error saving uploaded file: {e}")
        return {"error": "Failed to save uploaded file"}

    # Load presentation
    try:
        prs = Presentation(temp_pptx)
        logger.info(f"✅ Loaded PPTX with {len(prs.slides)} slides")
    except Exception as e:
        logger.error(f"❌ Error loading PPTX: {e}")
        return {"error": "Failed to load PPTX file"}

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
    logger.info(f"🖋️ Default font detected: {default_font_name}")

    # Process slides
    for slide_index, slide in enumerate(prs.slides):
        logger.info(f"📄 Processing slide {slide_index + 1}/{len(prs.slides)}")

        for shape_index, shape in enumerate(slide.shapes):
            logger.info(f"🔹 Shape {shape_index + 1}/{len(slide.shapes)} type={shape.shape_type}")

            # -----------------------------
            # Text boxes
            # -----------------------------
            if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX or shape.has_text_frame:
                try:
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
                except Exception as e:
                    logger.warning(f"⚠️ Failed to process text box: {e}")

            # -----------------------------
            # Tables
            # -----------------------------
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                try:
                    table = shape.table
                    word_table = doc.add_table(rows=len(table.rows), cols=len(table.columns))
                    word_table.style = "Table Grid"
                    for i, row in enumerate(table.rows):
                        for j, cell in enumerate(row.cells):
                            word_table.cell(i, j).text = cell.text.strip()
                    logger.info(f"✅ Table extracted ({len(table.rows)}x{len(table.columns)})")
                except Exception as e:
                    logger.warning(f"⚠️ Failed to process table: {e}")

            # -----------------------------
            # Images
            # -----------------------------
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image_stream = BytesIO(shape.image.blob)
                    doc.add_picture(image_stream, width=Inches(4))
                    logger.info(f"🖼️ Inserted image from shape {shape_index + 1}")
                except Exception as e:
                    logger.warning(f"⚠️ Failed to insert image: {e}")

            # -----------------------------
            # Charts (as images)
            # -----------------------------
            elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                try:
                    chart_part = shape.chart.plots[0].chart_part
                    chart_blob = chart_part.chart_space.blob
                    image_stream = BytesIO(chart_blob)
                    doc.add_picture(image_stream, width=Inches(4))
                    logger.info(f"📊 Inserted chart as image")
                except Exception as e:
                    logger.warning(f"⚠️ Failed to extract chart: {e}")

        # Add spacing between slides
        doc.add_paragraph("")
        doc.add_paragraph("")

    # Save Word file
    output_path = f"/tmp/{uuid.uuid4()}.docx"
    try:
        doc.save(output_path)
        logger.info(f"💾 Saved converted Word file: {output_path}")
    except Exception as e:
        logger.error(f"❌ Failed to save Word file: {e}")
        return {"error": "Failed to save Word file"}

    # Return file for download
    try:
        f = open(output_path, "rb")
        headers = {"Content-Disposition": 'attachment; filename=\"converted.docx\"'}
        logger.info(f"✅ Conversion successful, returning file.")
        return StreamingResponse(
            f,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )
    except Exception as e:
        logger.error(f"❌ Failed to return Word file: {e}")
        return {"error": "Failed to return Word file"}
