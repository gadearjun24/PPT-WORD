import os
import uuid
import logging
from io import BytesIO

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

from docx import Document
from docx.shared import Pt, RGBColor as DocxRGB

import matplotlib.pyplot as plt

# -------------------------
# Logging
# -------------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
logger = logging.getLogger(__name__)

# -------------------------
# FastAPI app
# -------------------------
app = FastAPI(title="PPTX → Word Converter (continuous mode)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    allow_credentials=True,
)

# -------------------------
# Helpers
# -------------------------
def pptx_color_to_rgb(color_obj):
    try:
        if color_obj and color_obj.type == 1:
            rgb = color_obj.rgb
            return DocxRGB(rgb[0], rgb[1], rgb[2])
    except Exception:
        pass
    return None

def safe_get_text(shape):
    try:
        if hasattr(shape, "text"):
            return shape.text or ""
    except Exception:
        pass
    return ""

def extract_image_from_shape(shape):
    """Try to extract image bytes from shape."""
    try:
        if hasattr(shape, "image") and shape.image:
            return shape.image.blob
    except Exception:
        pass

    try:
        pic = getattr(shape, "pic", None)
        if pic is not None and hasattr(pic, "image"):
            return pic.image.blob
    except Exception:
        pass

    try:
        blip = shape.element.xpath(".//a:blip")
        if blip:
            rId = blip[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            rel = shape.part.related_parts.get(rId)
            if rel:
                return rel.blob
    except Exception:
        pass

    raise RuntimeError("No image found in shape")

def render_chart_from_chart_data(chart):
    """Render chart data as matplotlib fallback."""
    try:
        data = chart.chart_data
        categories = [str(c) for c in data.categories] if data.categories else []
        fig, ax = plt.subplots(figsize=(6, 4))
        for s in data.series:
            values = list(s.values)
            ax.plot(categories if categories else range(len(values)), values, marker="o", label=s.name)
        ax.legend()
        ax.set_title("Chart Rendered (fallback)")
        plt.tight_layout()
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf.read()
    except Exception as e:
        raise RuntimeError(f"Failed to render chart: {e}")

# -------------------------
# Main endpoint
# -------------------------
@app.post("/convert/")
async def convert(file: UploadFile = File(...)):
    try:
        logger.info(f"Received file: {file.filename}")
        content = await file.read()
        pptx_path = f"/tmp/{uuid.uuid4()}.pptx"
        with open(pptx_path, "wb") as f:
            f.write(content)

        prs = Presentation(pptx_path)
        doc = Document()

        # Detect font
        default_font = "Utsaah"
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.name:
                                default_font = run.font.name
                                break
                        if default_font != "Utsaah":
                            break
                if default_font != "Utsaah":
                    break
            if default_font != "Utsaah":
                break

        logger.info(f"Using default font: {default_font}")

        # -----------------------------
        # Process slides continuously
        # -----------------------------
        for s_i, slide in enumerate(prs.slides, start=1):
            logger.info(f"Slide {s_i}/{len(prs.slides)}")

            # Optional: add slide header (not page break)
            doc.add_paragraph(f"--- Slide {s_i} ---").runs[0].bold = True

            for shape in slide.shapes:
                shape_type = shape.shape_type

                # TEXT
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        if not para.text.strip():
                            continue
                        p = doc.add_paragraph()
                        for run in para.runs:
                            r = p.add_run(run.text)
                            r.font.name = default_font
                            r.font.size = Pt(14)
                            r.bold = run.font.bold
                            r.italic = run.font.italic
                            rgb = pptx_color_to_rgb(run.font.color)
                            if rgb:
                                r.font.color.rgb = rgb

                # TABLE
                elif shape_type == MSO_SHAPE_TYPE.TABLE:
                    try:
                        table = shape.table
                        word_table = doc.add_table(rows=len(table.rows), cols=len(table.columns))
                        word_table.style = "Table Grid"
                        for r_i, row in enumerate(table.rows):
                            for c_i, cell in enumerate(row.cells):
                                word_table.cell(r_i, c_i).text = cell.text.strip()
                        doc.add_paragraph("")  # spacing after table
                    except Exception as e:
                        logger.warning(f"Table failed: {e}")

                # IMAGE
                elif shape_type == MSO_SHAPE_TYPE.PICTURE or "blip" in str(shape.element.xml):
                    try:
                        img_bytes = extract_image_from_shape(shape)
                        doc.add_picture(BytesIO(img_bytes), width=Inches(4))
                        doc.add_paragraph("")  # spacing
                    except Exception as e:
                        logger.warning(f"Image extraction failed: {e}")

                # CHART
                elif hasattr(shape, "chart"):
                    try:
                        try:
                            part = shape.chart.chart_part
                            blob = part.chart_space.blob
                            doc.add_picture(BytesIO(blob), width=Inches(5))
                        except Exception:
                            # fallback to rendering
                            img = render_chart_from_chart_data(shape.chart)
                            doc.add_picture(BytesIO(img), width=Inches(5))
                        doc.add_paragraph("")
                    except Exception as e:
                        logger.warning(f"Chart failed: {e}")

                # SHAPE (Rectangles, circles, etc.)
                elif shape_type in [
                    MSO_SHAPE_TYPE.AUTO_SHAPE,
                    MSO_SHAPE_TYPE.FREEFORM,
                    MSO_SHAPE_TYPE.GROUP,
                ]:
                    try:
                        shape_name = getattr(shape, "name", "Shape")
                        doc.add_paragraph(f"[Shape detected: {shape_name}]")
                    except Exception:
                        doc.add_paragraph("[Shape detected]")

            # Add 2 blank lines between slides (instead of page break)
            doc.add_paragraph("")
            doc.add_paragraph("")

        # Save and return
        out_path = f"/tmp/{uuid.uuid4()}.docx"
        doc.save(out_path)

        def iterfile():
            with open(out_path, "rb") as f:
                yield from f

        return StreamingResponse(
            iterfile(),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{os.path.splitext(file.filename)[0]}.docx"'},
        )

    except Exception as e:
        logger.exception("Conversion failed")
        return JSONResponse(content={"error": str(e)}, status_code=500)


@app.get("/")
def health():
    return {"status": "ok"}
