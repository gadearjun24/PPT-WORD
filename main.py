import os
import uuid
import logging
from io import BytesIO

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from docx import Document
from docx.shared import Pt, RGBColor as DocxRGB, Inches

# matplotlib is used to re-render charts when no image is available
import matplotlib.pyplot as plt

# -------------------------
# Logging
# -------------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
logger = logging.getLogger(__name__)

# -------------------------
# FastAPI app
# -------------------------
app = FastAPI(title="PPTX -> Word Converter (robust)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],    # restrict in prod if needed
    allow_methods=["*"],
    allow_headers=["*"],
    allow_credentials=True,
)

# -------------------------
# Helpers
# -------------------------
def pptx_color_to_rgb(color_obj):
    try:
        if color_obj and color_obj.type == 1:  # RGB
            rgb = color_obj.rgb
            return DocxRGB(rgb[0], rgb[1], rgb[2])
    except Exception as e:
        logger.debug(f"pptx_color_to_rgb: {e}")
    return None

def safe_get_text(shape):
    """Return text from a shape if available in a safe way."""
    try:
        if hasattr(shape, "text"):
            return shape.text or ""
    except Exception as e:
        logger.debug(f"safe_get_text: {e}")
    return ""

def save_stream_to_file(stream_bytes: bytes, ext: str = "png"):
    """Save bytes to a temporary file and return path."""
    path = f"/tmp/{uuid.uuid4()}.{ext}"
    with open(path, "wb") as f:
        f.write(stream_bytes)
    return path

def render_chart_from_chart_data(chart):
    """
    Attempt to extract series/categories from chart.chart_data (python-pptx ChartData)
    and render a simple image with matplotlib. Returns bytes or raises.
    """
    try:
        chart_data = chart.chart_data  # may raise if not available
    except Exception as e:
        raise RuntimeError(f"No chart_data available: {e}")

    # categories (x-axis)
    try:
        categories = list(chart_data.categories)
        categories = [str(c) for c in categories]
    except Exception:
        categories = None

    # series (list of Series)
    series_list = []
    try:
        for s in chart_data.series:
            label = s.name if hasattr(s, "name") else None
            values = list(s.values)
            series_list.append((label, values))
    except Exception:
        raise RuntimeError("Failed to read series from chart_data")

    if not series_list:
        raise RuntimeError("No series data found in chart_data")

    # choose chart type fallback: try to inspect chart.chart_type if present
    chart_type_name = ""
    try:
        chart_type_name = str(chart.chart_type).lower()
    except Exception:
        chart_type_name = ""

    # Create matplotlib figure
    fig, ax = plt.subplots(figsize=(6, 4))
    try:
        # If categories present and single series -> bar or pie
        if len(series_list) == 1:
            label, values = series_list[0]
            if "pie" in chart_type_name:
                ax.pie(values, labels=categories if categories else None, autopct="%1.1f%%")
            elif "bar" in chart_type_name or "column" in chart_type_name or (categories is not None):
                ax.bar(categories if categories else list(range(len(values))), values, label=label)
                if label:
                    ax.legend()
                ax.set_xticklabels(categories if categories else [str(i) for i in range(len(values))], rotation=45, ha="right")
            else:
                # default to line
                ax.plot(categories if categories else list(range(len(values))), values, marker="o", label=label)
                ax.legend()
        else:
            # multiple series -> grouped bar or line
            x = range(len(series_list[0][1]))
            width = 0.8 / max(1, len(series_list))
            for idx, (label, values) in enumerate(series_list):
                pos = [xi + (idx - len(series_list)/2) * width + width/2 for xi in x]
                ax.bar(pos, values, width=width, label=(label or f"Series {idx+1}"))
            ax.set_xticks(x)
            if categories:
                ax.set_xticklabels(categories, rotation=45, ha="right")
            ax.legend()
    except Exception as e:
        plt.close(fig)
        raise RuntimeError(f"Failed to render chart via matplotlib: {e}")

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def extract_image_from_shape(shape):
    """
    Try several ways to extract an image from a shape.
    Returns bytes or raises.
    """
    # 1) shape has .image attribute (PICTURE shapes)
    try:
        if hasattr(shape, "image") and shape.image is not None:
            return shape.image.blob
    except Exception as e:
        logger.debug(f"extract_image_from_shape: shape.image failed: {e}")

    # 2) some placeholders or shapes may contain a picture as part of their element - try shape.element
    try:
        pic = getattr(shape, "pic", None)
        if pic is not None and hasattr(pic, "image"):
            return pic.image.blob
    except Exception as e:
        logger.debug(f"extract_image_from_shape: shape.pic failed: {e}")

    # 3) try to read shape._element for an <a:blip> reference (advanced, might fail)
    try:
        blip = shape.element.xpath(".//a:blip")
        if blip:
            rId = blip[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            rel = shape.part.related_parts.get(rId)
            if rel:
                return rel.blob
    except Exception as e:
        logger.debug(f"extract_image_from_shape: element.blip failed: {e}")

    raise RuntimeError("No image found in shape")

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
        logger.info(f"Saved PPTX to {pptx_path}")

        prs = Presentation(pptx_path)
        logger.info(f"Loaded PPTX ({len(prs.slides)} slides)")

        # Create Word document
        doc = Document()

        # Try to detect default font from PPTX runs (keep previous logic)
        default_font_name = "Utsaah"
        for slide in prs.slides:
            for shape in slide.shapes:
                try:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                if run.font.name:
                                    default_font_name = run.font.name
                                    break
                            if default_font_name != "Utsaah":
                                break
                except Exception:
                    continue
            if default_font_name != "Utsaah":
                break
        logger.info(f"Default font detected: {default_font_name}")

        # Process slides (in-order)
        for s_i, slide in enumerate(prs.slides, start=1):
            logger.info(f"Processing slide {s_i}/{len(prs.slides)}")
            # Add slide heading in Word to separate slides
            doc.add_heading(f"Slide {s_i}", level=2)

            for sh_i, shape in enumerate(slide.shapes, start=1):
                logger.info(f"Processing shape {sh_i}/{len(slide.shapes)} type={shape.shape_type}")

                # 1) Text (including placeholders)
                text = ""
                try:
                    text = safe_get_text(shape).strip()
                except Exception as e:
                    logger.debug(f"safe_get_text error: {e}")

                if text:
                    # Break into paragraphs and preserve run-level formatting when possible
                    try:
                        if hasattr(shape, "text_frame") and shape.has_text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                if not paragraph.text.strip():
                                    continue
                                p = doc.add_paragraph()
                                for run in paragraph.runs:
                                    r = p.add_run(run.text)
                                    # basic formatting
                                    try:
                                        r.font.name = default_font_name
                                        r.font.size = Pt(14)
                                        r.bold = run.font.bold
                                        r.italic = run.font.italic
                                        r.underline = run.font.underline
                                        rgb = pptx_color_to_rgb(run.font.color)
                                        if rgb:
                                            r.font.color.rgb = rgb
                                    except Exception as e:
                                        logger.debug(f"run formatting failed: {e}")
                                # set paragraph alignment if available
                                try:
                                    p.alignment = paragraph.alignment
                                except Exception:
                                    pass
                        else:
                            doc.add_paragraph(text)
                        logger.info(f"Extracted text (len={len(text)})")
                    except Exception as e:
                        logger.warning(f"Failed to extract formatted text: {e}")
                        doc.add_paragraph(text)

                    # continue to next shape (but do not 'continue' if shape also contains image/table)
                    # we purposely do not skip further checks — some shapes may contain both text and image/table

                # 2) Table
                try:
                    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                        table = shape.table
                        rows, cols = len(table.rows), len(table.columns)
                        logger.info(f"Table detected ({rows}x{cols})")
                        word_table = doc.add_table(rows=rows, cols=cols)
                        word_table.style = "Table Grid"
                        for r_idx, row in enumerate(table.rows):
                            for c_idx, cell in enumerate(row.cells):
                                txt = cell.text.strip()
                                word_table.cell(r_idx, c_idx).text = txt
                        logger.info("Table written to Word")
                        # keep checking other content in shape (rare), so don't `continue`
                except Exception as e:
                    logger.warning(f"Table extraction failed: {e}")

                # 3) Image (try robust extraction)
                try:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or hasattr(shape, "image") or "blip" in str(shape.element.xml):
                        try:
                            img_bytes = extract_image_from_shape(shape)
                            img_stream = BytesIO(img_bytes)
                            # add caption/paragraph before image if text exists
                            if text:
                                doc.add_paragraph("")  # spacing
                            doc.add_picture(img_stream, width=Inches(4))
                            logger.info("Inserted image into Word")
                        except Exception as e_img:
                            logger.debug(f"No direct image extracted from shape: {e_img}")
                except Exception as e:
                    logger.debug(f"Image detection error: {e}")

                # 4) Chart handling
                try:
                    if hasattr(shape, "chart"):
                        logger.info("Chart found in shape; attempting extraction")
                        # First: try to get an embedded image (chart part) if available
                        chart_inserted = False
                        try:
                            # Some chart objects expose a chart_part -> chart_space -> blob
                            # This may fail on some chart types; attempt defensively
                            chart_part = getattr(shape.chart, "chart_part", None)
                            if chart_part is not None:
                                chart_blob = chart_part.chart_space.blob
                                img_stream = BytesIO(chart_blob)
                                doc.add_picture(img_stream, width=Inches(5))
                                logger.info("Inserted chart image from chart_part")
                                chart_inserted = True
                        except Exception as e_part:
                            logger.debug(f"chart_part extraction failed: {e_part}")

                        # Second fallback: try to render chart via chart_data using matplotlib
                        if not chart_inserted:
                            try:
                                logger.info("Attempting to render chart from chart_data")
                                img_bytes = render_chart_from_chart_data(shape.chart)
                                img_stream = BytesIO(img_bytes)
                                doc.add_picture(img_stream, width=Inches(5))
                                logger.info("Rendered chart via matplotlib and inserted")
                                chart_inserted = True
                            except Exception as e_chart:
                                logger.warning(f"Chart render fallback failed: {e_chart}")
                                # As a last resort add a textual summary
                                try:
                                    chart_type = getattr(shape.chart, "chart_type", "Unknown")
                                    doc.add_paragraph(f"[Chart could not be rendered. Chart type: {chart_type}]")
                                except Exception:
                                    doc.add_paragraph("[Chart could not be rendered]")
                                logger.info("Inserted chart placeholder text")
                except Exception as e:
                    logger.debug(f"Chart detection error: {e}")

            # end per-slide shapes
            # Add a page break between slides to keep things separated
            doc.add_page_break()

        # Save Word document
        out_path = f"/tmp/{uuid.uuid4()}.docx"
        doc.save(out_path)
        logger.info(f"Saved Word to {out_path}")

        # Stream back
        def iterfile():
            with open(out_path, "rb") as f:
                yield from f

        logger.info("Conversion finished; returning file")
        return StreamingResponse(iterfile(),
                                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                headers={"Content-Disposition": f'attachment; filename="{os.path.splitext(file.filename)[0]}.docx"'})
    except Exception as e:
        logger.exception("Conversion failed")
        return JSONResponse(content={"error": str(e)}, status_code=500)

# Health endpoint
@app.get("/")
def health():
    return {"status": "ok"}
