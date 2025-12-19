from pathlib import Path

from .convert_images import convert_wmf_to_png
from .flow_parser import parse_flow
from .formula import extract_formulas
from .images import extract_images
from .questions import parse_questions
from .text import extract_text_with_highlight

# region Phân tích file .docx

# Hàm để phân tích file .docx và trích xuất văn bản, công thức, ảnh, và cấu trúc luồng
def parse_docx(
    docx_path: str,
    *,
    image_output_dir: str | Path = "extracted_images",
    image_public_dir: str | None = None,
):
    image_output_dir = Path(image_output_dir)
    image_public_dir = (image_public_dir or image_output_dir.name).strip("/")

    images = extract_images(
        docx_path,
        output_dir=image_output_dir,
        public_dir=image_public_dir,
    )
    try:
        convert_wmf_to_png(str(image_output_dir))
    except Exception as e:
        print(f"[wmf] convert_wmf_to_png failed: {e}", flush=True)

    final_images: list[str] = []
    for img in images:
        if img.lower().endswith(".wmf"):
            png_disk = (image_output_dir / Path(img).name).with_suffix(".png")
            png_public = Path(img).with_suffix(".png").as_posix()
            final_images.append(png_public if png_disk.exists() else img)
        else:
            final_images.append(img)

    texts = extract_text_with_highlight(docx_path)
    flow = parse_flow(docx_path, image_dir=image_public_dir)

    def replace_wmf_in_blocks(blocks: list[dict]):
        for block in blocks:
            if not isinstance(block, dict):
                continue

            if block.get("type") == "image":
                src = block.get("src")
                if isinstance(src, str) and src.lower().endswith(".wmf"):
                    png_disk = (image_output_dir / Path(src).name).with_suffix(".png")
                    if png_disk.exists():
                        block["src"] = Path(src).with_suffix(".png").as_posix()

            if block.get("type") == "table":
                rows = block.get("rows")
                if not isinstance(rows, list):
                    continue
                for row in rows:
                    if not isinstance(row, list):
                        continue
                    for cell in row:
                        if not isinstance(cell, dict):
                            continue
                        cell_blocks = cell.get("blocks")
                        if isinstance(cell_blocks, list):
                            replace_wmf_in_blocks(cell_blocks)

    replace_wmf_in_blocks(flow)

    return {
        "texts": texts,
        "formulas": extract_formulas(docx_path),
        "images": final_images,
        "flow": flow,
        "questions": parse_questions(texts=texts, flow=flow),
    }
# endregion
