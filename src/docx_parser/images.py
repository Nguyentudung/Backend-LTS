import zipfile
from pathlib import Path

# region Ảnh từ file .docx

# Hàm để trích xuất ảnh từ file .docx
def extract_images(
    docx_path: str,
    output_dir: str | Path = "extracted_images",
    public_dir: str | None = None,
):
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True, parents=True)
    public_dir = (public_dir or output_dir.name).strip("/")

    images: list[str] = []

    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if not name.startswith("word/media/"):
                continue

            data = z.read(name)
            filename = Path(name).name
            out_path = output_dir / filename

            with open(out_path, "wb") as f:
                f.write(data)

            # Use POSIX-style paths so they work as browser URLs (Windows paths use `\`).
            if public_dir:
                images.append(f"{public_dir}/{filename}")
            else:
                images.append(filename)

    return images
# endregion