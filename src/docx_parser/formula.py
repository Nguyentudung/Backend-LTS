import zipfile
import xml.etree.ElementTree as ET

NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M = f"{{{NS_M}}}"

# region Công thức OMML

# Hàm để trích xuất công thức từ file .docx
def extract_formulas(docx_path: str):
    """
    Trả về danh sách công thức (OMML XML string).

    Lưu ý: Nhiều file Word cũ/MathType/Equation Editor sẽ lưu công thức dạng OLE + ảnh preview,
    khi đó danh sách này có thể rỗng và công thức sẽ được lấy qua `flow` như image.
    """
    formulas: list[str] = []

    with zipfile.ZipFile(docx_path) as z:
        xml_content = z.read("word/document.xml")

    root = ET.fromstring(xml_content)

    for elem in root.iter(f"{M}oMath"):
        formulas.append(ET.tostring(elem, encoding="unicode"))

    return formulas
# endregion