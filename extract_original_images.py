"""
extract_original_images.py — Word COM으로 문서를 Flat XML로 저장한 뒤,
그 안에 base64로 내장된 그림을 원본 바이트 그대로 폴더에 꺼낸다.
(웹 페이지 저장은 96dpi로 재샘플링하지만, 이 방식은 Word가 보관한 원본을 그대로 준다)

사용:
    python extract_original_images.py "C:\\test\\sample.docx"
    → C:\\test\\sample_orig_media\\image1.jpeg ... 와 각 파일의 픽셀 크기 출력
"""
import sys, re, base64, pathlib
import win32com.client

WD_FLAT_XML = 19


def main():
    if len(sys.argv) < 2:
        print("사용: python extract_original_images.py <docx 경로>")
        sys.exit(1)
    src = pathlib.Path(sys.argv[1]).resolve()
    out_dir = src.with_name(src.stem + "_orig_media")
    out_dir.mkdir(exist_ok=True)
    xml_path = src.with_suffix(".flat.xml")

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
    try:
        doc.SaveAs2(str(xml_path), FileFormat=WD_FLAT_XML)
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

    head = xml_path.read_bytes()[:16]
    if not head.lstrip().startswith(b"<"):
        print("Flat XML 결과가 DRM으로 감싸져 있음 — 이 경로는 막힌 것")
        sys.exit(2)

    xml = xml_path.read_text(encoding="utf-8", errors="ignore")
    parts = re.finditer(
        r'<pkg:part pkg:name="/word/media/([^"]+)"[^>]*>\s*<pkg:binaryData>([^<]+)</pkg:binaryData>',
        xml, re.S)
    n = 0
    for m in parts:
        name, b64 = m.group(1), re.sub(r"\s+", "", m.group(2))
        data = base64.b64decode(b64)
        (out_dir / name).write_bytes(data)
        n += 1
        size = ""
        try:
            from PIL import Image
            with Image.open(out_dir / name) as im:
                size = f"{im.width}x{im.height}px"
        except Exception:
            pass
        print(f"{name}\t{len(data)//1024} KB\t{size}")

    print(f"\n{n}개 추출 → {out_dir}")
    if n == 0:
        print("그림이 없거나 XML 구조가 예상과 다름. flat.xml에서 'pkg:name=\"/word/media' 검색해 확인")


if __name__ == "__main__":
    main()
