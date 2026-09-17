"""
docx2md_com_v2.py — Word COM(pywin32)으로 .docx를 열어 agent용 구조화 데이터로 변환하는 배치 변환기 (v2).
LLM 불필요. DRM 연동 Word가 문서를 열어주면 Python이 객체 모델에서 내용을 읽어 평문으로 저장한다.

사용:
    pip install pywin32 pillow
    python docx2md_com_v2.py <입력폴더> <출력폴더> [--limit N] [--no-images] [--doc-id-from filename|title]

출력 (문서당 하위 폴더):
    <출력폴더>/<문서ID>/<문서ID>.md          사람 검수용
    <출력폴더>/<문서ID>/<문서ID>.json        메타 + 블록 배열 (agent 입력)
    <출력폴더>/<문서ID>/<문서ID>.jsonl       블록당 1줄 (스트리밍 인덱싱용)
    <출력폴더>/<문서ID>/media/img_001.png ...
    <출력폴더>/_manifest.jsonl               문서 메타 1줄씩 (인벤토리 입력)
    <출력폴더>/_log.csv

v1 대비 변경: 블록 ID·제목 경로·페이지, 떠 있는 도형 텍스트, 캡션 연결, 타 문서 참조, 문서 메타,
              리스트 계층, 상용구 절 플래그, JSONL, manifest, EMF/WMF→PNG
v2.1: 그림 추출 1순위를 Flat XML(원본 바이트)로 변경 — 웹 페이지 저장의 96dpi 재샘플링 회피
"""
import sys, os, re, json, csv, time, shutil, argparse, pathlib, glob

import win32com.client
import pythoncom

# =========================== 사내 표준서에 맞게 조정하는 부분 ===========================
# 제목 스타일명 → 계층. 사내 자체 스타일이 있으면 여기에 추가 (정규식, 그룹1이 레벨 숫자)
HEADING_PATTERNS = [
    re.compile(r"^(?:제목|Heading)\s*(\d)$", re.I),
    # re.compile(r"^표준서\s*장제목\s*(\d)$"),
]
# 본문에서 찾을 타 문서 번호 패턴 (예: IWP-AB-1234, SOP_1234 등). 사내 번호 체계로 교체.
DOC_NO_RE = re.compile(r"\b(?:IWP|IDMS|PLM|SOP|STD|WI)[-_ ]?[A-Z0-9]{2,}[-_][A-Z0-9-]{2,}\b", re.I)
# 참조를 뜻하는 문맥 (문서번호가 없어도 "○○ 표준서 참조" 형태를 잡기 위함)
REF_CONTEXT_RE = re.compile(r"([가-힣A-Za-z0-9 ]{2,30}(?:표준서|기준서|절차서|지침서|규정))\s*(?:을|를)?\s*(?:참조|준용|따른다|참고)")
# 상용구 절 제목 (추출 대상이 아닌 절). 정규식, 제목 텍스트에 매칭
BOILERPLATE_RE = re.compile(r"(목\s*적|적용\s*범위|용어\s*(의\s*)?정의|개정\s*이력|개정\s*내역|관련\s*문서|참고\s*문서|첨\s*부|부\s*록)")
CAPTION_RE = re.compile(r"^\s*(?:<\s*)?(그림|표|Figure|Table)\s*[\d.-]+\s*[>.:\-]?\s*(.*)$")
# ======================================================================================

WD_FILTERED_HTML = 10
WD_FLAT_XML = 19
WD_STAT_PAGES = 2
WD_ACTIVE_END_PAGE = 3  # Range.Information(3) = wdActiveEndAdjustedPageNumber


def clean(text: str) -> str:
    return text.replace("\r\x07", "").replace("\x07", "").replace("\r", "\n").replace("\x0c", "").strip()


def heading_level(style_name: str):
    s = style_name.strip()
    for pat in HEADING_PATTERNS:
        m = pat.match(s)
        if m:
            return int(m.group(1))
    return None


def md_cell(s: str) -> str:
    return s.replace("|", "\\|").replace("\n", "<br>")


def table_to_grid(tbl):
    grid, max_r, max_c = {}, 0, 0
    for cell in tbl.Range.Cells:
        r, c = cell.RowIndex, cell.ColumnIndex
        grid[(r, c)] = clean(cell.Range.Text)
        max_r, max_c = max(max_r, r), max(max_c, c)
    return [[grid.get((r, c), "") for c in range(1, max_c + 1)] for r in range(1, max_r + 1)]


def grid_to_md(rows):
    if not rows:
        return ""
    w = max(len(r) for r in rows)
    rows = [r + [""] * (w - len(r)) for r in rows]
    out = ["| " + " | ".join(md_cell(c) for c in rows[0]) + " |", "|" + "---|" * w]
    out += ["| " + " | ".join(md_cell(c) for c in r) + " |" for r in rows[1:]]
    return "\n".join(out)


def safe_page(rng):
    try:
        return int(rng.Information(WD_ACTIVE_END_PAGE))
    except Exception:
        return None


def to_png(path: pathlib.Path):
    """EMF/WMF/BMP 등을 PNG로. Windows의 Pillow는 GDI로 EMF/WMF를 읽을 수 있다. 실패하면 원본 유지."""
    if path.suffix.lower() == ".png":
        return path
    try:
        from PIL import Image
        im = Image.open(path)
        im.load()
        png = path.with_suffix(".png")
        im.convert("RGB").save(png)
        path.unlink(missing_ok=True)
        return png
    except Exception as e:
        print(f"    [img] {path.name} → png 실패({e}), 원본 유지")
        return path


def export_images_via_flatxml(doc, media_dir: pathlib.Path):
    """Flat XML(FileFormat=19)로 저장해 Word가 보관한 원본 그림 바이트를 그대로 꺼낸다.
    웹 페이지 저장은 96dpi로 재샘플링하지만 이 경로는 원본 해상도가 유지된다.
    순서는 본문(document.xml)의 r:embed 등장 순서를 rels로 풀어 문서 내 순서와 맞춘다."""
    import base64
    tmp = media_dir / "_flat"
    tmp.mkdir(parents=True, exist_ok=True)
    xml_path = tmp / "doc.xml"
    try:
        doc.SaveAs2(str(xml_path), FileFormat=WD_FLAT_XML)
    except Exception as e:
        print(f"    [img] Flat XML export blocked: {e}")
        shutil.rmtree(tmp, ignore_errors=True)
        return []
    head = xml_path.read_bytes()[:16]
    if not head.lstrip().startswith(b"<"):
        print("    [img] Flat XML output is DRM-wrapped")
        shutil.rmtree(tmp, ignore_errors=True)
        return []
    xml = xml_path.read_text(encoding="utf-8", errors="ignore")

    # 1) media 파트: 이름 → 바이트
    media = {}
    for m in re.finditer(r'<pkg:part pkg:name="/word/media/([^"]+)"[^>]*>\s*<pkg:binaryData>([^<]+)</pkg:binaryData>', xml, re.S):
        media[m.group(1)] = base64.b64decode(re.sub(r"\s+", "", m.group(2)))
    if not media:
        shutil.rmtree(tmp, ignore_errors=True)
        return []

    # 2) rels 파트: rId → media 이름
    rels = {}
    rm = re.search(r'<pkg:part pkg:name="/word/_rels/document.xml.rels".*?</pkg:part>', xml, re.S)
    if rm:
        for r in re.finditer(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="media/([^"]+)"', rm.group(0)):
            rels[r.group(1)] = r.group(2)
        for r in re.finditer(r'<Relationship [^>]*Target="media/([^"]+)"[^>]*Id="([^"]+)"', rm.group(0)):
            rels[r.group(2)] = r.group(1)

    # 3) 본문 파트: 등장 순서대로 r:embed / r:link
    dm = re.search(r'<pkg:part pkg:name="/word/document.xml".*?</pkg:part>', xml, re.S)
    order = []
    if dm:
        for e in re.finditer(r'r:(?:embed|link)="([^"]+)"', dm.group(0)):
            name = rels.get(e.group(1))
            if name and name in media:
                order.append(name)
    if not order:  # rels를 못 풀면 파일명 숫자순으로 대체
        order = sorted(media, key=lambda n: int(re.sub(r"\D", "", n) or 0))

    paths = []
    for i, name in enumerate(order, 1):
        dst = media_dir / f"img_{i:03d}{pathlib.Path(name).suffix.lower()}"
        dst.write_bytes(media[name])
        paths.append(to_png(dst))
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"    [img] flat xml: {len(paths)} images (original bytes)")
    return paths


def export_images_via_html(doc, media_dir: pathlib.Path):
    tmp = media_dir / "_html"
    tmp.mkdir(parents=True, exist_ok=True)
    html_path = str(tmp / "doc.htm")
    try:
        doc.SaveAs2(html_path, FileFormat=WD_FILTERED_HTML)
    except Exception as e:
        print(f"    [img] HTML export blocked: {e}")
        shutil.rmtree(tmp, ignore_errors=True)
        return []
    with open(html_path, "rb") as f:
        head = f.read(16)
    if not head.lstrip().startswith(b"<"):
        print("    [img] HTML output is DRM-wrapped")
        shutil.rmtree(tmp, ignore_errors=True)
        return []
    files = sorted(glob.glob(str(tmp / "doc_files" / "image*")) + glob.glob(str(tmp / "doc.files" / "image*")))
    paths = []
    for i, f in enumerate(files, 1):
        dst = media_dir / f"img_{i:03d}{pathlib.Path(f).suffix.lower()}"
        shutil.move(f, dst)
        paths.append(to_png(dst))
    shutil.rmtree(tmp, ignore_errors=True)
    return paths


def export_images_via_clipboard(doc, media_dir: pathlib.Path):
    try:
        from PIL import ImageGrab
    except ImportError:
        return []
    paths = []
    for i in range(1, doc.InlineShapes.Count + 1):
        try:
            doc.InlineShapes(i).Range.CopyAsPicture()
            time.sleep(0.15)
            im = ImageGrab.grabclipboard()
            if im is None:
                print("    [img] clipboard blocked")
                return paths
            dst = media_dir / f"img_{i:03d}.png"
            im.save(dst)
            paths.append(dst)
        except Exception as e:
            print(f"    [img] shape {i} failed: {e}")
    return paths


def collect_floating_shapes(doc):
    """떠 있는 도형·텍스트박스·그룹의 글자를 앵커 위치와 함께 수집 (순서도 노드명이 여기 있다)."""
    out = []

    def walk(shape):
        try:
            if shape.Type == 6:  # msoGroup
                for i in range(1, shape.GroupItems.Count + 1):
                    walk(shape.GroupItems(i))
                return
            if shape.TextFrame.HasText:
                t = clean(shape.TextFrame.TextRange.Text)
                if t:
                    out.append(t)
        except Exception:
            pass

    for s in doc.Shapes:
        try:
            anchor = s.Anchor.Start
        except Exception:
            anchor = None
        before = len(out)
        walk(s)
        for t in out[before:]:
            yield anchor, t


def doc_metadata(doc, tables_grid_first):
    meta = {}
    try:
        props = doc.BuiltInDocumentProperties
        for key in ("Title", "Subject", "Keywords", "Revision number", "Last save time"):
            try:
                v = props(key).Value
                if v:
                    meta[key.lower().replace(" ", "_")] = str(v)
            except Exception:
                pass
    except Exception:
        pass
    # 첫 표가 문서번호/개정이력 표인 경우가 많음
    if tables_grid_first:
        flat = " ".join(" ".join(r) for r in tables_grid_first)
        if re.search(r"(문서\s*번호|개정|Rev)", flat):
            meta["header_table"] = tables_grid_first
            m = DOC_NO_RE.search(flat)
            if m:
                meta["doc_no"] = m.group(0)
    return meta


def convert_one(word, src: pathlib.Path, out_root: pathlib.Path, want_images: bool, id_mode: str):
    doc = word.Documents.Open(str(src.resolve()), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    try:
        pages = doc.ComputeStatistics(WD_STAT_PAGES)
        tables = [(t.Range.Start, t.Range.End, t) for t in doc.Tables]
        inline_pos = sorted(s.Range.Start for s in doc.InlineShapes)
        floating = sorted([(a if a is not None else 0, t) for a, t in collect_floating_shapes(doc)])
        first_grid = table_to_grid(tables[0][2]) if tables else None
        meta = doc_metadata(doc, first_grid)

        doc_id = src.stem
        if id_mode == "title" and meta.get("doc_no"):
            doc_id = re.sub(r"[^\w\-]", "_", meta["doc_no"])
        out_dir = out_root / doc_id
        media_dir = out_dir / "media"
        out_dir.mkdir(parents=True, exist_ok=True)

        img_paths = []
        if want_images and inline_pos:
            media_dir.mkdir(exist_ok=True)
            img_paths = (export_images_via_flatxml(doc, media_dir)
                         or export_images_via_html(doc, media_dir)
                         or export_images_via_clipboard(doc, media_dir))

        blocks, md = [], []
        seq = 0
        heading_path = []          # [(level, text)]
        section_no = []            # 번호 경로 문자열용
        cur_boiler = False
        emitted_tables = set()
        img_idx, fl_idx = 0, 0
        refs = set()
        last_block = None

        def hpath():
            return [t for _, t in heading_path]

        def new_block(b, page):
            nonlocal seq, last_block
            seq += 1
            b.update({"id": f"{doc_id}#{seq:04d}", "heading_path": hpath(),
                      "section": " / ".join(hpath()), "page": page, "boilerplate": cur_boiler})
            blocks.append(b)
            last_block = b
            return b

        def scan_refs(text):
            for m in DOC_NO_RE.finditer(text):
                refs.add(m.group(0))
            for m in REF_CONTEXT_RE.finditer(text):
                refs.add(m.group(1).strip())

        for p in doc.Paragraphs:
            rng = p.Range
            start, end = rng.Start, rng.End
            page = safe_page(rng)

            # 이 문단 앞에 앵커된 떠 있는 도형 텍스트를 먼저 흘려보냄
            while fl_idx < len(floating) and floating[fl_idx][0] < end:
                t = floating[fl_idx][1]
                b = new_block({"type": "shape_text", "text": t}, page)
                md.append(f"> [도형] {t}")
                md.append("")
                scan_refs(t)
                fl_idx += 1

            in_table = next(((ti, t) for ti, (ts, te, t) in enumerate(tables) if ts <= start < te), None)
            if in_table:
                ti, t = in_table
                if ti not in emitted_tables:
                    emitted_tables.add(ti)
                    rows = table_to_grid(t)
                    b = new_block({"type": "table", "rows": rows, "caption": None}, page)
                    # 바로 앞 문단이 "표 N" 캡션이면 연결
                    if last_block is not None and len(blocks) >= 2 and blocks[-2].get("caption_for") == "table":
                        b["caption"] = blocks[-2]["text"]
                    md.append(grid_to_md(rows))
                    md.append("")
                    scan_refs(" ".join(" ".join(r) for r in rows))
                continue

            text = clean(rng.Text)

            # 문단 안 인라인 그림
            while img_idx < len(inline_pos) and start <= inline_pos[img_idx] < end:
                rel = f"media/{img_paths[img_idx].name}" if img_idx < len(img_paths) else f"media/img_{img_idx+1:03d}.MISSING"
                b = new_block({"type": "image", "path": rel, "caption": None}, page)
                md.append(f"![]({rel})")
                md.append("")
                img_idx += 1

            if not text:
                continue

            style = p.Style.NameLocal
            lvl = heading_level(style)
            try:
                num = p.Range.ListFormat.ListString or ""
                list_level = int(p.Range.ListFormat.ListLevelNumber) if num else 0
            except Exception:
                num, list_level = "", 0

            if lvl:
                heading_path = [(l, t) for l, t in heading_path if l < lvl] + [(lvl, (num + " " + text).strip())]
                cur_boiler = bool(BOILERPLATE_RE.search(text))
                new_block({"type": "heading", "level": lvl, "text": text, "num": num}, page)
                md.append("#" * lvl + " " + (num + " " if num else "") + text)
                md.append("")
                continue

            cap = CAPTION_RE.match(text)
            b = new_block({"type": "paragraph", "text": text, "style": style, "num": num, "list_level": list_level}, page)
            if cap:
                kind = "image" if cap.group(1) in ("그림", "Figure") else "table"
                b["caption_for"] = kind
                # 바로 앞 블록이 그림/표면 캡션을 그쪽에 붙임
                if len(blocks) >= 2 and blocks[-2]["type"] == kind and blocks[-2].get("caption") is None:
                    blocks[-2]["caption"] = text
            scan_refs(text)
            indent = "    " * max(list_level - 1, 0) if num else ""
            md.append(indent + (num + " " if num else "") + text)
            md.append("")

        # 하이퍼링크
        try:
            for h in doc.Hyperlinks:
                a = clean(h.TextToDisplay or "")
                refs.add(f"{a} <{h.Address}>" if h.Address else a)
        except Exception:
            pass

        meta.update({"doc_id": doc_id, "source": str(src), "pages": pages, "tables": len(tables),
                     "images": len(inline_pos), "images_extracted": len(img_paths),
                     "shape_texts": len(floating), "blocks": len(blocks),
                     "references": sorted(refs),
                     "headings": [b["text"] for b in blocks if b["type"] == "heading"]})

        (out_dir / f"{doc_id}.md").write_text("\n".join(md), encoding="utf-8")
        (out_dir / f"{doc_id}.json").write_text(
            json.dumps({"meta": meta, "blocks": blocks}, ensure_ascii=False, indent=1), encoding="utf-8")
        with open(out_dir / f"{doc_id}.jsonl", "w", encoding="utf-8") as jf:
            for b in blocks:
                jf.write(json.dumps({"doc_id": doc_id, **b}, ensure_ascii=False) + "\n")
        return meta
    finally:
        doc.Close(SaveChanges=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src_dir", nargs="?", default=".")
    ap.add_argument("out_dir", nargs="?", default="out")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--no-images", action="store_true")
    ap.add_argument("--doc-id-from", choices=["filename", "title"], default="filename")
    a = ap.parse_args()

    src_root, out_root = pathlib.Path(a.src_dir).resolve(), pathlib.Path(a.out_dir).resolve()
    out_root.mkdir(parents=True, exist_ok=True)
    files = [f for f in sorted(src_root.rglob("*.docx")) if out_root not in f.parents and not f.name.startswith("~$")]
    if a.limit:
        files = files[:a.limit]

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    log_path = out_root / "_log.csv"
    new_log = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8-sig") as lf, \
         open(out_root / "_manifest.jsonl", "a", encoding="utf-8") as mf:
        w = csv.writer(lf)
        if new_log:
            w.writerow(["file", "status", "pages", "tables", "images", "images_extracted", "shape_texts", "headings", "sec", "error"])
        for i, f in enumerate(files, 1):
            if (out_root / f.stem / f"{f.stem}.md").exists():
                continue
            t0 = time.time()
            print(f"[{i}/{len(files)}] {f.name}")
            try:
                m = convert_one(word, f, out_root, not a.no_images, a.doc_id_from)
                w.writerow([f.name, "ok", m["pages"], m["tables"], m["images"], m["images_extracted"],
                            m["shape_texts"], len(m["headings"]), round(time.time() - t0, 1), ""])
                mf.write(json.dumps({k: v for k, v in m.items() if k != "header_table"}, ensure_ascii=False) + "\n")
            except Exception as e:
                w.writerow([f.name, "fail", "", "", "", "", "", "", round(time.time() - t0, 1), str(e)[:200]])
                print(f"    FAIL: {e}")
            lf.flush(); mf.flush()
    word.Quit()


if __name__ == "__main__":
    main()
