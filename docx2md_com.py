"""
docx2md_com.py — Word COM(pywin32)으로 .docx를 열어 Markdown + JSON + 이미지로 변환하는 배치 변환기.
LLM 불필요. DRM 연동 Word가 문서를 열어주면 Python이 객체 모델에서 내용을 읽어 평문으로 저장한다.

사용:
    pip install pywin32 pillow
    python docx2md_com.py <입력폴더> <출력폴더> [--limit N] [--no-images]

출력 (파일명 기준 하위 폴더):
    <출력폴더>/<문서명>/<문서명>.md
    <출력폴더>/<문서명>/<문서명>.json
    <출력폴더>/<문서명>/media/img_001.png ...
    <출력폴더>/_log.csv
"""
import sys, os, re, json, csv, time, shutil, argparse, pathlib, glob

import win32com.client
import pythoncom

# ---- Word 상수 ----
WD_FILTERED_HTML = 10
WD_STAT_PAGES = 2

HEADING_RE = re.compile(r"^(제목|Heading)\s*(\d)$", re.I)


def clean(text: str) -> str:
    return text.replace("\r\x07", "").replace("\x07", "").replace("\r", "\n").strip()


def heading_level(style_name: str):
    m = HEADING_RE.match(style_name.strip())
    return int(m.group(2)) if m else None


def md_escape_cell(s: str) -> str:
    return s.replace("|", "\\|").replace("\n", "<br>")


def table_to_grid(tbl):
    """병합 셀을 견디도록 Range.Cells를 순회해 (row, col) 사전으로 만든다."""
    grid, max_r, max_c = {}, 0, 0
    for cell in tbl.Range.Cells:
        r, c = cell.RowIndex, cell.ColumnIndex
        grid[(r, c)] = clean(cell.Range.Text)
        max_r, max_c = max(max_r, r), max(max_c, c)
    rows = [[grid.get((r, c), "") for c in range(1, max_c + 1)] for r in range(1, max_r + 1)]
    return rows


def grid_to_md(rows):
    if not rows:
        return ""
    w = max(len(r) for r in rows)
    rows = [r + [""] * (w - len(r)) for r in rows]
    out = ["| " + " | ".join(md_escape_cell(c) for c in rows[0]) + " |",
           "|" + "---|" * w]
    for r in rows[1:]:
        out.append("| " + " | ".join(md_escape_cell(c) for c in r) + " |")
    return "\n".join(out)


def export_images_via_html(doc, media_dir: pathlib.Path, base: str):
    """웹 페이지(필터링됨) 저장으로 그림을 폴더로 떨어뜨린다. DRM이 저장을 막으면 [] 반환."""
    tmp = media_dir / "_html"
    tmp.mkdir(parents=True, exist_ok=True)
    html_path = str(tmp / "doc.htm")
    try:
        doc.SaveAs2(html_path, FileFormat=WD_FILTERED_HTML)
    except Exception as e:
        print(f"    [img] HTML export blocked: {e}")
        return []
    with open(html_path, "rb") as f:
        head = f.read(16)
    if not head.lstrip().startswith(b"<"):
        print("    [img] HTML output is DRM-wrapped")
        return []
    files = sorted(glob.glob(str(tmp / "doc_files" / "image*")) +
                   glob.glob(str(tmp / "doc.files" / "image*")))
    paths = []
    for i, f in enumerate(files, 1):
        ext = pathlib.Path(f).suffix.lower()
        dst = media_dir / f"img_{i:03d}{ext}"
        shutil.move(f, dst)
        paths.append(dst)
    shutil.rmtree(tmp, ignore_errors=True)
    return paths


def export_images_via_clipboard(doc, media_dir: pathlib.Path):
    """HTML 저장이 막혔을 때 예비: 그림을 클립보드로 복사해 PIL로 저장. 클립보드가 막히면 None이 온다."""
    try:
        from PIL import ImageGrab
    except ImportError:
        print("    [img] pillow 미설치 — 클립보드 경로 생략")
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


def convert_one(word, src: pathlib.Path, out_root: pathlib.Path, want_images: bool):
    base = src.stem
    out_dir = out_root / base
    media_dir = out_dir / "media"
    out_dir.mkdir(parents=True, exist_ok=True)

    doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    try:
        pages = doc.ComputeStatistics(WD_STAT_PAGES)

        # 표 범위와 그림 위치를 먼저 수집
        tables = [(t.Range.Start, t.Range.End, t) for t in doc.Tables]
        shapes = sorted(s.Range.Start for s in doc.InlineShapes)

        # 그림 파일 추출 (문서 순서 = 파일 순서 가정)
        img_paths = []
        if want_images and shapes:
            media_dir.mkdir(exist_ok=True)
            img_paths = export_images_via_html(doc, media_dir, base)
            if not img_paths:
                img_paths = export_images_via_clipboard(doc, media_dir)

        blocks, md = [], []
        emitted_tables = set()
        img_idx = 0

        for p in doc.Paragraphs:
            start = p.Range.Start
            # 표 안 문단이면 표 전체를 한 번만 출력
            in_table = None
            for ti, (ts, te, t) in enumerate(tables):
                if ts <= start < te:
                    in_table = (ti, t)
                    break
            if in_table:
                ti, t = in_table
                if ti not in emitted_tables:
                    emitted_tables.add(ti)
                    rows = table_to_grid(t)
                    blocks.append({"type": "table", "rows": rows})
                    md.append(grid_to_md(rows))
                    md.append("")
                continue

            text = clean(p.Range.Text)
            # 문단 안의 그림
            while img_idx < len(shapes) and start <= shapes[img_idx] < p.Range.End:
                rel = f"media/{img_paths[img_idx].name}" if img_idx < len(img_paths) else f"media/img_{img_idx+1:03d}.MISSING"
                blocks.append({"type": "image", "path": rel})
                md.append(f"![]({rel})")
                md.append("")
                img_idx += 1
            if not text:
                continue

            style = p.Style.NameLocal
            lvl = heading_level(style)
            try:
                num = p.Range.ListFormat.ListString or ""
            except Exception:
                num = ""
            if lvl:
                blocks.append({"type": "heading", "level": lvl, "text": text, "num": num})
                md.append("#" * lvl + " " + (num + " " if num else "") + text)
            else:
                blocks.append({"type": "paragraph", "text": text, "style": style, "num": num})
                md.append((num + " " if num else "") + text)
            md.append("")

        meta = {"source": str(src), "pages": pages, "tables": len(tables),
                "images": len(shapes), "images_extracted": len(img_paths)}
        (out_dir / f"{base}.md").write_text("\n".join(md), encoding="utf-8")
        (out_dir / f"{base}.json").write_text(
            json.dumps({"meta": meta, "blocks": blocks}, ensure_ascii=False, indent=1), encoding="utf-8")
        return meta
    finally:
        doc.Close(SaveChanges=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src_dir")
    ap.add_argument("out_dir")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--no-images", action="store_true")
    a = ap.parse_args()

    src_root, out_root = pathlib.Path(a.src_dir), pathlib.Path(a.out_dir)
    out_root.mkdir(parents=True, exist_ok=True)
    files = sorted(src_root.rglob("*.docx"))
    if a.limit:
        files = files[:a.limit]

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    log_path = out_root / "_log.csv"
    new_log = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf)
        if new_log:
            w.writerow(["file", "status", "pages", "tables", "images", "images_extracted", "sec", "error"])
        for i, f in enumerate(files, 1):
            if (out_root / f.stem / f"{f.stem}.md").exists():
                continue  # 재실행 시 건너뛰기
            t0 = time.time()
            print(f"[{i}/{len(files)}] {f.name}")
            try:
                m = convert_one(word, f, out_root, not a.no_images)
                w.writerow([f.name, "ok", m["pages"], m["tables"], m["images"], m["images_extracted"],
                            round(time.time() - t0, 1), ""])
            except Exception as e:
                w.writerow([f.name, "fail", "", "", "", "", round(time.time() - t0, 1), str(e)[:200]])
                print(f"    FAIL: {e}")
            lf.flush()
    word.Quit()


if __name__ == "__main__":
    main()
