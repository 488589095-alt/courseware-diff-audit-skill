# -*- coding: utf-8 -*-
"""
render_slides.py — 免 PowerPoint/PyObjC 的渲染路径：LibreOffice(soffice) pptx→PDF + PyMuPDF PDF→PNG。

take_screenshots.py 依赖 Microsoft PowerPoint(AppleScript) + PyObjC(Quartz)；
未装 PowerPoint 的 Mac 用本模块替代。run_audit 在 take_screenshots 失败时回退到这里。

⚠️ 字体注意：LibreOffice 缺自定义字体（阿里巴巴普惠体/可口可乐在乎体/汉仪/字魂等）时会**替换字体**渲染，
所以**字体外观不保真**；但**版式/坐标/重叠/溢出/错位是保留的** —— 排版类问题用本渲染足够判断。
接口与 take_screenshots.run 对齐：写出 slide_{i:03d}.png 到 out_dir，返回代表页列表。

用法: python3 render_slides.py <pptx> <out_dir> [--dpi 110]
"""
import argparse
import os
import sys
import subprocess
import tempfile
from pathlib import Path

SKILL_DIR = Path(__file__).parent
sys.path.insert(0, str(SKILL_DIR))

SOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "soffice", "libreoffice",
]


def _find_soffice():
    for c in SOFFICE_CANDIDATES:
        if os.path.sep in c:
            if Path(c).exists():
                return c
        else:
            from shutil import which
            if which(c):
                return c
    raise RuntimeError("未找到 LibreOffice(soffice)；装 PowerPoint 走 take_screenshots，或装 LibreOffice 走本路径")


def pptx_to_pdf(pptx_path, out_dir):
    soffice = _find_soffice()
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    # 独立 profile 避免与正在运行的 LibreOffice 抢锁
    with tempfile.TemporaryDirectory() as prof:
        r = subprocess.run(
            [soffice, f"-env:UserInstallation=file://{prof}", "--headless",
             "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path)],
            capture_output=True, text=True, timeout=300,
        )
    pdf = out_dir / (Path(pptx_path).stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError(f"soffice 转 PDF 失败:\n{r.stdout}\n{r.stderr}")
    return str(pdf)


def pdf_to_pngs(pdf_path, out_dir, dpi=110):
    import fitz
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    paths = []
    for i in range(doc.page_count):
        pix = doc[i].get_pixmap(dpi=dpi)
        p = out_dir / f"slide_{i+1:03d}.png"
        pix.save(str(p))
        paths.append(str(p))
    doc.close()
    return paths


def run(pptx_path, out_dir, dpi=110):
    print(f"\n  [render_slides] LibreOffice 渲染: {Path(pptx_path).name}")
    pdf = pptx_to_pdf(pptx_path, out_dir)
    pngs = pdf_to_pngs(pdf, out_dir, dpi=dpi)
    print(f"  共 {len(pngs)} 页 (字体可能被替换，版式/位置保真)")
    try:
        from take_screenshots import select_representative
        reps = select_representative(pptx_path, pngs)
    except Exception:
        reps = [(i + 1, "?", "页", p) for i, p in enumerate(pngs)]
    return reps


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("out_dir")
    ap.add_argument("--dpi", type=int, default=110)
    a = ap.parse_args()
    reps = run(a.pptx, a.out_dir, a.dpi)
    for num, layout, reason, png in reps:
        print(f"    第{num:02d}页 [{reason}] {layout} → {png}")
