"""
Convert PPTX to PNG screenshots via PowerPoint (AppleScript) → PDF → PNG (Quartz).
PowerPoint renders OMML formulas and fonts identically to the final output.

Usage:
    python3 tests/take_screenshots.py path/to/file.pptx /tmp/output_dir/
"""
import os
import sys
import re
import subprocess
import tempfile
from pptx import Presentation

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ── Step 1: PPTX → PDF via PowerPoint AppleScript ─────────────────────────────

def _pptx_to_pdf(pptx_path, pdf_path):
    """Export all slides to PDF using Microsoft PowerPoint."""
    pptx_abs = os.path.abspath(pptx_path)
    pdf_abs   = os.path.abspath(pdf_path)
    script = f"""\
tell application "Microsoft PowerPoint"
    open POSIX file "{pptx_abs}"
    delay 3
    save active presentation in POSIX file "{pdf_abs}" as save as PDF
    close active presentation saving no
end tell
"""
    with tempfile.NamedTemporaryFile(suffix=".applescript", mode="w", delete=False) as f:
        f.write(script)
        script_path = f.name
    try:
        result = subprocess.run(
            ["osascript", script_path],
            capture_output=True, text=True, timeout=120,
        )
    finally:
        os.unlink(script_path)
    if result.returncode != 0 or not os.path.exists(pdf_abs):
        raise RuntimeError(
            f"PowerPoint PDF export failed:\n{result.stdout}\n{result.stderr}"
        )


# ── Step 2: PDF → PNG via macOS Quartz (PyObjC) ───────────────────────────────

def _pdf_to_png(pdf_path, out_dir, dpi=144):
    """Render each PDF page to PNG using macOS Quartz. Returns sorted PNG list."""
    from Foundation import NSURL
    from Quartz import (
        PDFDocument,
        CGColorSpaceCreateDeviceRGB,
        CGBitmapContextCreate,
        CGContextDrawPDFPage,
        CGBitmapContextCreateImage,
        CGImageDestinationCreateWithURL,
        CGImageDestinationAddImage,
        CGImageDestinationFinalize,
        CGContextSetRGBFillColor,
        CGContextFillRect,
        CGContextScaleCTM,
        CGRectMake,
        kCGBitmapByteOrderDefault,
        kCGImageAlphaPremultipliedLast,
    )

    os.makedirs(out_dir, exist_ok=True)
    url = NSURL.fileURLWithPath_(os.path.abspath(pdf_path))
    doc = PDFDocument.alloc().initWithURL_(url)
    if doc is None:
        raise RuntimeError(f"Cannot open PDF: {pdf_path}")

    scale = dpi / 72.0
    paths = []
    for i in range(doc.pageCount()):
        page   = doc.pageAtIndex_(i)
        bounds = page.boundsForBox_(0)
        w_px   = int(bounds.size.width  * scale)
        h_px   = int(bounds.size.height * scale)

        cs  = CGColorSpaceCreateDeviceRGB()
        ctx = CGBitmapContextCreate(
            None, w_px, h_px, 8, w_px * 4, cs,
            kCGBitmapByteOrderDefault | kCGImageAlphaPremultipliedLast,
        )
        CGContextSetRGBFillColor(ctx, 1.0, 1.0, 1.0, 1.0)
        CGContextFillRect(ctx, CGRectMake(0, 0, w_px, h_px))
        CGContextScaleCTM(ctx, scale, scale)
        CGContextDrawPDFPage(ctx, page.pageRef())

        img  = CGBitmapContextCreateImage(ctx)
        path = os.path.join(out_dir, f"slide_{i+1:03d}.png")
        dest = CGImageDestinationCreateWithURL(
            NSURL.fileURLWithPath_(os.path.abspath(path)), "public.png", 1, None
        )
        CGImageDestinationAddImage(dest, img, None)
        CGImageDestinationFinalize(dest)
        paths.append(path)

    return paths


# ── Representative slide selection ────────────────────────────────────────────

def _has_overflow_font(slide):
    for rpr in slide._element.findall(f".//{{{A_NS}}}rPr"):
        sz = rpr.get("sz")
        if sz and int(sz) <= 1800:
            return True
    return False


def _has_image(slide):
    return bool(slide._element.findall(f".//{{{P_NS}}}pic"))


def select_representative(pptx_path, png_paths):
    """
    Pick representative slides:
    - First occurrence of each unique layout name
    - Any slide with overflow-shrunk font (≤18pt)
    - Any slide with images
    Returns list of (slide_number_1based, layout_name, reason, png_path).
    """
    prs = Presentation(pptx_path)
    seen_layouts = set()
    selected = []

    for i, (slide, png) in enumerate(zip(prs.slides, png_paths), 1):
        layout = slide.slide_layout.name
        reasons = []
        if layout not in seen_layouts:
            seen_layouts.add(layout)
            reasons.append("新版式")
        if _has_overflow_font(slide):
            reasons.append("溢出缩小")
        if _has_image(slide):
            reasons.append("含图片")
        if reasons:
            selected.append((i, layout, "+".join(reasons), png))

    return selected


# ── Public API ────────────────────────────────────────────────────────────────

def run(pptx_path, out_dir):
    """Full pipeline: PPTX → PDF → PNG → representative selection → report."""
    print(f"\n{'='*60}")
    print(f"截图: {os.path.basename(pptx_path)}")

    os.makedirs(out_dir, exist_ok=True)
    pdf_path = os.path.join(out_dir, "_export.pdf")

    print("  Step 1/3: PowerPoint 导出 PDF...")
    _pptx_to_pdf(pptx_path, pdf_path)

    print("  Step 2/3: Quartz 渲染 PDF → PNG...")
    png_paths = _pdf_to_png(pdf_path, out_dir)
    print(f"  共生成 {len(png_paths)} 张截图")

    print("  Step 3/3: 选取代表性页面...")
    slides = select_representative(pptx_path, png_paths)
    print(f"  代表性页面 ({len(slides)} 张):")
    for slide_num, layout, reason, png in slides:
        print(f"    第{slide_num:02d}页 [{reason}] {layout}")
        print(f"    → {png}")

    return slides


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python3 take_screenshots.py <pptx> <out_dir>")
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
