"""
Structural checks based on 课件生成需求汇总.md.
Each check returns a list of (rule, status, detail) tuples.
status is "PASS", "FAIL", or "WARN".
"""
import os
import re
import zipfile
from lxml import etree
from pptx import Presentation

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _slide_roots(pptx_path):
    with zipfile.ZipFile(pptx_path, "r") as z:
        names = sorted(
            [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml", n)],
            key=lambda n: int(re.search(r"\d+", n.split("/")[-1]).group()),
        )
        for name in names:
            yield name, etree.fromstring(z.read(name))


def _layout_name(slide):
    """Return the slide layout name for a python-pptx slide object."""
    return slide.slide_layout.name


def _has_table(root):
    return bool(root.findall(f".//{{{A_NS}}}tbl"))


def _all_ea_fonts(root):
    return [ea.get("typeface", "") for ea in root.findall(f".//{{{A_NS}}}ea")]


def _all_line_spacings(root):
    results = []
    for spcPct in root.findall(f".//{{{A_NS}}}spcPct"):
        results.append(int(spcPct.get("val", "0")))
    return results


def _bare_limupp(root):
    found = []
    for limupp in root.findall(f".//{{{M_NS}}}limUpp"):
        if f"{{{M_NS}}}e" not in {c.tag for c in limupp}:
            found.append(limupp)
    return found


def _zwsp_in_omath(root):
    found = []
    for mt in root.findall(f".//{{{M_NS}}}t"):
        if mt.text and "​" in mt.text:
            found.append(mt)
    return found


# ── Individual checks ────────────────────────────────────────────────────────

def check_wps_patterns(prs, label):
    results = []
    for i, slide in enumerate(prs.slides, 1):
        root = slide._element
        bare = _bare_limupp(root)
        if bare:
            results.append(("WPS冻结:bare_limUpp", "FAIL",
                             f"[{label}] 第{i}页: {len(bare)}个bare m:limUpp"))
        zwsp = _zwsp_in_omath(root)
        if zwsp:
            results.append(("WPS冻结:U+200B_in_oMath", "FAIL",
                             f"[{label}] 第{i}页: {len(zwsp)}个U+200B在m:t中"))
    if not results:
        results.append(("WPS冻结", "PASS", f"[{label}] 无bare limUpp/U+200B"))
    return results


def check_fonts(prs, label):
    results = []
    bad_slides = []
    for i, slide in enumerate(prs.slides, 1):
        root = slide._element
        bad = [tf for tf in _all_ea_fonts(root)
               if tf and tf not in ("微软雅黑", "+mj-ea", "+mn-ea")]
        if bad:
            bad_slides.append(f"第{i}页:{bad[:2]}")
    if bad_slides:
        results.append(("字体:ea_typeface", "FAIL",
                         f"[{label}] ea字体非微软雅黑: " + "; ".join(bad_slides[:5])))
    else:
        results.append(("字体:ea_typeface", "PASS", f"[{label}] 全部ea字体为微软雅黑"))
    return results


def check_line_spacing(prs, label):
    """Line spacing should be 140000 (1.4×) or 120000 (1.2×), not others."""
    results = []
    bad_slides = []
    for i, slide in enumerate(prs.slides, 1):
        root = slide._element
        bad = [v for v in _all_line_spacings(root)
               if v not in (0, 100000, 120000, 140000, 150000)]
        if bad:
            bad_slides.append(f"第{i}页:{bad[:3]}")
    if bad_slides:
        results.append(("行距", "WARN",
                         f"[{label}] 非标准行距值: " + "; ".join(bad_slides[:5])))
    else:
        results.append(("行距", "PASS", f"[{label}] 行距均为标准值"))
    return results


def check_xuexizhinan_empty(prs, label):
    """学习指南 slide must not contain any table."""
    results = []
    for i, slide in enumerate(prs.slides, 1):
        name = _layout_name(slide)
        if "学习指南" in name or "学习目标" in name:
            if _has_table(slide._element):
                results.append(("学习指南:无表格", "FAIL",
                                 f"[{label}] 第{i}页({name})含有表格，应为空白"))
    if not any(r[1] == "FAIL" for r in results):
        results.append(("学习指南:无表格", "PASS", f"[{label}] 学习指南页无表格"))
    return results


def check_zhidaoma_followed_by_directory(prs, label):
    """你知道吗 section must be followed by a 目录页 layout slide."""
    results = []
    layouts = [_layout_name(s) for s in prs.slides]
    for i, name in enumerate(layouts):
        if "你知道吗" in name:
            # Find the next non-你知道吗 slide
            next_slides = [layouts[j] for j in range(i+1, min(i+4, len(layouts)))]
            has_dir = any("目录" in n for n in next_slides)
            if not has_dir:
                results.append(("你知道吗:后跟目录页", "FAIL",
                                 f"[{label}] 第{i+1}页({name})后未找到目录页版式"))
    if not any(r[1] == "FAIL" for r in results):
        results.append(("你知道吗:后跟目录页", "PASS",
                         f"[{label}] 你知道吗后均有目录页"))
    return results


def check_xinzhi_followed_by_blank(prs, label):
    """新知速递/真题速递 must end with a same-layout blank page."""
    results = []
    layouts = [_layout_name(s) for s in prs.slides]
    xinzhi_indices = [i for i, n in enumerate(layouts)
                      if "新知速递" in n or "真题速递" in n]
    if not xinzhi_indices:
        return results  # not in this PPT
    last_idx = max(xinzhi_indices)
    if last_idx + 1 >= len(layouts):
        results.append(("新知速递:后跟空白页", "FAIL",
                         f"[{label}] 新知速递是最后一页，没有跟随空白页"))
    else:
        next_name = layouts[last_idx + 1]
        if "新知速递" not in next_name and "真题速递" not in next_name:
            results.append(("新知速递:后跟空白页", "FAIL",
                             f"[{label}] 新知速递后跟的是'{next_name}'而非同版式空白页"))
        else:
            results.append(("新知速递:后跟空白页", "PASS",
                             f"[{label}] 新知速递后有同版式空白页"))
    return results


def check_page_structure_main(prs, label):
    """正课 fixed page order: 你知道吗→目录→热身→标题→探索→[模块]→本讲巩固→我也会→再见."""
    results = []
    layouts = [_layout_name(s) for s in prs.slides]
    # Check key milestones appear in correct relative order
    milestones = [
        ("你知道吗",    lambda n: "你知道吗" in n),
        ("标题页",      lambda n: "标题" in n and "模块" not in n),
        ("探索任务",    lambda n: "探索" in n),
        ("本讲巩固标题", lambda n: "本讲巩固" in n and "标题" in n),
        ("再见页",      lambda n: "再见" in n or "拜拜" in n),
    ]
    prev_idx = -1
    for name, matcher in milestones:
        idx = next((i for i, n in enumerate(layouts) if matcher(n)), None)
        if idx is None:
            results.append(("正课页面顺序", "WARN", f"[{label}] 未找到'{name}'版式"))
        elif idx < prev_idx:
            results.append(("正课页面顺序", "FAIL",
                             f"[{label}] '{name}'(第{idx+1}页)出现在前序里程碑之前"))
        else:
            prev_idx = idx
    if not any(r[1] in ("FAIL", "WARN") for r in results):
        results.append(("正课页面顺序", "PASS", f"[{label}] 页面顺序里程碑检查通过"))
    return results


def check_page_structure_small(prs, label):
    """小班课 fixed order: 标题→目录→榜上有名→知识回顾→思维挑战→计算比拼→新知速递→本讲巩固→再见."""
    results = []
    layouts = [_layout_name(s) for s in prs.slides]
    milestones = [
        ("标题页",      lambda n: "标题" in n and "模块" not in n and "巩固" not in n),
        ("榜上有名",    lambda n: "榜上有名" in n),
        ("知识回顾",    lambda n: "知识回顾" in n),
        ("本讲巩固标题", lambda n: "本讲巩固" in n),
        ("再见页",      lambda n: "再见" in n or "拜拜" in n),
    ]
    prev_idx = -1
    for name, matcher in milestones:
        idx = next((i for i, n in enumerate(layouts) if matcher(n)), None)
        if idx is None:
            results.append(("小班课页面顺序", "WARN", f"[{label}] 未找到'{name}'版式"))
        elif idx < prev_idx:
            results.append(("小班课页面顺序", "FAIL",
                             f"[{label}] '{name}'(第{idx+1}页)顺序错误"))
        else:
            prev_idx = idx
    if not any(r[1] in ("FAIL", "WARN") for r in results):
        results.append(("小班课页面顺序", "PASS", f"[{label}] 页面顺序里程碑检查通过"))
    return results


# ── Master runner ────────────────────────────────────────────────────────────

def run_all_checks(pptx_path, ppt_type="main"):
    """Run all structural checks. ppt_type='main' or 'small'."""
    label = os.path.basename(pptx_path)
    prs = Presentation(pptx_path)
    results = []
    results += check_wps_patterns(prs, label)
    results += check_fonts(prs, label)
    results += check_line_spacing(prs, label)
    results += check_xuexizhinan_empty(prs, label)
    if ppt_type == "main":
        results += check_zhidaoma_followed_by_directory(prs, label)
        results += check_page_structure_main(prs, label)
    if ppt_type == "small":
        results += check_xinzhi_followed_by_blank(prs, label)
        results += check_page_structure_small(prs, label)
    return results
