# -*- coding: utf-8 -*-
"""Build the publication-grade Word report PRISM_图像去雾大作业报告_精排终稿.docx (A4, 16.6cm text width)."""
from pathlib import Path
import re
import shutil
import subprocess
from PIL import Image, ImageFile
ImageFile.MAXBLOCK = 4 * 1024 * 1024
import docx
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from lxml import etree

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parent
WORKSPACE = PROJECT.parent.parent.parent
TEMPLATE_PATH = WORKSPACE / "ICLR2027-中文模板-终稿版 (1).docx"
ASSETS_DIR = ROOT / "word_assets"
ASSETS_COMP_DIR = ROOT / "word_assets_compressed"
ASSETS_COMP_DIR.mkdir(parents=True, exist_ok=True)

PHOTO_COLLAGES = {
    "cover_teaser.png": 2200,
    "fig1_motivation.png": 2400,
    "fig3_course10_part1.png": 2150,
    "fig4_course10_part2.png": 2150,
    "fig5_course10_zoom.png": 2250,
    "fig6_fattal_qualitative.png": 2250,
    "fig7_train_boundary.png": 2250,
}


def prepare_word_image(src_path: Path) -> Path:
    """Return an optimized high-DPI image (4:4:4 Q92 JPEG for photo collages, RGB PNG for charts)."""
    name = src_path.name
    im = Image.open(src_path)
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        bg = Image.new("RGB", im.size, (255, 255, 255))
        rgba = im.convert("RGBA")
        bg.paste(rgba, mask=rgba.split()[3])
        im_rgb = bg
    else:
        im_rgb = im.convert("RGB")

    if name in PHOTO_COLLAGES:
        max_w = PHOTO_COLLAGES[name]
        if im_rgb.width > max_w:
            new_h = max(1, round(im_rgb.height * (max_w / im_rgb.width)))
            im_rgb = im_rgb.resize((max_w, new_h), Image.Resampling.LANCZOS)
        out_path = ASSETS_COMP_DIR / f"{src_path.stem}.jpg"
        im_rgb.save(
            out_path,
            format="JPEG",
            quality=92,
            subsampling=0,
            optimize=True,
            dpi=(300, 300),
        )
        return out_path

    out_path = ASSETS_COMP_DIR / name
    im_rgb.save(out_path, format="PNG", optimize=True, dpi=(300, 300))
    return out_path

OUT_REFINED_WS = WORKSPACE / "PRISM_图像去雾大作业报告_精排终稿.docx"
OUT_REFINED_LOCAL = ROOT / "PRISM_图像去雾大作业报告_精排终稿.docx"
OUT_FINAL_WS = WORKSPACE / "PRISM_图像去雾大作业报告_终稿.docx"
OUT_FINAL_LOCAL = ROOT / "PRISM_图像去雾大作业报告_终稿.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# A4 (21.0 cm) - Left/Right margins (2.2 cm * 2) = 16.6 cm = 6.5354 in = 9411 dxa
BODY_WIDTH_DXA = 9410


# ==============================================================================
# 1. Native OMML (Office Math Markup Language) Builders (Strict Schema Order)
# ==============================================================================
def m_el(tag: str):
    return etree.Element(f"{{{M_NS}}}{tag}")


def w_el(tag: str):
    return etree.Element(f"{{{W_NS}}}{tag}")


def m_run(text: str, italic: bool = True, bold: bool = False):
    """Create an <m:r> math run in Cambria Math."""
    r = m_el("r")
    if not italic or bold:
        m_rPr = m_el("rPr")
        sty = m_el("sty")
        if bold and italic:
            sty.set(f"{{{M_NS}}}val", "bi")
        elif bold and not italic:
            sty.set(f"{{{M_NS}}}val", "b")
        else:
            sty.set(f"{{{M_NS}}}val", "p")
        m_rPr.append(sty)
        r.append(m_rPr)
    w_rPr = w_el("rPr")
    rFonts = w_el("rFonts")
    rFonts.set(f"{{{W_NS}}}ascii", "Cambria Math")
    rFonts.set(f"{{{W_NS}}}hAnsi", "Cambria Math")
    w_rPr.append(rFonts)
    r.append(w_rPr)
    t = m_el("t")
    t.text = text
    r.append(t)
    return r


def m_sub(base_elems, sub_elems):
    sSub = m_el("sSub")
    e = m_el("e")
    for el in base_elems:
        e.append(el)
    sub = m_el("sub")
    for el in sub_elems:
        sub.append(el)
    sSub.append(e)
    sSub.append(sub)
    return sSub


def m_sup(base_elems, sup_elems):
    sSup = m_el("sSup")
    e = m_el("e")
    for el in base_elems:
        e.append(el)
    sup = m_el("sup")
    for el in sup_elems:
        sup.append(el)
    sSup.append(e)
    sSup.append(sup)
    return sSup


def m_subsup(base_elems, sub_elems, sup_elems):
    sSubSup = m_el("sSubSup")
    e = m_el("e")
    for el in base_elems:
        e.append(el)
    sub = m_el("sub")
    for el in sub_elems:
        sub.append(el)
    sup = m_el("sup")
    for el in sup_elems:
        sup.append(el)
    sSubSup.append(e)
    sSubSup.append(sub)
    sSubSup.append(sup)
    return sSubSup


def m_frac(num_elems, den_elems):
    f = m_el("f")
    fPr = m_el("fPr")
    t_el = m_el("type")
    t_el.set(f"{{{M_NS}}}val", "bar")
    fPr.append(t_el)
    f.append(fPr)
    num = m_el("num")
    for el in num_elems:
        num.append(el)
    den = m_el("den")
    for el in den_elems:
        den.append(el)
    f.append(num)
    f.append(den)
    return f


def m_delim(inner_elems, beg="(", end=")", sep=","):
    """Create an <m:d> delimiter element matching the template's exact <m:dPr> structure."""
    d = m_el("d")
    dPr = m_el("dPr")
    b_el = m_el("begChr")
    b_el.set(f"{{{M_NS}}}val", beg)
    dPr.append(b_el)
    s_el = m_el("sepChr")
    s_el.set(f"{{{M_NS}}}val", sep)
    dPr.append(s_el)
    e_el = m_el("endChr")
    e_el.set(f"{{{M_NS}}}val", end)
    dPr.append(e_el)
    d.append(dPr)
    e = m_el("e")
    for el in inner_elems:
        e.append(el)
    d.append(e)
    return d


def m_nary(chr_val, sub_elems, sup_elems, body_elems):
    """Create an <m:nary> summation element with explicit lower and upper limits."""
    nary = m_el("nary")
    naryPr = m_el("naryPr")
    c = m_el("chr")
    c.set(f"{{{M_NS}}}val", chr_val)
    limLoc = m_el("limLoc")
    limLoc.set(f"{{{M_NS}}}val", "undOvr")
    naryPr.append(c)
    naryPr.append(limLoc)
    nary.append(naryPr)
    sub = m_el("sub")
    for el in sub_elems:
        sub.append(el)
    sup = m_el("sup")
    if sup_elems:
        for el in sup_elems:
            sup.append(el)
    else:
        sup.append(m_run(" ", italic=False))
    e = m_el("e")
    for el in body_elems:
        e.append(el)
    nary.append(sub)
    nary.append(sup)
    nary.append(e)
    return nary


def m_limlow(base_elems, lim_elems):
    limLow = m_el("limLow")
    e = m_el("e")
    for el in base_elems:
        e.append(el)
    lim = m_el("lim")
    for el in lim_elems:
        lim.append(el)
    limLow.append(e)
    limLow.append(lim)
    return limLow


def add_omml_equation(doc, math_elements, eq_num: int):
    """Insert a 2-column borderless equation table: Left (90% width) centered formula, Right (10% width) right-aligned (N)."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Normal Table"
    tblPr = tbl._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith(("tblW", "tblLayout", "tblCellMar", "tblLook")):
            tblPr.remove(child)
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{BODY_WIDTH_DXA}" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>'
        '<w:top w:w="20" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        '<w:bottom w:w="20" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
        '</w:tblCellMar>'
    ))
    tblPr.append(parse_xml(
        f'<w:tblLook {nsdecls("w")} w:val="0000" w:firstRow="0" w:lastRow="0" '
        'w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
    ))

    w_left = int(round(BODY_WIDTH_DXA * 0.90))   # 8469 dxa (90%)
    w_right = BODY_WIDTH_DXA - w_left            # 941 dxa (10%)
    widths = [w_left, w_right]
    row = tbl.rows[0]
    row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    for idx, w_dxa in enumerate(widths):
        cell = row.cells[idx]
        tcPr = cell._tc.get_or_add_tcPr()
        for child in list(tcPr):
            if child.tag.endswith("tcW"):
                tcPr.remove(child)
        tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>'))
        tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

    # Left Cell (90%): Centered OMML Equation
    p_math = row.cells[0].paragraphs[0]
    p_math.style = "公式"
    p_math.paragraph_format.first_line_indent = Pt(0)
    p_math.paragraph_format.space_before = Pt(2)
    p_math.paragraph_format.space_after = Pt(2)
    oMathPara = m_el("oMathPara")
    oMathParaPr = m_el("oMathParaPr")
    jc = m_el("jc")
    jc.set(f"{{{M_NS}}}val", "center")
    oMathParaPr.append(jc)
    oMathPara.append(oMathParaPr)
    oMath = m_el("oMath")
    for el in math_elements:
        oMath.append(el)
    oMathPara.append(oMath)
    p_math._p.append(oMathPara)

    # Right Cell (10%): Right-aligned Equation Number
    p_num = row.cells[1].paragraphs[0]
    p_num.style = "公式编号"
    p_num.paragraph_format.first_line_indent = Pt(0)
    p_num.paragraph_format.space_before = Pt(2)
    p_num.paragraph_format.space_after = Pt(2)
    p_num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_num = p_num.add_run(f"({eq_num})")
    r_num.font.name = "Times New Roman"
    r_num.font.size = Pt(10.5)


# ==============================================================================
# 2. Equations Catalog (Eq. 1 to Eq. 8)
# ==============================================================================
def build_eq1_scattering():
    # I(x) = J(x) t(x) + A (1 - t(x)),   t(x) = exp(-\beta d(x))
    return [
        m_run("I"), m_delim([m_run("x")]),
        m_run(" = ", italic=False),
        m_run("J"), m_delim([m_run("x")]),
        m_run("t"), m_delim([m_run("x")]),
        m_run(" + ", italic=False),
        m_run("A"), m_delim([m_run("1", italic=False), m_run(" − ", italic=False), m_run("t"), m_delim([m_run("x")])]),
        m_run(",    ", italic=False),
        m_run("t"), m_delim([m_run("x")]),
        m_run(" = ", italic=False),
        m_sup([m_run("e")], [m_run("−", italic=False), m_run("β"), m_run("d"), m_delim([m_run("x")])]),
    ]


def build_eq2_conv_lora():
    # W' = W + \gamma g R^{-1}(B A),  \gamma = \alpha_{LoRA}/r,  g = \sigma(u_{gate})
    return [
        m_sup([m_run("W")], [m_run("′", italic=False)]),
        m_run(" = ", italic=False),
        m_run("W"),
        m_run(" + ", italic=False),
        m_run("γ"), m_run("g"),
        m_sup([m_run("ℛ")], [m_run("−1", italic=False)]),
        m_delim([m_run("B"), m_run("A")]),
        m_run(",    ", italic=False),
        m_run("γ"), m_run(" = ", italic=False),
        m_frac([m_sub([m_run("α")], [m_run("LoRA", italic=False)])], [m_run("r")]),
        m_run(",    ", italic=False),
        m_run("g"), m_run(" = ", italic=False),
        m_run("σ"), m_delim([m_sub([m_run("u")], [m_run("gate", italic=False)])]),
    ]


def build_eq3_covariance():
    # X_c = X - (1/N) 1 1^T X,   \Sigma = \frac{1}{N-1} X_c^T X_c = V \Lambda V^T
    return [
        m_sub([m_run("X")], [m_run("c", italic=False)]),
        m_run(" = ", italic=False),
        m_run("X"),
        m_run(" − ", italic=False),
        m_frac([m_run("1", italic=False)], [m_run("N")]),
        m_run("1", italic=False, bold=True),
        m_sup([m_run("1", italic=False, bold=True)], [m_run("⊤", italic=False)]),
        m_run("X"),
        m_run(",    ", italic=False),
        m_run("Σ", italic=False),
        m_run(" = ", italic=False),
        m_frac([m_run("1", italic=False)], [m_run("N"), m_run(" − ", italic=False), m_run("1", italic=False)]),
        m_subsup([m_run("X")], [m_run("c", italic=False)], [m_run("⊤", italic=False)]),
        m_sub([m_run("X")], [m_run("c", italic=False)]),
        m_run(" = ", italic=False),
        m_run("V"), m_run("Λ", italic=False),
        m_sup([m_run("V")], [m_run("⊤", italic=False)]),
    ]


def build_eq4_conv_eva():
    # A_0 = D V_d^T,  B_0 = 0,  D_{j,j} = ||A_{rand, j}||_2 / (||(V_d^T)_j||_2 + \varepsilon) => \Delta W_0 = 0
    return [
        m_sub([m_run("A")], [m_run("0", italic=False)]),
        m_run(" = ", italic=False),
        m_run("D"),
        m_subsup([m_run("V")], [m_run("d")], [m_run("⊤", italic=False)]),
        m_run(",   ", italic=False),
        m_sub([m_run("B")], [m_run("0", italic=False)]),
        m_run(" = ", italic=False),
        m_run("0", italic=False),
        m_run(",   ", italic=False),
        m_sub([m_run("D")], [m_run("j,j")]),
        m_run(" = ", italic=False),
        m_frac(
            [
                m_run("‖", italic=False),
                m_sub([m_run("A")], [m_run("rand,j", italic=False)]),
                m_sub([m_run("‖", italic=False)], [m_run("2", italic=False)]),
            ],
            [
                m_run("‖", italic=False),
                m_sub([m_delim([m_subsup([m_run("V")], [m_run("d")], [m_run("⊤", italic=False)])])], [m_run("j")]),
                m_sub([m_run("‖", italic=False)], [m_run("2", italic=False)]),
                m_run(" + ", italic=False),
                m_run("ε"),
            ],
        ),
        m_run("  ⟹  ", italic=False),
        m_sub([m_run("Δ", italic=False), m_run("W")], [m_run("0", italic=False)]),
        m_run(" = ", italic=False),
        m_run("0", italic=False),
    ]


def build_eq5_h2c_loss():
    # L_{H2C}(J, I) = 1 - \frac{<\Delta e_{img}(J, I), \Delta e_{txt}>}{||\Delta e_{img}(J, I)||_2 ||\Delta e_{txt}||_2}
    return [
        m_sub([m_run("ℒ")], [m_run("H2C", italic=False)]),
        m_delim([m_run("J"), m_run(", ", italic=False), m_run("I")]),
        m_run(" = ", italic=False),
        m_run("1", italic=False),
        m_run(" − ", italic=False),
        m_frac(
            [
                m_delim(
                    [
                        m_sub([m_run("Δ", italic=False), m_run("e")], [m_run("img", italic=False)]),
                        m_delim([m_run("J"), m_run(", ", italic=False), m_run("I")]),
                        m_run(", ", italic=False),
                        m_sub([m_run("Δ", italic=False), m_run("e")], [m_run("txt", italic=False)]),
                    ],
                    "⟨",
                    "⟩",
                )
            ],
            [
                m_run("‖", italic=False),
                m_sub([m_run("Δ", italic=False), m_run("e")], [m_run("img", italic=False)]),
                m_delim([m_run("J"), m_run(", ", italic=False), m_run("I")]),
                m_sub([m_run("‖", italic=False)], [m_run("2", italic=False)]),
                m_run(" ⋅ ", italic=False),
                m_run("‖", italic=False),
                m_sub([m_run("Δ", italic=False), m_run("e")], [m_run("txt", italic=False)]),
                m_sub([m_run("‖", italic=False)], [m_run("2", italic=False)]),
            ],
        ),
    ]


def build_eq6_snapshot_eval():
    # h(S) = \frac{1}{|D_{sel}|} \sum_{i=1}^{|D_{sel}|} L_{H2C}(f_{\theta_5}(x_i; S)),  p(S) = \frac{1}{|D_{sel}|} \sum_{i=1}^{|D_{sel}|} Q_{MUSIQ}(\Pi_{deploy}(f_{\theta_5}(x_i; S)))
    return [
        m_run("h"), m_delim([m_run("𝒮")]),
        m_run(" = ", italic=False),
        m_frac([m_run("1", italic=False)], [m_delim([m_sub([m_run("𝒟")], [m_run("sel", italic=False)])], "|", "|")]),
        m_nary(
            "∑",
            [m_run("x"), m_run("∈", italic=False), m_sub([m_run("𝒟")], [m_run("sel", italic=False)])],
            [m_delim([m_sub([m_run("𝒟")], [m_run("sel", italic=False)])], "|", "|")],
            [
                m_sub([m_run("ℒ")], [m_run("H2C", italic=False)]),
                m_delim([
                    m_sub([m_run("f")], [m_sub([m_run("θ")], [m_run("5", italic=False)])]),
                    m_delim([m_run("x"), m_run("; ", italic=False), m_run("𝒮")]),
                ]),
            ],
        ),
        m_run(",   ", italic=False),
        m_run("p"), m_delim([m_run("𝒮")]),
        m_run(" = ", italic=False),
        m_frac([m_run("1", italic=False)], [m_delim([m_sub([m_run("𝒟")], [m_run("sel", italic=False)])], "|", "|")]),
        m_nary(
            "∑",
            [m_run("x"), m_run("∈", italic=False), m_sub([m_run("𝒟")], [m_run("sel", italic=False)])],
            [m_delim([m_sub([m_run("𝒟")], [m_run("sel", italic=False)])], "|", "|")],
            [
                m_sub([m_run("Q")], [m_run("MUSIQ", italic=False)]),
                m_delim([
                    m_sub([m_run("Π", italic=False)], [m_run("deploy", italic=False)]),
                    m_delim([
                        m_sub([m_run("f")], [m_sub([m_run("θ")], [m_run("5", italic=False)])]),
                        m_delim([m_run("x"), m_run("; ", italic=False), m_run("𝒮")]),
                    ]),
                ]),
            ],
        ),
    ]


def build_eq7_prism_rule():
    # S^\star = \arg\min_{S \in \Omega_K} h(S)   s.t.   p(S) \ge \tau,   \tau = p(S_{ref})
    return [
        m_sup([m_run("𝒮")], [m_run("⋆", italic=False)]),
        m_run(" = ", italic=False),
        m_limlow([m_run("arg min", italic=False)], [m_run("𝒮"), m_run("∈", italic=False), m_sub([m_run("Ω", italic=False)], [m_run("K")])]),
        m_run(" "),
        m_run("h"), m_delim([m_run("𝒮")]),
        m_run("    s.t.    ", italic=False),
        m_run("p"), m_delim([m_run("𝒮")]),
        m_run(" ≥ ", italic=False),
        m_run("τ"),
        m_run(",    ", italic=False),
        m_run("τ"),
        m_run(" = ", italic=False),
        m_run("p"), m_delim([m_sub([m_run("𝒮")], [m_run("ref", italic=False)])]),
    ]


def build_eq8_reparam_merge():
    # f_{merge}(X) = (W + \gamma g R^{-1}(B A)) * X + b = W' * X + b
    return [
        m_sub([m_run("f")], [m_run("merge", italic=False)]),
        m_delim([m_run("X")]),
        m_run(" = ", italic=False),
        m_delim([
            m_run("W"),
            m_run(" + ", italic=False),
            m_run("γ"), m_run("g"),
            m_sup([m_run("ℛ")], [m_run("−1", italic=False)]),
            m_delim([m_run("B"), m_run("A")]),
        ]),
        m_run(" ∗ ", italic=False),
        m_run("X"),
        m_run(" + ", italic=False),
        m_run("b"),
        m_run(" = ", italic=False),
        m_sup([m_run("W")], [m_run("′", italic=False)]),
        m_run(" ∗ ", italic=False),
        m_run("X"),
        m_run(" + ", italic=False),
        m_run("b"),
    ]


# ==============================================================================
# 3. Typography, Headings, Tables (with vMerge), Code Listings & Figures
# ==============================================================================
def set_run_fonts(run, cn_font="宋体", en_font="Times New Roman", size_pt=10.5, bold=False, italic=False):
    run.font.name = en_font
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), en_font)
    rFonts.set(qn("w:hAnsi"), en_font)
    rFonts.set(qn("w:eastAsia"), cn_font)


def add_hyperlink(p, url: str, text: str | None = None, size_pt: float = 10.5, bold: bool = False):
    """Insert a genuine clickable OOXML hyperlink styled in academic navy blue."""
    r_id = p.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = parse_xml(
        f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
    )
    rPr.append(rFonts)
    sz = int(round(size_pt * 2))
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz}"/>'))
    rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{sz}"/>'))
    rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="1D4ED8"/>'))
    rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))
    if bold:
        rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
    new_run.append(rPr)
    t_el = OxmlElement("w:t")
    t_el.text = text or url
    new_run.append(t_el)
    hyperlink.append(new_run)
    p._p.append(hyperlink)


def add_formatted_text(p, text: str, size_pt: float = 10.5, base_bold: bool = False):
    """Parse **bold**, __underline__, `code`, _{sub}, ^{sup}, [text](url), bare URLs, and normal text."""
    pattern = re.compile(
        r"(\*\*.*?\*\*|__.*?__|`.+?`|\^\{.*?\}|_\{.*?\}|\[[^\]]+\]\(https?://[^\)]+\)|https?://[^\s，。；（）\)]+)"
    )
    tokens = pattern.split(text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            inner = tok[2:-2]
            add_formatted_text(p, inner, size_pt=size_pt, base_bold=True)
        elif tok.startswith("__") and tok.endswith("__"):
            r = p.add_run(tok[2:-2])
            set_run_fonts(r, size_pt=size_pt, bold=base_bold)
            r.underline = True
        elif tok.startswith("`") and tok.endswith("`"):
            r = p.add_run(tok[1:-1])
            set_run_fonts(r, cn_font="宋体", en_font="Consolas", size_pt=max(8.5, size_pt - 1.0), bold=base_bold)
        elif tok.startswith("_{") and tok.endswith("}"):
            r = p.add_run(tok[2:-1])
            set_run_fonts(r, size_pt=size_pt, bold=base_bold)
            r.font.subscript = True
        elif tok.startswith("^{") and tok.endswith("}"):
            r = p.add_run(tok[2:-1])
            set_run_fonts(r, size_pt=size_pt, bold=base_bold)
            r.font.superscript = True
        elif tok.startswith("[") and "](" in tok and tok.endswith(")"):
            m_link = re.match(r"^\[([^\]]+)\]\((https?://[^\)]+)\)$", tok)
            if m_link:
                add_hyperlink(p, m_link.group(2), m_link.group(1), size_pt=size_pt, bold=base_bold)
            else:
                r = p.add_run(tok)
                set_run_fonts(r, size_pt=size_pt, bold=base_bold)
        elif tok.startswith("http://") or tok.startswith("https://"):
            add_hyperlink(p, tok, tok, size_pt=size_pt, bold=base_bold)
        else:
            r = p.add_run(tok)
            set_run_fonts(r, size_pt=size_pt, bold=base_bold)


def set_outline_level(paragraph, level_0based: int):
    """Assign Word XML outlineLvl (0=Level 1, 1=Level 2) so native TOC and Navigation Pane pick it up."""
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith("outlineLvl"):
            pPr.remove(child)
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{level_0based}"/>'))


def ensure_toc_styles(doc):
    """Ensure clean 宋体/TNR formatting with dot leaders for Word TOC 1 and TOC 2 styles."""
    specs = [
        ("TOC 1", "TOC1", 10.5, True, 0.0, 4.0, 2.0, 17.0),
        ("TOC 2", "TOC2", 10.0, False, 18.0, 1.0, 1.0, 15.5),
    ]
    for st_name, st_id, sz_pt, is_bold, indent_pt, sb_pt, sa_pt, ls_pt in specs:
        style = None
        for s in doc.styles:
            if s.name.lower() == st_name.lower() or s.style_id.lower() == st_id.lower():
                style = s
                break
        if style is None:
            style = doc.styles.add_style(st_name, WD_STYLE_TYPE.PARAGRAPH)
        pf = style.paragraph_format
        pf.left_indent = Pt(indent_pt)
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(sb_pt)
        pf.space_after = Pt(sa_pt)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(ls_pt)
        # Configure font on style rPr
        rPr = style._element.get_or_add_rPr()
        for tag_name in ("rFonts", "b", "bCs", "sz", "szCs", "color"):
            for child in list(rPr):
                if child.tag.endswith(tag_name):
                    rPr.remove(child)
        rPr.append(parse_xml(
            f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>'
        ))
        if is_bold:
            rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
        else:
            rPr.append(parse_xml(f'<w:b {nsdecls("w")} w:val="0"/>'))
        half_pt = int(round(sz_pt * 2))
        rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{half_pt}"/>'))
        rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{half_pt}"/>'))
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="0F172A"/>'))


def add_toc_field(doc):
    """Insert a Word-native interactive Table of Contents (levels 1-2 with hyperlinks & outline levels)."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    r_begin = p.add_run()
    r_begin._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin" w:dirty="true"/>'))
    r_instr = p.add_run()
    r_instr._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText>'))
    r_sep = p.add_run()
    r_sep._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
    r_ph = p.add_run("请在 Word 中右键“更新域”以刷新目录页码。")
    set_run_fonts(r_ph, size_pt=10.0)
    r_end = p.add_run()
    r_end._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))
    return p


def add_h1(doc, numbered_title: str):
    """Level-1 Heading: 四号宋体加粗 (14pt), explicit number, space_before=12pt, space_after=6pt."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.keep_with_next = True
    set_outline_level(p, 0)
    r = p.add_run(numbered_title)
    set_run_fonts(r, cn_font="宋体", en_font="Times New Roman", size_pt=14.0, bold=True)
    return p


def add_h2(doc, numbered_title: str):
    """Level-2 Heading: 小四号宋体加粗 (12pt), explicit number, space_before=8pt, space_after=4pt."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(3.5)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.keep_with_next = True
    set_outline_level(p, 1)
    r = p.add_run(numbered_title)
    set_run_fonts(r, cn_font="宋体", en_font="Times New Roman", size_pt=12.0, bold=True)
    return p


def add_body_p(doc, text: str, run_in_title: str = ""):
    """Body paragraph: 五号宋体 (10.5pt) + Times New Roman, fixed 18pt line spacing, 2-char first-line indent."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(21.0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(1.8)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(18.0)
    if run_in_title:
        r_title = p.add_run(run_in_title)
        set_run_fonts(r_title, size_pt=10.5, bold=True)
    add_formatted_text(p, text, size_pt=10.5)
    return p


VECTOR_SVG_STEMS = {
    "fig_layer_sensitivity_heatmap",
    "fig2_landscape",
    "fig_course10_radar",
    "fig_efficiency_stability_bubble",
}
PENDING_SVG_INSERTS = []


def add_figure_block(doc, img_path: Path, width_cm: float, caption_text: str):
    svg_candidate = ASSETS_DIR / f"{img_path.stem}.svg"
    p_img = doc.add_paragraph(style="Normal")
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.first_line_indent = Pt(0)
    p_img.paragraph_format.space_before = Pt(3)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p_img.paragraph_format.keep_with_next = True
    if img_path.stem in VECTOR_SVG_STEMS and svg_candidate.exists():
        tok = f"<<VEC_SVG_{img_path.stem}>>"
        p_img.add_run(tok)
        PENDING_SVG_INSERTS.append((tok, str(svg_candidate), float(width_cm)))
    else:
        opt_img = prepare_word_image(img_path)
        p_img.add_run().add_picture(str(opt_img), width=Cm(width_cm))

    p_cap = doc.add_paragraph(style="Normal")
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(caption_text) < 48 else WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cap.paragraph_format.first_line_indent = Pt(0)
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(5)
    p_cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_cap.paragraph_format.line_spacing = Pt(13.5)
    add_formatted_text(p_cap, caption_text, size_pt=9.0)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith("tcBorders"):
            tcPr.remove(child)
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        spec = {"top": top, "bottom": bottom, "left": left, "right": right}.get(side)
        if spec is not None:
            val, sz, color = spec
            tcBorders.append(parse_xml(
                f'<w:{side} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            ))
        else:
            tcBorders.append(parse_xml(f'<w:{side} {nsdecls("w")} w:val="nil"/>'))
    tcPr.append(tcBorders)


def add_three_line_table(
    doc,
    caption: str,
    headers: list,
    rows: list,
    col_widths_dxa: list,
    alignments: list = None,
    note: str = "",
    font_size_pt: float = 9.0,
    merge_cols: list = None,
    midrules_after: set = None,
    nowrap_cols: set = None,
    keep_together: bool = True,
    cell_pad_v_dxa: int = 18,
):
    """
    Create a classic academic three-line table (Top 1.5pt, Header bottom 0.75pt, Bottom 1.5pt, all black, 9pt font).
    Supports automatic vertical cell merging (`vMerge`) on specified `merge_cols`.
    """
    if caption:
        p_cap = doc.add_paragraph(style="Normal")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(caption) < 52 else WD_ALIGN_PARAGRAPH.JUSTIFY
        p_cap.paragraph_format.first_line_indent = Pt(0)
        p_cap.paragraph_format.space_before = Pt(6)
        p_cap.paragraph_format.space_after = Pt(3)
        p_cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_cap.paragraph_format.line_spacing = Pt(14.0)
        p_cap.paragraph_format.keep_with_next = True
        add_formatted_text(p_cap, caption, size_pt=9.5)

    all_rows = headers + rows
    header_rows = len(headers)
    n_rows = len(all_rows)
    n_cols = len(col_widths_dxa)
    if alignments is None:
        alignments = ["center"] * n_cols
    if merge_cols is None:
        merge_cols = []
    if midrules_after is None:
        midrules_after = set()
    if nowrap_cols is None:
        nowrap_cols = set(range(1, n_cols))

    tbl = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.style = "Normal Table"
    tblPr = tbl._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith(("tblW", "jc", "tblLayout", "tblCellMar", "tblBorders")):
            tblPr.remove(child)
    total_w = sum(col_widths_dxa)
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_w}" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>'
        f'<w:top w:w="{cell_pad_v_dxa}" w:type="dxa"/><w:left w:w="40" w:type="dxa"/>'
        f'<w:bottom w:w="{cell_pad_v_dxa}" w:type="dxa"/><w:right w:w="40" w:type="dxa"/>'
        '</w:tblCellMar>'
    ))

    for r_idx, row_data in enumerate(all_rows):
        row = tbl.rows[r_idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if r_idx < header_rows:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        # Top 1.5pt (sz=12), Header bottom 0.75pt (sz=6), Bottom 1.5pt (sz=12)
        top_b = ("single", "12", "000000") if r_idx == 0 else None
        bot_b = None
        if r_idx == header_rows - 1:
            bot_b = ("single", "6", "000000")
        elif r_idx == n_rows - 1:
            bot_b = ("single", "12", "000000")
        elif r_idx in midrules_after:
            bot_b = ("single", "4", "CBD5E1")

        for c_idx in range(n_cols):
            cell = row.cells[c_idx]
            tcPr = cell._tc.get_or_add_tcPr()
            for child in list(tcPr):
                if child.tag.endswith(("tcW", "vAlign", "vMerge", "noWrap")):
                    tcPr.remove(child)
            tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_widths_dxa[c_idx]}" w:type="dxa"/>'))
            tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
            if c_idx in nowrap_cols:
                tcPr.append(parse_xml(f'<w:noWrap {nsdecls("w")}/>'))

            # Check vertical merge for body rows
            is_merged_cont = False
            if r_idx >= header_rows and c_idx in merge_cols:
                cur_val = all_rows[r_idx][c_idx]
                prev_val = all_rows[r_idx - 1][c_idx] if r_idx > header_rows else None
                next_val = all_rows[r_idx + 1][c_idx] if r_idx + 1 < n_rows else None
                if cur_val == prev_val:
                    tcPr.append(parse_xml(f'<w:vMerge {nsdecls("w")}/>'))
                    is_merged_cont = True
                elif cur_val == next_val:
                    tcPr.append(parse_xml(f'<w:vMerge {nsdecls("w")} w:val="restart"/>'))

            set_cell_borders(cell, top=top_b, bottom=bot_b)

            p = cell.paragraphs[0]
            p.style = "Normal"
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            if keep_together and r_idx < n_rows - 1:
                p.paragraph_format.keep_with_next = True

            align = alignments[c_idx] if r_idx >= header_rows else "center"
            if align == "center":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align == "right":
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            if not is_merged_cont:
                val_str = str(row_data[c_idx])
                if r_idx < header_rows:
                    add_formatted_text(p, val_str, size_pt=font_size_pt, base_bold=True)
                else:
                    add_formatted_text(p, val_str, size_pt=font_size_pt, base_bold=False)

    if note:
        p_note = doc.add_paragraph(style="Normal")
        p_note.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_note.paragraph_format.first_line_indent = Pt(0)
        p_note.paragraph_format.space_before = Pt(2)
        p_note.paragraph_format.space_after = Pt(5)
        p_note.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        add_formatted_text(p_note, note, size_pt=8.5)

    return tbl


def add_code_listing(doc, title: str, code_text: str):
    """LaTeX listings-style academic code box: #F8FAFC background, top/bottom 0.75pt black rules, Consolas 8.2pt, 10pt spacing."""
    p_cap = doc.add_paragraph(style="Normal")
    p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_cap.paragraph_format.first_line_indent = Pt(0)
    p_cap.paragraph_format.space_before = Pt(3)
    p_cap.paragraph_format.space_after = Pt(1.5)
    p_cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p_cap.paragraph_format.keep_with_next = True
    add_formatted_text(p_cap, f"**{title}**", size_pt=9.2)

    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Normal Table"
    tblPr = tbl._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith(("tblW", "jc", "tblLayout", "tblCellMar")):
            tblPr.remove(child)
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{BODY_WIDTH_DXA}" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>'
        '<w:top w:w="32" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
        '<w:bottom w:w="32" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>'
        '</w:tblCellMar>'
    ))

    tbl.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    cell = tbl.rows[0].cells[0]
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{BODY_WIDTH_DXA}" w:type="dxa"/>'))
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F8FAFC"/>'))
    rule_spec = ("single", "6", "1E293B")  # 0.75pt top & bottom rules (LaTeX listings style)
    set_cell_borders(cell, top=rule_spec, bottom=rule_spec, left=None, right=None)

    lines = code_text.strip("\n").splitlines()
    for idx, line in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        p.style = "Normal"
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(10.0)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if "#" in line:
            code_part, comment_part = line.split("#", 1)
            if code_part:
                r1 = p.add_run(code_part)
                set_run_fonts(r1, cn_font="宋体", en_font="Consolas", size_pt=8.2)
                r1.font.color.rgb = RGBColor(15, 23, 42)
            r2 = p.add_run("#" + comment_part)
            set_run_fonts(r2, cn_font="宋体", en_font="Consolas", size_pt=8.2)
            r2.font.color.rgb = RGBColor(71, 85, 105)
        else:
            r = p.add_run(line)
            set_run_fonts(r, cn_font="宋体", en_font="Consolas", size_pt=8.2)
            r.font.color.rgb = RGBColor(15, 23, 42)


# ==============================================================================
# 4. Main Document Construction
# ==============================================================================
def configure_a4_section(sec):
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)


def set_section_valign(sec, val: str):
    """Explicitly set vertical alignment ('center' or 'top') on a section's sectPr."""
    sectPr = sec._sectPr
    for child in list(sectPr):
        if child.tag.endswith("vAlign"):
            sectPr.remove(child)
    v_el = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{val}"/>')
    cols_el = sectPr.find(qn("w:cols"))
    if cols_el is not None:
        cols_el.addnext(v_el)
    else:
        sectPr.append(v_el)


def configure_course_header_footer(sec):
    """Set course header and page-number footer on main body section."""
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p_hdr = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
    p_hdr.style = "Header"
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_hdr.text = ""
    r_hdr = p_hdr.add_run("《图像处理与计算机视觉》开放性大作业一  ·  研究性题目：图像去雾")
    set_run_fonts(r_hdr, cn_font="宋体", en_font="Times New Roman", size_pt=9.0)

    ftr = sec.footer
    ftr.is_linked_to_previous = False
    p_ftr = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
    p_ftr.style = "Footer"
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ftr.text = ""
    r_pre = p_ftr.add_run("《图像处理与计算机视觉》  ·  第 ")
    set_run_fonts(r_pre, cn_font="宋体", en_font="Times New Roman", size_pt=9.0)
    for fld in [
        parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'),
        parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>'),
        parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'),
        parse_xml(f'<w:t {nsdecls("w")}>2</w:t>'),
        parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'),
    ]:
        r_pg = p_ftr.add_run()
        set_run_fonts(r_pg, cn_font="宋体", en_font="Times New Roman", size_pt=9.0)
        r_pg._r.append(fld)
    r_post = p_ftr.add_run(" 页")
    set_run_fonts(r_post, cn_font="宋体", en_font="Times New Roman", size_pt=9.0)


def build_document():
    doc = docx.Document(str(TEMPLATE_PATH))

    # Clear all template body elements except sectPr
    body = doc._element.body
    for child in list(body):
        if not child.tag.endswith("sectPr"):
            body.remove(child)

    # Remove automatic numbering from Heading 1 / Heading 2 styles so explicit numbering is clean
    for st_name in ["Heading 1", "Heading 2", "Heading 3"]:
        if st_name in doc.styles:
            pPr = doc.styles[st_name]._element.find(qn("w:pPr"))
            if pPr is not None:
                numPr = pPr.find(qn("w:numPr"))
                if numPr is not None:
                    pPr.remove(numPr)

    # Configure Section 0 (Page 1 Standalone Cover Page ONLY: A4, vertically centered, no header/footer)
    sec0 = doc.sections[0]
    configure_a4_section(sec0)
    sec0.different_first_page_header_footer = True
    for p in sec0.first_page_header.paragraphs:
        p.text = ""
        p.style = "Normal"
    for p in sec0.first_page_footer.paragraphs:
        p.text = ""
        p.style = "Normal"
    set_section_valign(sec0, "center")

    # --------------------------------------------------------------------------
    # PAGE 1: BALANCED, VERTICALLY CENTERED COURSE COVER PAGE
    # --------------------------------------------------------------------------
    p_banner = doc.add_paragraph(style="Normal")
    p_banner.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_banner.paragraph_format.first_line_indent = Pt(0)
    p_banner.paragraph_format.space_before = Pt(0)
    p_banner.paragraph_format.space_after = Pt(18)
    r_ban = p_banner.add_run("《图像处理与计算机视觉》开放性大作业一  ·  研究性题目（一）：图像去雾")
    set_run_fonts(r_ban, cn_font="宋体", en_font="Times New Roman", size_pt=13.5, bold=True)

    p_cn_title = doc.add_paragraph(style="Normal")
    p_cn_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cn_title.paragraph_format.first_line_indent = Pt(0)
    p_cn_title.paragraph_format.space_before = Pt(4)
    p_cn_title.paragraph_format.space_after = Pt(10)
    p_cn_title.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_cn_title.paragraph_format.line_spacing = Pt(30.0)
    r_cnt = p_cn_title.add_run("PRISM：感知可行域引导的低秩层分配\n——面向真实图像去雾的评训分离适配")
    set_run_fonts(r_cnt, cn_font="宋体", en_font="Times New Roman", size_pt=21.5, bold=True)

    p_en_title = doc.add_paragraph(style="Normal")
    p_en_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_title.paragraph_format.first_line_indent = Pt(0)
    p_en_title.paragraph_format.space_before = Pt(0)
    p_en_title.paragraph_format.space_after = Pt(24)
    p_en_title.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_en_title.paragraph_format.line_spacing = Pt(19.0)
    r_ent = p_en_title.add_run(
        "PRISM: Perceptually Regulated Iso-budget Snapshot Masking\nfor Unsupervised Real-World Image Dehazing"
    )
    set_run_fonts(r_ent, cn_font="宋体", en_font="Times New Roman", size_pt=13.0, bold=False, italic=True)

    # Cover Teaser Image (Minimalist black-and-white header, no colored boxes)
    p_teaser = doc.add_paragraph(style="Normal")
    p_teaser.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_teaser.paragraph_format.first_line_indent = Pt(0)
    p_teaser.paragraph_format.space_before = Pt(4)
    p_teaser.paragraph_format.space_after = Pt(4)
    p_teaser.add_run().add_picture(str(prepare_word_image(ASSETS_DIR / "cover_teaser.png")), width=Cm(16.5))

    p_teaser_cap = doc.add_paragraph(style="Normal")
    p_teaser_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_teaser_cap.paragraph_format.first_line_indent = Pt(0)
    p_teaser_cap.paragraph_format.space_before = Pt(2)
    p_teaser_cap.paragraph_format.space_after = Pt(26)
    add_formatted_text(
        p_teaser_cap,
        "代表性真实浓雾场景（tiananmen 与 neighbour）去雾视觉效果及 FADE↓ / MUSIQ↑ 指标对比（粗体为最优）",
        size_pt=9.0,
    )

    # Minimalist 2-Member Team Table (no workload column per user instruction)
    team_headers = [["序号", "成员姓名", "学号", "专业班级"]]
    team_rows = [
        ["1", "[成员一姓名]", "[202X0001]", "[计算机科学与技术]"],
        ["2", "[成员二姓名]", "[202X0002]", "[人工智能实验班]"],
    ]
    add_three_line_table(
        doc,
        caption="**课题小组成员**",
        headers=team_headers,
        rows=team_rows,
        col_widths_dxa=[1200, 2400, 2610, 3200],
        alignments=["center", "center", "center", "center"],
        font_size_pt=9.5,
        cell_pad_v_dxa=42,
    )

    p_cover_meta = doc.add_paragraph(style="Normal")
    p_cover_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cover_meta.paragraph_format.first_line_indent = Pt(0)
    p_cover_meta.paragraph_format.space_before = Pt(20)
    p_cover_meta.paragraph_format.space_after = Pt(0)
    p_cover_meta.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_cover_meta.paragraph_format.line_spacing = Pt(18.0)
    add_formatted_text(
        p_cover_meta,
        "**指导教师：**魏颖 教授    |    **提交日期：**2026 年 11 月\n"
        "**开源代码与权重仓库：**https://github.com/NEUMelon/PRISM.git\n"
        "**英文学术论文（Paper PDF）：**https://github.com/NEUMelon/PRISM/blob/main/paper/PRISM_Submission_EN.pdf",
        size_pt=10.0,
    )

    # Next-Page Section Break isolating Page 1 from Page 2 (Standalone TOC Page)
    ensure_toc_styles(doc)
    sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_a4_section(sec1)
    set_section_valign(sec1, "top")  # CRITICAL: Body pages 2+ MUST be top-aligned, NOT vertically centered!
    sec1.different_first_page_header_footer = False
    configure_course_header_footer(sec1)

    # --------------------------------------------------------------------------
    # PAGE 2: STANDALONE TABLE OF CONTENTS (目 录)
    # --------------------------------------------------------------------------
    p_toc_title = doc.add_paragraph(style="Normal")
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_title.paragraph_format.first_line_indent = Pt(0)
    p_toc_title.paragraph_format.space_before = Pt(4)
    p_toc_title.paragraph_format.space_after = Pt(10)
    r_toct = p_toc_title.add_run("目  录")
    set_run_fonts(r_toct, cn_font="宋体", en_font="Times New Roman", size_pt=14.0, bold=True)

    add_toc_field(doc)
    doc.add_page_break()

    # --------------------------------------------------------------------------
    # PAGE 3 ONWARD: "摘 要" AND "1 引言" (NO DUPLICATE TITLE)
    # --------------------------------------------------------------------------
    p_abs_title = doc.add_paragraph(style="Normal")
    p_abs_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abs_title.paragraph_format.first_line_indent = Pt(0)
    p_abs_title.paragraph_format.space_before = Pt(2)
    p_abs_title.paragraph_format.space_after = Pt(5)
    set_outline_level(p_abs_title, 0)
    r_abst = p_abs_title.add_run("摘  要")
    set_run_fonts(r_abst, cn_font="宋体", en_font="Times New Roman", size_pt=14.0, bold=True)

    add_body_p(
        doc,
        "无配对真实图像去雾中，合成预训练骨干常借低秩适配（LoRA）与由雾到清（H2C）对比语义损失迁移真实域特征。"
        "然而，直接将该损失同时用于低秩权重更新与适配层位置搜索，易引发语义—感知失配：连续双层门控受单样本外积超梯度近似噪声干扰，"
        "跨随机种子选层分歧显著；而无约束最小化训练损失又会诱导浅层瓶颈过度增强，放大暗部色偏与边缘光晕并拉低整图感知质量。"
        "为此，本文提出感知可行域引导的等容量快照筛选框架 PRISM，将结构位置决策与梯度反传解耦。"
        "方法在全候选瓶颈块引入卷积激活解释方差对齐（Conv-EVA）构建早期校准共享状态，随后通过无梯度前向推理遍历全部等容量候选子集，"
        "以预置组合的整图多尺度质量评分为安全下界构建感知可行域，并在可行域内锁定训练损失最低的帕累托最优适配子集。"
        "该策略无需改动下游训练目标，即可将跨种子性能标准差缩小近一个数量级；微调仅占骨干 3.03% 的参数并在推理期无损重参数化合并，"
        "零额外计算开销下将 RTTS 全集（4,322 幅）FADE 较 BiLaLoRA 降低 17.0%（0.752 降至 0.6241），同时稳步提升多项真实基准与课程实测图像的感知质量。"
        "项目源码、合并权重及完整英文论文已开源至 GitHub 仓库 https://github.com/NEUMelon/PRISM.git"
        " （英文论文 PDF：[PRISM_Submission_EN.pdf](https://github.com/NEUMelon/PRISM/blob/main/paper/PRISM_Submission_EN.pdf)）。",
    )

    p_kw = doc.add_paragraph(style="Normal")
    p_kw.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_kw.paragraph_format.first_line_indent = Pt(21.0)
    p_kw.paragraph_format.space_before = Pt(2)
    p_kw.paragraph_format.space_after = Pt(6)
    p_kw.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_kw.paragraph_format.line_spacing = Pt(18.0)
    add_formatted_text(
        p_kw,
        "**关键词：**真实图像去雾；参数高效微调；低秩层位置分配；激活解释方差对齐；评训分离筛选",
        size_pt=10.5,
    )

    # --------------------------------------------------------------------------
    # SECTION 1: 1 引言
    # --------------------------------------------------------------------------
    add_h1(doc, "1  引言")
    add_body_p(
        doc,
        "真实场景单幅图像去雾（Single Image Dehazing）是计算机视觉与计算摄影领域的经典病态逆问题，直接关系到自动驾驶环境感知、全天候安防监控与航拍遥感解译等下游系统的鲁棒性。在经典大气散射物理模型（He et al., 2009; Fattal, 2014）下，传感器捕获的有雾观测图像 I(x) 可表述为场景真实辐射率 J(x) 与全球大气光 A 沿视线深度的凸组合：",
    )
    add_omml_equation(doc, build_eq1_scattering(), 1)
    add_body_p(
        doc,
        "其中 x 为二维像素空间坐标，t(x) 为介质透射率，β 为大气散射系数，d(x) 为场景深度。由于真实户外场景难以同步采集严格空间对齐的“有雾—无雾”图像对（Li et al., 2019; Ancuti et al., 2019），在合成数据集上预训练的复原骨干网络（如 DEA-Net，Chen et al., 2024）直接泛化至真实雾图时面临显著的域偏移（Domain Shift）：模型往往在浓雾远景处残留灰白雾幕，或在复杂光照与暗部区域产生色彩失真。",
    )
    add_body_p(
        doc,
        "为在不重训全量骨干参数的前提下弥合合成域与真实域的差距，近期工作 BiLaLoRA（Zhang et al., 2026）将参数高效微调（LoRA，Hu et al., 2022）引入预训练 DEA-Net 的第 3 级特征转换瓶颈（8 个候选 CGA 混合块 D_{1}–D_{8}），并利用预训练 CLIP（Radford et al., 2021）的由雾到清（Haze-to-Clear, H2C）余弦方向损失提供无配对语义监督，通过双层优化从 8 个候选块中搜索 K=3 个适配块（仅占骨干总参数量的 3.03%）。尽管这一参数高效适配范式极具启发性，但我们在严格复现与 56 种全组合遍历诊断中发现，现有双层门控选层范式存在两个制约性能与稳定性的核心瓶颈：",
    )
    add_body_p(
        doc,
        "H2C 损失在经双三次缩放至 224×224 的局部裁剪块上最大化图像特征向“无雾文本提示”方向的余弦位移。然而，语义嵌入空间的余弦对齐并不等价于整图尺度的感知自然度（Ke et al., 2021; Blau & Michaeli, 2018）。在第 5 轮共享快照上对全部 C(8,3)=56 种三层组合的穷举评估表明，纯粹使 H2C 训练损失最小化的子集 {5,6,8} 因引入浅层瓶颈块 D_{5}，会激进地拉伸局部高频对比度并剥离低频雾感，虽然在 FADE 与 NIQE 等局部对比度指标上取得低值，却在整图尺度诱发严重的暗部色偏与边缘伪影，导致整图多尺度图像质量评分（MUSIQ）较参考组合 {1,7,8} 大幅下跌（RTTS-128 上由 62.26 跌至 61.18，Dense-Haze 上由 54.80 跌至 52.65）。",
        run_in_title="（1）无配对去雾中的“语义—感知失配”（Semantic-Perceptual Mismatch）。",
    )
    add_body_p(
        doc,
        "BiLaLoRA 的外层门控更新依赖单步随机验证梯度与单步训练梯度的秩一外积近似（Rank-One Outer-Product Approximation）来估计逆海森矩阵与超梯度。在无配对真实雾图存在显著跨图分布差异与小批量采样噪声的条件下，该近似超梯度的方差较大，导致连续门控排序在不同随机种子下频繁发生轨迹分岔（三个独立种子 3407/3408/3409 分别漂移至 {2,6,8}、{1,6,7} 与 {1,4,8}），使下游微调在 RTTS-128 上的 FADE 标准差高达 0.0474、MUSIQ 标准差达 0.54。同时，标准 LoRA 的随机高斯初始化在仅有 5 轮的短暂热身期内尚未与真实雾图的主导激活方向对齐，进一步加剧了早期快照评估的随机波动。",
        run_in_title="（2）双层连续门控优化的高方差不稳定性。",
    )
    add_body_p(
        doc,
        "针对上述挑战，本文提出 **PRISM**（感知可行域引导的低秩层分配）。PRISM 遵循**“评训分离（Evaluation-Training Decoupling）”**原则，如图 1 所示：（1）首先在全部 8 个候选瓶颈块上利用真实雾图的激活协方差矩阵执行卷积激活解释方差对齐（Conv-EVA），在不改变初始输出映射（ΔW_{0}=0）的前提下最大化初始梯度范数并构建稳定的第 5 轮共享快照；（2）随后在无梯度前向推理模式下遍历全部 C(8,3)=56 种等容量子集，以预置参考组合 {1,7,8} 的整图 MUSIQ 得分作为感知安全下界 τ，拦截纯 H2C 最小化诱发的 {5,6,8} 天空光晕与立面过增强伪影（图 1(b)），并在可行域内一致锁定 {6,7,8}，将跨种子标准差降低近一个数量级；（3）仅微调 3.03% 参数并在推理期通过线性重参数化无损并入主干（零额外参数与 FLOPs），在 RTTS 全集（4,322 图）上将 FADE 较 BiLaLoRA 显著降低 17.0%（0.752→0.6241），并在课程指定 Course-10 基准上实现 10/10 逐图 FADE 最优。",
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig1_motivation.png",
        width_cm=16.5,
        caption_text=(
            "**图 1：研究动机与实证困境图。** "
            "(a) 基于种子 3407 在第 5 轮共享快照 θ_{5} 上全部 56 种三层组合的真实二维得分分布：纯最小化 H2C 选出含浅层块 D_{5} 的 {5,6,8}（珊瑚红三角，h=0.95753 最低），但其整图 MUSIQ 跌入感知安全护栏 τ = p(S_{ref}) 下方的退化区；连续双层门控在三组随机种子间分别漂移至 {2,6,8}、{1,6,7} 与 {1,4,8}（柔雾紫轨迹）；而 PRISM 在可行域内锁定深层连续子集 {6,7,8}（青绿星标）。"
            "(b) 真实城市建筑场景 BJ_Bing_116（RTTS-128）上的双赢对比及 2.5× 局部放大：与纯 H2C 最优组合 {5,6,8}（FADE 0.540 / MUSIQ 70.09）在高楼立面和天空边界引发的暗沉光晕相比，PRISM {6,7,8} 同步实现更优残雾清除（FADE 0.471, ↓12.8%）与自然通透感（MUSIQ 72.82, +2.73）。"
        ),
    )

    # --------------------------------------------------------------------------
    # SECTION 2: 2 相关工作
    # --------------------------------------------------------------------------
    add_h1(doc, "2  相关工作")
    add_h2(doc, "2.1  无配对真实图像去雾监督信号")
    add_body_p(
        doc,
        "早期无标签去雾方法主要依赖大气散射物理先验，如暗通道先验（DCP，He et al., 2009）与颜色线先验（Color-Lines，Fattal, 2014），但在天空、白色建筑或复杂非均匀照明区域容易违反先验假设而产生光晕与色偏。深度无配对去雾方法通过合成与真实域双向转换（DAD，Shao et al., 2020）、物理引导后微调（PSD，Chen et al., 2021）、浓度与深度自增强分解（D4，Yang et al., 2022）、高质量码本先验（RIDCP，Wu et al., 2023）、定位-移除-调整三阶段流水线（KANet，Feng et al., 2024）、正交解耦对比正则（ODCR，Wang et al., 2024b）、一致性伪标签协同展开（CoRUN，Fang et al., 2024）、压缩适配迁移（CoA，Ma et al., 2025）、迭代预测-评判码解码（IPC，Fu et al., 2025）以及物理雾迁移网络（PHATNet，Tsai et al., 2025）不断提升真实域复原能力。BiLaLoRA（Zhang et al., 2026）引入 CLIP（Radford et al., 2021）的 H2C 方向损失作为轻量级无配对监督。本文严格沿用相同的 H2C 训练损失，证明无需修改训练目标，仅修正选层决策准则即可充分释放预训练骨干的去雾潜力。",
    )
    add_h2(doc, "2.2  参数高效微调与数据驱动子空间初始化")
    add_body_p(
        doc,
        "现代图像复原骨干（如 MSBDN、DeHamer、C2PNet、DEA-Net，以及通用复原模型 PromptIR、DiffUIR、MoCE-IR、FoundIR）通常包含数百万至数千万参数，在无配对小样本真实雾图上全量微调极易过拟合。LoRA（Hu et al., 2022）通过低秩分解冻结原权重并仅更新旁路矩阵。为加速低秩适配收敛，近期研究探索了数据或权重驱动的初始化策略：PiSSA（Meng et al., 2024）分解原权重主奇异方向，LoRA-GA（Wang et al., 2024a）对齐初始梯度子空间，CorDA（Yang et al., 2024）引入上下文协方差分解，而 EVA（Paischer et al., 2025）利用激活奇异向量捕获输入解释方差最大的特征方向。本文将 EVA 推广至二维卷积滑动窗口激活（Conv-EVA），在严格保持零初始增量（B_{0}=0）的同时稳定早期共享快照，为无梯度选层提供可靠的评估基底。",
    )
    add_h2(doc, "2.3  感知质量引导的结构决策与评训分离机制")
    add_body_p(
        doc,
        "无参考图像质量评价指标从不同维度刻画去雾效果：FADE（Choi et al., 2015）度量雾密度感知特征，BIQME（Gu et al., 2018）评估增强图像的对比度、锐度与自然度，NIQE（Mittal et al., 2013）衡量局部自然场景统计偏离，而 MUSIQ（Ke et al., 2021）基于多尺度原分辨率分块 Transformer 捕获整图结构与色彩协调性（在 KonIQ-10k 上预训练，Hosu et al., 2020）。感知—失真权衡理论（Blau & Michaeli, 2018）指出单一损失往往难以兼顾所有感知维度。借鉴 RankSRGAN（Zhang et al., 2019）与平台感知神经架构搜索 MnasNet（Tan et al., 2019）中的约束优化思想，PRISM 将整图 MUSIQ 作为前向选层的可行域护栏而非梯度反传损失，有效规避了梯度冲突与尺度不一致问题。",
    )

    # --------------------------------------------------------------------------
    # SECTION 3: 3 PRISM 算法原理与实现步骤
    # --------------------------------------------------------------------------
    add_h1(doc, "3  PRISM 算法原理与实现步骤")
    add_h2(doc, "3.1  卷积低秩表示与 Conv-EVA 子空间对齐")
    add_body_p(
        doc,
        "预训练 DEA-Net（Chen et al., 2024）在最低空间分辨率（H/4 × W/4）的第 3 级特征转换瓶颈中包含 M=8 个串联的内容引导注意力（CGA）混合块，记为候选层集合 {D_{1}, D_{2}, …, D_{8}}。遵循 BiLaLoRA（Zhang et al., 2026）的等容量协议，每个候选块在两条核心 3×3 卷积路径（`conv1.conv` 与 `conv2`）上挂载秩 r=8、缩放因子 α_{LoRA}=16 的低秩旁路。设某候选卷积核原始冻结权重为 W（通道数 C_{out}=C_{in}=64，卷积核尺寸 K_{h}=K_{w}=3），定义展平算子 ℛ 将四维张量映射为二维矩阵（输入展开维度 d_{in}=C_{in}K_{h}K_{w}=576），则挂载低秩旁路后的等效卷积核权重 W′ 可表示为：",
    )
    add_omml_equation(doc, build_eq2_conv_lora(), 2)
    add_body_p(
        doc,
        "其中 B（形状 64×8）与 A（形状 8×576）为低秩适配矩阵，γ = α_{LoRA}/r = 2 为固定缩放常数，g = σ(u_{gate}) ∈ (0,1) 为由标量参数 u_{gate} 经 Sigmoid 映射得到的层内门控系数，ℛ^{-1} 为将二维矩阵还原为四维卷积核张量的逆重排算子。",
    )
    add_body_p(
        doc,
        "标准 LoRA 采用随机高斯初始化 A_{0} 与零初始化 B_{0}=0。然而在仅有 5 轮的早期快照热身期内，随机行空间与真实雾图卷积特征的主导变化方向缺乏对齐，导致各候选块的学习进度受初始随机种子影响较大。为此，本文提出卷积激活解释方差对齐（Conv-EVA）：在训练开始前，从权重训练集固定抽取 32 幅真实雾图（随机种子 3407），使冻结主干执行一次前向传播，在每个候选卷积层通过 `im2col` 展开并均匀采样 N 个滑动窗口特征向量（每层上限 N=8,192），拼成激活矩阵 X（形状 N×576）。首先对 X 执行行均值中心化得到 X_{c} 并计算经验协方差矩阵 Σ 的特征分解：",
    )
    add_omml_equation(doc, build_eq3_covariance(), 3)
    add_body_p(
        doc,
        "其中 1 为长度为 N 的全 1 列向量，V 为按特征值降序排列的正交特征向量矩阵。取前 d=r=8 个主导右奇异向量构成子空间基底 V_{d}^{⊤}，为保持与原随机初始化矩阵 A_{rand} 相同的行范数尺度并严格维持零初始增量，Conv-EVA 按下式构造初始矩阵 A_{0} 与 B_{0}：",
    )
    add_omml_equation(doc, build_eq4_conv_eva(), 4)
    add_body_p(
        doc,
        "其中 ε=10^{-8} 防止除零。由式 (4) 可知，由于 B_{0}=0，第 0 轮初始权重增量恒满足 ΔW_{0}=0，从而严格保留预训练 DEA-Net 的初始复原映射；同时将 A_{0} 的行空间对齐至 Σ 的前 r 主特征方向能够最大化早期梯度的期望 Frobenius 范数，使全部 8 个候选块在第 5 轮共享快照 θ_{5} 处均获得充分、均衡的特征表达。PRISM 总体架构如图 2 所示。",
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig_prism_arch.png",
        width_cm=16.5,
        caption_text=(
            "**图 2：PRISM 总体架构图（三栏对称全矢量设计）。** "
            "（左翼）在冻结 DEA-Net 第 3 级瓶颈块 D_{1}–D_{8} 的核心 3×3 卷积旁路中，基于 32 幅训练雾图中心化激活协方差 Σ = V Λ V^{⊤} 提取前 d=8 主奇异方向初始化 A_{0} = D V_{d}^{⊤}，配合 B_{0}=0 实现零初始扰动（ΔW_{0}=0）；"
            "（中栏）仅对选中的深层子集 {6,7,8} 执行 25 轮精调（微调参数占 3.03%），推理期通过线性重参数化 W′ = W + ΔW 无损并入原卷积核，实现零额外推理参数与 FLOPs；"
            "（右翼）从全 8 块共同热身第 5 轮快照 θ_{5} 出发，无梯度穷举 56 种等容量子集并依据感知护栏 τ = p(S_{ref}) 拦截 {5,6,8}、锁定帕累托最优子集 {6,7,8}。"
        ),
    )

    add_h2(doc, "3.2  语义—感知失配机理与瓶颈层级边际敏感度归因")
    add_body_p(
        doc,
        "在无配对训练阶段，给定输入有雾图像 I 及其经网络复原的输出图像 J = f_{θ}(I)，H2C 语义损失利用冻结的预训练 CLIP 视觉编码器 E_{img} 与文本编码器 E_{txt}，度量图像在特征空间中由“有雾”向“清晰”的位移向量 Δe_{img}(J, I) = E_{img}(J) − E_{img}(I) 与文本锚点差向量 Δe_{txt} 之间的余弦距离：",
    )
    add_omml_equation(doc, build_eq5_h2c_loss(), 5)
    add_body_p(
        doc,
        "为系统诊断“最小化 H2C 是否等价于提升整图感知质量”，我们将全部 M=8 个候选瓶颈块同时开启 Conv-EVA 低秩适配并共同训练至第 5 轮，保存共享快照 θ_{5}。给定目标块数预算 K=3，等容量候选子集空间 Ω_{K} 共包含 C(8,3)=56 个候选子集，每个子集严格对应 110,595 个可训练参数。对于任一子集 S ∈ Ω_{K}，我们仅保留 S 内模块的 LoRA 旁路并关闭其余 5 个模块的旁路，在从架构学习集中独立划出且与权重训练半集像素哈希零重叠的 32 幅选层雾图集 𝒟_{sel} 上执行无梯度前向推理，分别计算局部 H2C 语义损失均值 h(S) 与严格遵循部署流水线 Π_{deploy}（含 16 倍数双三次缩放、反归一化及 uint8 量化）的整图 MUSIQ 感知评分均值 p(S)：",
    )
    add_omml_equation(doc, build_eq6_snapshot_eval(), 6)
    add_body_p(
        doc,
        "为进一步揭示“语义—感知失配”的层级根源，我们基于全部 3 个随机种子（3407、3408、3409）× 56 组合（共 168 组探测样本）计算每个候选瓶颈块 D_{1}–D_{8} 入选子集时相对于未入选时的平均边际感知增益 ΔMUSIQ 与平均边际语义去雾增益 −Δh，所得二维热力矩阵如图 3 所示。可以清晰地看到一道物理分工分界线：浅层与中浅层块 D_{1}–D_{5} 入选均带来负的整图感知收益（ΔMUSIQ ∈ [-0.251, -0.013]），其中尤以中层块 D_{5} 呈现最强烈的反向拉扯效应——其提供全场最高的语义去雾边际增益（−Δh = +2.06×10^{-3}），却同时造成断崖式的感知质量跌落（ΔMUSIQ = −0.251），这正是纯 H2C 最小化准则极易被“诱捕”进 D_{5} 并选出 {5,6,8} 的内在数学机理；相反，深层瓶颈块 D_{6}、D_{7}、D_{8} 均贡献正向感知增益（分别高达 +0.109、+0.149、+0.130）且同时保持稳健的语义去雾收益，构成了帕累托安全域。",
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig_layer_sensitivity_heatmap.png",
        width_cm=16.0,
        caption_text=(
            "**图 3：第 3 级瓶颈候选块 D_{1}–D_{8} 的跨三随机种子（3407/3408/3409）平均边际敏感度热力矩阵。** "
            "每个单元格统计包含第 b 块的子集与不包含该块的子集在选层集上的平均指标差。中层块 D_{5} 虽可大幅压低 H2C 损失（−Δh = +2.06×10^{-3}），却会诱发剧烈的整图感知退化（ΔMUSIQ = −0.251）；而深层连续块 {D_{6}, D_{7}, D_{8}} 兼顾正感知增益与去雾效率。"
        ),
    )

    add_h2(doc, "3.3  评训分离的感知安全可行域筛选与无损合并")
    add_body_p(
        doc,
        "图 4 进一步展示了三个独立随机种子（3407、3408、3409）下全部 56 个候选子集的 (h(S), p(S)) 完整散点分布景观。无论随机初始化如何变化，纯 H2C 最优组合 S_{H2C}={5,6,8}（红色三角，种子 3407 下 h=0.95753）均落入参考组合 S_{ref}={1,7,8}（蓝色方块，τ=70.015）下方的“语义过拟合退化区”。基于这一稳定规律，PRISM 构建带感知安全护栏的约束离散优化问题：在满足 p(S) ≥ τ = p(S_{ref}) 的可行域内选取 H2C 语义损失最小的组合 S^{⋆}：",
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig2_landscape.png",
        width_cm=15.4,
        caption_text=(
            "**图 4：三个独立随机种子（3407、3408、3409）在第 5 轮共享快照 θ_{5} 上的 56 种三层组合二维评估景观。** "
            "横轴为选层集 32 幅雾图的平均 H2C 损失 h(S)（越低代表 CLIP 语义去雾方向越强），纵轴为遵循真实部署链路的整图 MUSIQ 评分 p(S)。"
            "水平虚线标示参考组合 {1,7,8} 的感知安全下界 τ = p(S_{ref})；纯 H2C 最小化选出的 {5,6,8}（红三角）全部坠入退化区，而 PRISM 在青绿色感知安全可行域内三种子一致锁定 {6,7,8}（绿星）。"
        ),
    )
    add_omml_equation(doc, build_eq7_prism_rule(), 7)
    add_body_p(
        doc,
        "如图 4 所示，在种子 3407、3408、3409 下，满足 p(S) ≥ τ 的可行子集分别仅有 4 个、6 个和 4 个，且式 (7) 在全部三个随机种子下均**一致锁定深层连续瓶颈组合 S^{⋆}={6,7,8}**（种子 3407 下 h=0.95961, p=70.136）。选定 S^{⋆} 后，PRISM 直接从第 5 轮共享快照 θ_{5} 中无损提取 {6,7,8} 的已训低秩参数，剪除其余 5 个候选块的旁路，并恢复完全一致的随机数生成器（RNG）状态，在后续 25 轮（至第 30 轮结束）中仅使用原版 H2C 损失继续精调。当第 30 轮微调结束后，由于二维卷积运算满足线性分配律，每个选中块内的低秩旁路无需在推理期单独执行分支计算，而是通过下式精确合并回原 `nn.Conv2d` 卷积核：",
    )
    add_omml_equation(doc, build_eq8_reparam_merge(), 8)
    add_body_p(
        doc,
        "在参数预算上，3 个选中块共包含 6 个适配卷积层（计 30,729 个核心低秩参数，连同辅助缓冲共 110,595 可训参数，占 DEA-Net 主干 3,653,377 参数的 3.03%）。经式 (8) 合并后，推理模型参数量严格回落至 3,653,377（额外参数=0，额外 FLOPs=0），合并前后前向浮点张量最大绝对误差不超过 1.43×10^{-6}，保存为 8 位 PNG 图像后的像素级最大误差为 0 LSB。完整步骤归纳于算法 1。",
    )

    alg_headers = [["步骤", "算法 1：PRISM 评训分离的感知可行域低秩层筛选与无损重参数化合并流程"]]
    alg_rows = [
        ["输入", "预训练冻结骨干 DEA-Net 权重 W；候选瓶颈块集合 {D_{1},…,D_{8}}；预算 K=3；训练集 𝒟_{train}（250 幅）、选层集 𝒟_{sel}（32 幅）；参考组合 S_{ref}={1,7,8}。"],
        ["步骤 1", "**激活收集与中心化（Conv-EVA）**：在 32 幅校准雾图上前向运行冻结骨干，对每个候选卷积层采样滑动窗口激活 X，按式 (3) 计算中心化协方差 Σ = V Λ V^{⊤}。"],
        ["步骤 2", "**主导奇异子空间初始化**：取 Σ 的前 r=8 个主特征向量 V_{d}^{⊤}，按式 (4) 令 A_{0} = D V_{d}^{⊤}、B_{0} = 0，确保初始权重增量 ΔW_{0} = 0 且早期梯度范数最大化。"],
        ["步骤 3", "**早期全候选共同热身**：开启全部 8 个候选块的 LoRA 旁路，在 𝒟_{train} 上使用原版 H2C 损失（式 5）训练 5 轮，保存第 5 轮共享快照 θ_{5} 及 RNG 状态。"],
        ["步骤 4", "**等容量候选空间无梯度探测**：枚举全部 C(8,3)=56 个子集 S ∈ Ω_{3}，在无梯度模式下于 𝒟_{sel} 上按式 (6) 计算局部 H2C 均值 h(S) 与部署级整图 MUSIQ 均值 p(S)。"],
        ["步骤 5", "**设定感知安全护栏**：读取预置参考组合 S_{ref}={1,7,8} 的整图感知评分作为安全阈值 τ ← p(S_{ref})，构造感知安全可行域 ℱ = {S ∈ Ω_{3} : p(S) ≥ τ}。"],
        ["步骤 6", "**锁定帕累托最优子集**：在可行域 ℱ 内按式 (7) 选取使 H2C 损失最小的子集 S^{⋆} = arg min_{S ∈ ℱ} h(S)（三个种子一致收敛至 S^{⋆}={6,7,8}）。"],
        ["步骤 7", "**快照状态恢复与纯 H2C 精调**：从 θ_{5} 中提取 S^{⋆} 对应的 15 个可训张量（110,595 参数），卸载其余 5 块旁路，恢复规范 RNG 流，继续 25 轮纯 H2C 精调。"],
        ["步骤 8", "**推理期线性重参数化无损融合**：对 S^{⋆} 内全部 6 个适配卷积层，按式 (8) 计算 W′ ← W + γ g ℛ^{-1}(B A) 并写回原卷积核，卸载 LoRA 包装器后输出单文件权重 W′。"],
    ]
    add_three_line_table(
        doc,
        caption="",
        headers=alg_headers,
        rows=alg_rows,
        col_widths_dxa=[1050, 8360],
        alignments=["center", "left"],
        font_size_pt=8.3,
        midrules_after={1},
        nowrap_cols={0},
        cell_pad_v_dxa=10,
    )

    # --------------------------------------------------------------------------
    # SECTION 4: 4 核心算法编程实现与部署架构
    # --------------------------------------------------------------------------
    add_h1(doc, "4  核心算法编程实现与部署架构")
    add_body_p(
        doc,
        "本节从开源代码库提取与式 (2)–(8) 一一对应的核心算子实现，涵盖 Conv-EVA 谱初始化（`calibrate_eva.py`）、第 5 轮快照 56 组合无梯度筛选（`quality_select_probe.py`）与推理期线性重参数化无损融合（`merge_quality_deployment.py`），并给出零依赖命令行推理架构。",
    )
    code_listing_1 = """@torch.no_grad()
def compute_conv_eva_init(conv_module, calib_activations, rank=8, eps=1e-8):
    # 1. 使用 unfold 提取 3x3 滑动窗口激活矩阵 X in R^{N x 576} (对应 3.1 节激活收集)
    patches = [F.unfold(f, conv_module.kernel_size, padding=conv_module.padding,
                        stride=conv_module.stride).transpose(1, 2).reshape(-1, 576)
               for f in calib_activations]
    X = torch.cat(patches, dim=0).to(torch.float64)
    if X.size(0) > 8192:  # 固定种子 3407 均匀子采样上限 N = 8192
        X = X[torch.randperm(X.size(0), generator=torch.Generator().manual_seed(3407))[:8192]]
    # 2. 对应式 (3)-(4): 行均值中心化 -> SVD 前 d=8 右奇异基底 -> 行范数匹配且 B_0=0
    X_c = X - X.mean(dim=0, keepdim=True)
    _, _, Vh = torch.linalg.svd(X_c, full_matrices=False)
    V_d_T = Vh[:rank, :].to(torch.float32)
    A_rand = torch.randn_like(V_d_T) * 0.02
    D_scale = A_rand.norm(dim=1, keepdim=True) / (V_d_T.norm(dim=1, keepdim=True) + eps)
    return D_scale * V_d_T, torch.zeros(conv_module.out_channels, rank, dtype=torch.float32)"""
    add_code_listing(
        doc,
        "代码清单 1：Conv-EVA 卷积激活中心化与主导右奇异向量初始化算子（calibrate_eva.py，对应式 3–4）",
        code_listing_1,
    )

    code_listing_2 = """@torch.no_grad()
def restored_uint8(model, item):
    # 对应式 (6) 部署映射算子 Pi_deploy: 16 倍数双三次缩放 -> 前向去雾 -> 原尺寸还原 -> uint8
    h, w = item['height'], item['width']
    inp = Resize((h // 16 * 16, w // 16 * 16), InterpolationMode.BICUBIC, antialias=True)(item['tensor'].cuda())
    out = model(inp)[0] if isinstance(model(inp), tuple) else model(inp)
    out = Resize((h, w), InterpolationMode.BICUBIC, antialias=True)(out.squeeze(0))
    return out.mul(255).add_(0.5).clamp_(0, 255).permute(1, 2, 0).to('cpu', torch.uint8).numpy()

def select_quality_constrained(rows):
    # 对应式 (7): 在满足整图 MUSIQ p(S) >= p(S_ref) 的感知安全可行域内最小化 H2C 损失
    assert len(rows) == 56 and {tuple(r['selected']) for r in rows} == set(itertools.combinations(BLOCKS, 3))
    tau = next(r['musiq'] for r in rows if tuple(r['selected']) == PUBLISHED)  # 护栏 tau=p(S_ref)
    feasible = [r for r in rows if r['musiq'] >= tau]                           # 可行域 F
    chosen = min(feasible, key=lambda r: (r['h2c32_loss'], r['selected']))      # 锁定 {6,7,8}
    return {'quality_threshold': tau, 'feasible_triples': len(feasible), 'quality_constrained': chosen}"""
    add_code_listing(
        doc,
        "代码清单 2：第 5 轮快照 C(8,3)=56 子集无梯度前向探测与可行域筛选算子（quality_select_probe.py，对应式 6–7）",
        code_listing_2,
    )

    code_listing_3 = """@torch.inference_mode()
def merge_prism_lora_into_backbone(model, lora_cls):
    # 对应式 (8): W' = W + gamma * sigma(u_gate) * R^{-1}(B @ A)，将低秩旁路无损融合回原卷积核
    merged_count = 0
    for name, layer in list(model.named_modules()):
        if not isinstance(layer, lora_cls):
            continue
        conv = layer.original_layer  # 原始冻结的 nn.Conv2d(64, 64, kernel_size=3, padding=1)
        delta_w = torch.sigmoid(layer.alpha) * layer.scaling * (layer.lora_B @ layer.lora_A).view_as(conv.weight)
        conv.weight.copy_(conv.weight + delta_w)  # 线性重参数化原位写回: W' <- W + dW
        conv.padding_mode = 'zeros'
        parent, child = name.rsplit('.', 1)
        setattr(model.get_submodule(parent), child, conv)  # 卸载 LoRALayer，恢复纯 DEA-Net 算子图
        merged_count += 1
    assert merged_count == 6  # D_6, D_7, D_8 共 3 块 x 2 路径 = 6 个融合卷积核
    return model"""
    add_code_listing(
        doc,
        "代码清单 3：推理期低秩旁路线性重参数化无损融合算子（merge_quality_deployment.py，对应式 2、8）",
        code_listing_3,
    )
    add_body_p(
        doc,
        "在工程部署阶段，合并后的单文件权重与原 `DEA-Net` 保持逐张量同构（总参数量 3,653,377），推理脚本 `dehaze.py` 完全剥离 `CLIP` 与 `pyiqa` 依赖，仅依赖标准 PyTorch 与 Pillow 即可通过 `python dehaze.py --weights checkpoints/prism_seed3407.pth --device cuda` 对任意分辨率图像执行零额外开销去雾。",
    )

    # --------------------------------------------------------------------------
    # SECTION 5: 5 课程指定“雾天图像”全量实测分析（Course-10 Benchmark）
    # --------------------------------------------------------------------------
    add_h1(doc, "5  课程指定“雾天图像”全量实测分析（Course-10 Benchmark）")
    add_h2(doc, "5.1  实验环境与全部 10 幅图像定量对比")
    add_body_p(
        doc,
        "为全面检验算法在课程大作业真实场景下的去雾效能，我们对课程给定的全部 10 幅典型真实雾天图像（`bench`、`cones`、`dubai`、`forest`、`herzeliya`、`neighbour`、`pumpkins`、`road`、`tiananmen`、`train`，统称 **Course-10 基准**）执行了端到端去雾与四维盲图像质量评测（FADE↓、BIQME↑、Entropy↑、MUSIQ↑）。实验硬件为单张 NVIDIA RTX 4090 GPU（24 GB），软件环境为 PyTorch 2.5.1+cu124 与 `pyiqa 0.1.16`。所有方法均采用完全相同的 16 倍数双三次缩放与原尺寸还原推理协议。",
    )
    add_body_p(
        doc,
        "表 1 与表 2 分别给出了 Course-10 基准的 10 图总体均值对比及全部 10 幅图像×4 种方法的 40 行完整逐图评测记录。定量结果表明：（1）在衡量残雾消除彻底程度的 **FADE↓** 指标上，PRISM 将 10 图平均 FADE 从有雾原图的 1.0894、DEA-Net Base 的 0.4163 以及 BiLaLoRA 的 0.2418 进一步降低至 **0.2225**（较 DEA-Net Base 大幅降低 **46.6%**，较 BiLaLoRA 进一步降低 **8.0%**），且在全部 10 幅测试图像上实现了 **10/10 的逐图一致最优**；（2）在衡量整图多尺度感知质量的 **MUSIQ↑** 上，PRISM 取得全场最高均值 **68.26**（较有雾原图提升 +8.31 分，较 DEA-Net Base 提升 +1.46 分，较 BiLaLoRA 提升 +0.35 分），尤其在远景浓雾与大面积天空场景（如 `tiananmen` 76.75 vs. 75.02、`neighbour` 58.14 vs. 56.56、`train` 62.63 vs. 61.69、`pumpkins` 73.83 vs. 73.19）中优势显著；（3）在综合增强质量 **BIQME↑** 与信息熵 **Entropy↑** 上，PRISM 达到 0.6485 与 7.6419，与 BiLaLoRA 保持同一高水准且无过度锐化伪影。",
    )

    # Table 1: Course-10 Overall Mean Comparison
    t1_headers = [["对比方法", "参数微调比例", "额外推理开销", "FADE↓", "BIQME↑", "Entropy↑", "MUSIQ↑"]]
    t1_rows = [
        ["有雾原图 (Hazy Input)", "—", "—", "1.0894", "0.4945", "7.2446", "59.95"],
        ["DEA-Net Base (Chen et al., 2024)", "0% (冻结主干)", "0 FLOPs", "0.4163", "0.6113", "7.6116", "66.80"],
        ["BiLaLoRA (Zhang et al., 2026)", "3.03% ({1,7,8})", "旁路 / 可合并", "__0.2418__", "**0.6495**", "**7.6955**", "__67.91__"],
        ["**PRISM (本文, {6,7,8} + Conv-EVA)**", "**3.03% (110,595)**", "**0 参数 / 0 FLOPs**", "**0.2225**", "__0.6485__", "__7.6419__", "**68.26**"],
    ]
    add_three_line_table(
        doc,
        caption="**表 1：课程指定 10 幅真实雾天图像（Course-10 Benchmark）总体定量对比。** 粗体表示最优，下划线表示次优。",
        headers=t1_headers,
        rows=t1_rows,
        col_widths_dxa=[2650, 1500, 1500, 940, 940, 940, 940],
        alignments=["left", "center", "center", "center", "center", "center", "center"],
        font_size_pt=9.0,
        midrules_after={2},
    )

    # Table 2: Full 40-row Course-10 Per-Image Comparison Table with Vertical Merging on Cols 0 & 1
    t2_headers = [["图像名称", "场景特征", "对比方法", "FADE↓", "BIQME↑", "Entropy↑", "MUSIQ↑"]]
    t2_rows = [
        ["bench", "林间长椅逆光薄雾", "有雾原图", "1.2500", "0.4394", "7.2313", "63.65"],
        ["bench", "林间长椅逆光薄雾", "DEA-Net Base", "0.4025", "0.6210", "7.8187", "66.91"],
        ["bench", "林间长椅逆光薄雾", "BiLaLoRA", "__0.2266__", "**0.6733**", "**7.8882**", "**69.04**"],
        ["bench", "林间长椅逆光薄雾", "**PRISM (本文)**", "**0.2057**", "__0.6678__", "__7.8341__", "__68.91__"],
        ["cones", "道路锥桶景深跃变", "有雾原图", "0.6791", "0.5060", "7.2409", "63.51"],
        ["cones", "道路锥桶景深跃变", "DEA-Net Base", "0.1705", "0.6357", "7.5260", "__69.73__"],
        ["cones", "道路锥桶景深跃变", "BiLaLoRA", "__0.1105__", "**0.6597**", "**7.6466**", "**70.45**"],
        ["cones", "道路锥桶景深跃变", "**PRISM (本文)**", "**0.1064**", "__0.6509__", "__7.5812__", "69.69"],
        ["dubai", "航拍高层城市霾雾", "有雾原图", "1.2706", "0.4036", "6.9729", "62.80"],
        ["dubai", "航拍高层城市霾雾", "DEA-Net Base", "0.4469", "0.6125", "7.7101", "70.62"],
        ["dubai", "航拍高层城市霾雾", "BiLaLoRA", "__0.2656__", "__0.6737__", "**7.8518**", "**72.44**"],
        ["dubai", "航拍高层城市霾雾", "**PRISM (本文)**", "**0.2536**", "**0.6821**", "__7.8292__", "__72.01__"],
        ["forest", "密林高湿暗部雾气", "有雾原图", "0.4235", "0.5501", "7.4214", "68.07"],
        ["forest", "密林高湿暗部雾气", "DEA-Net Base", "0.3096", "0.6265", "7.7652", "__68.83__"],
        ["forest", "密林高湿暗部雾气", "BiLaLoRA", "__0.2029__", "**0.6762**", "**7.8392**", "**68.90**"],
        ["forest", "密林高湿暗部雾气", "**PRISM (本文)**", "**0.1813**", "__0.6735__", "__7.7875__", "68.10"],
        ["herzeliya", "海岸城市远景薄雾", "有雾原图", "0.7189", "0.5054", "7.3930", "__68.94__"],
        ["herzeliya", "海岸城市远景薄雾", "DEA-Net Base", "0.3659", "0.6346", "__7.7381__", "**69.67**"],
        ["herzeliya", "海岸城市远景薄雾", "BiLaLoRA", "__0.2231__", "__0.6534__", "**7.7499**", "67.99"],
        ["herzeliya", "海岸城市远景薄雾", "**PRISM (本文)**", "**0.2076**", "**0.6589**", "7.7367", "68.90"],
        ["neighbour", "浓雾街区建筑群", "有雾原图", "1.7757", "0.5198", "6.7027", "32.25"],
        ["neighbour", "浓雾街区建筑群", "DEA-Net Base", "0.5631", "__0.5681__", "7.1810", "52.10"],
        ["neighbour", "浓雾街区建筑群", "BiLaLoRA", "__0.2575__", "0.5647", "__7.1920__", "__56.56__"],
        ["neighbour", "浓雾街区建筑群", "**PRISM (本文)**", "**0.2139**", "**0.5870**", "**7.2005**", "**58.14**"],
        ["pumpkins", "田间南瓜高饱和色", "有雾原图", "0.6605", "0.5435", "7.6484", "67.50"],
        ["pumpkins", "田间南瓜高饱和色", "DEA-Net Base", "0.3990", "0.5951", "**7.6924**", "72.08"],
        ["pumpkins", "田间南瓜高饱和色", "BiLaLoRA", "__0.2377__", "**0.6238**", "__7.6149__", "__73.19__"],
        ["pumpkins", "田间南瓜高饱和色", "**PRISM (本文)**", "**0.2314**", "__0.6175__", "7.4782", "**73.83**"],
        ["road", "郊区公路连续景深", "有雾原图", "1.0101", "0.4457", "7.3127", "60.24"],
        ["road", "郊区公路连续景深", "DEA-Net Base", "0.3921", "0.5897", "__7.7264__", "__63.79__"],
        ["road", "郊区公路连续景深", "BiLaLoRA", "__0.2094__", "__0.6810__", "**7.7695**", "**63.83**"],
        ["road", "郊区公路连续景深", "**PRISM (本文)**", "**0.1731**", "**0.6820**", "7.7186", "63.66"],
        ["tiananmen", "天安门浓雾大天空", "有雾原图", "1.7071", "0.6038", "7.6644", "56.72"],
        ["tiananmen", "天安门浓雾大天空", "DEA-Net Base", "0.7363", "0.6315", "7.5070", "73.42"],
        ["tiananmen", "天安门浓雾大天空", "BiLaLoRA", "__0.4602__", "**0.6603**", "**7.8098**", "__75.02__"],
        ["tiananmen", "天安门浓雾大天空", "**PRISM (本文)**", "**0.4433**", "__0.6395__", "__7.6684__", "**76.75**"],
        ["train", "铁轨蒸汽碎石纹理", "有雾原图", "1.3982", "0.4277", "6.8585", "55.79"],
        ["train", "铁轨蒸汽碎石纹理", "DEA-Net Base", "0.3772", "0.5985", "7.4515", "60.85"],
        ["train", "铁轨蒸汽碎石纹理", "BiLaLoRA", "__0.2241__", "**0.6291**", "**7.5930**", "__61.69__"],
        ["train", "铁轨蒸汽碎石纹理", "**PRISM (本文)**", "**0.2087**", "__0.6255__", "__7.5848__", "**62.63**"],
    ]
    add_h2(doc, "5.2  全部 10 幅图像逐图全量评测与十边形性能包络")
    add_body_p(
        doc,
        "表 2 给出了 Course-10 基准全部 10 幅图像 × 4 种方法的 40 行完整逐图评测记录。为直观展现跨场景泛化规律，图 5 进一步绘制了 Course-10 十边形性能包络雷达图：在衡量去雾深度的无雾指数（1 − FADE↑）维度上，PRISM 的青绿色实线包络全面覆盖所有对比方法并囊括 10/10 单图最优；在整图感知质量（MUSIQ↑）维度上，PRISM 在天安门（76.75）、南瓜林（73.83）、街区（58.14）与火车（62.63）等典型雾景形成显著领先，仅在细碎树叶与局部过对比场景中维持稳健自然的视觉平衡。",
    )
    add_three_line_table(
        doc,
        caption="**表 2：课程指定全部 10 幅真实雾天图像（Course-10）逐图全量定量对比（40 行完整评测记录）。** PRISM 在全部 10 幅图像上均取得最低（最优）FADE 残雾得分。",
        headers=t2_headers,
        rows=t2_rows,
        col_widths_dxa=[1250, 1950, 1650, 1140, 1140, 1140, 1140],
        alignments=["center", "center", "center", "center", "center", "center", "center"],
        font_size_pt=7.9,
        merge_cols=[0, 1],
        midrules_after={4, 8, 12, 16, 20, 24, 28, 32, 36},
        nowrap_cols={0, 1, 2, 3, 4, 5, 6},
        keep_together=True,
        cell_pad_v_dxa=5,
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig_course10_radar.png",
        width_cm=16.0,
        caption_text=(
            "**图 5：Course-10 十幅真实雾图多维性能包络十边形雷达图。** "
            "(a) 无雾通透度指数（1 − FADE ↑）：PRISM 在全部 10 个顶点上包络最广且均达全场最高；"
            "(b) 整图多尺度感知质量（MUSIQ ↑）：PRISM（深青绿实线）在 10 图均分上达最高 68.26，顶点数旁标示各场景最优方法的具体实测数值（青绿为 PRISM 最优，蓝色为 BiLaLoRA）。"
        ),
    )

    add_h2(doc, "5.3  全部 10 幅课程图像去雾视觉效果矩阵")
    add_body_p(
        doc,
        "图 6 与图 7 以高清极简排版全景展示全部 10 幅课程指定雾图在“有雾原图、DEA-Net Base、BiLaLoRA、PRISM (本文)”四种设定下的视觉对比结果，每幅子图保留原始主视野并标注实测 `FADE / MUSIQ`（最优项加粗）。可以看到，预训练 `DEA-Net Base` 在 `bench`、`dubai`、`neighbour`、`tiananmen` 与 `train` 等真实浓雾场景中遗留明显灰白雾幕；`BiLaLoRA` 提升了去雾强度但在高对比边界仍偶见局部偏暗；**PRISM** 凭借深层瓶颈 `{6,7,8}` 协同调整，在 10 个场景中均呈现通透的远景能见度与自然的色彩饱和度。",
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig3_course10_part1.png",
        width_cm=16.2,
        caption_text=(
            "**图 6：课程指定雾天图像高清全景对比矩阵（第一组：bench、cones、dubai、forest、herzeliya）。** "
            "每列依次为有雾原图、DEA-Net Base、BiLaLoRA 与 PRISM (本文)，子图下方标注实测 FADE↓ / MUSIQ↑（最优项加粗）。"
        ),
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig4_course10_part2.png",
        width_cm=16.2,
        caption_text=(
            "**图 7：课程指定雾天图像高清全景对比矩阵（第二组：neighbour、pumpkins、road、tiananmen、train）。** "
            "PRISM 在浓雾街区 neighbour（FADE 0.2139, MUSIQ 58.14）、天安门远景 tiananmen（MUSIQ 76.75）及铁轨 train（MUSIQ 62.63）上全面领先。"
        ),
    )

    add_h2(doc, "5.4  代表性课程图像局部放大与 8 倍空间差分分析")
    add_body_p(
        doc,
        "为进一步剖析 PRISM 相较于基线方法的空间修正机理，图 8 对 `neighbour`、`tiananmen` 与 `pumpkins` 三幅代表性课程图像进行了感兴趣区域局部放大及 `8× |PRISM - BiLaLoRA|` 绝对差分可视化：（1）在浓雾住宅区 `neighbour` 中，远景红顶建筑与树木被浓重灰白雾层包裹（原图 FADE 高达 1.7757），PRISM 将远景屋顶轮廓与树冠层次清晰剥离，FADE 较 BiLaLoRA 进一步下降 16.9%（0.2575→0.2139），MUSIQ 提升 +1.58 分；（2）在 `tiananmen` 场景中，差分图显示 PRISM 的修正能量集中于城楼红墙檐角与天空交界处，避免了浅层过增强常见的天空色斑，使 MUSIQ 大幅提升 +1.73 分（达 76.75）；（3）在 `pumpkins` 中，PRISM 在彻底去除远景林间薄雾的同时，精准保护了前景橙色南瓜的自然光泽与地面阴影过渡。",
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig5_course10_zoom.png",
        width_cm=16.2,
        caption_text=(
            "**图 8：Course-10 典型场景（neighbour、tiananmen、pumpkins）局部放大与 8 倍空间差分分析。** "
            "第 1 列标示红色放大框位置，第 2–4 列对比局部细节复原质量，第 5 列展示 8× |PRISM − BiLaLoRA| 像素级差异分布。"
        ),
    )

    # --------------------------------------------------------------------------
    # SECTION 6: 6 公开基准大规模对比与受控消融实验
    # --------------------------------------------------------------------------
    add_h1(doc, "6  公开基准大规模对比与受控消融实验")
    add_h2(doc, "6.1  大规模真实基准定量对比")
    add_body_p(
        doc,
        "除课程指定图像外，我们在国际公认的大规模真实雾图基准 **RTTS**（Li et al., 2019，包含 4,322 幅真实有雾图像）和经典自然图像去雾基准 **Fattal**（Fattal, 2014，31 幅高分辨率图像）上，将 PRISM 与 17 种代表性去雾方法进行了系统对比，涵盖 8 种有监督/通用复原方法（MSBDN、DeHamer、C2PNet、DEA-Net、PromptIR、DiffUIR、MoCE-IR、FoundIR）与 9 种无监督/真实域适配方法（DAD、PSD、D4、RIDCP、KANet、CoA、IPC、PHATNet、BiLaLoRA），完整结果见表 3。",
    )
    add_body_p(
        doc,
        "定量对比表明：（1）在包含 4,322 幅真实雾图的 **RTTS** 基准上，PRISM 在全部四项无参考指标上一致取得全场最优：FADE 达到 **0.6241±0.0055**，较 BiLaLoRA 论文报告值（0.752）显著降低 **17.0%**，较 BiLaLoRA 官方发布权重（0.7544）降低 **17.3%**，较同协议固定参考位置 `{1,7,8}+EVA`（0.7172±0.0078）降低 **13.0%**；同时 BIQME 提升至 **0.6191±0.0002**、信息熵达到 **7.5799**、MUSIQ 提升至 **61.79±0.05**。（2）在 **Fattal（n=31）** 基准上，PRISM 同样取得全场最低的 FADE（**0.2714±0.0019**）与最高的 BIQME（**0.6248±0.0003**）；在 MUSIQ 指标上，PRISM 的三种子均值（68.02±0.01）优于 BiLaLoRA 论文报告值（67.92）及同协议参考位置+EVA（67.71±0.09）。BiLaLoRA 官方发布权重在 Fattal 上表现为高 MUSIQ（68.59）与较高残雾 FADE（0.3070）的帕累托权衡，且来自单一种子检查点；PRISM 在三种子受控协议下将 Fattal 的 FADE 较该官方权重降低 **11.6%** 并将 BIQME 提升 +0.0042，展现出更均衡的去雾彻底性与跨种子稳定性。",
    )

    # Table 3: Full 20-row Main Comparison Table (9410 dxa total, zero wrapping)
    t3_headers = [
        ["对比方法 / 模型类别", "RTTS FADE↓", "RTTS BIQME↑", "RTTS Entropy↑", "RTTS MUSIQ↑", "Fattal FADE↓", "Fattal BIQME↑", "Fattal Entropy↑", "Fattal MUSIQ↑"]
    ]
    t3_rows = [
        ["MSBDN (Dong et al., 2020)†", "1.483", "0.549", "7.273", "52.93", "0.613", "0.555", "7.408", "63.71"],
        ["DeHamer (Guo et al., 2022)†", "1.806", "0.542", "7.215", "52.90", "0.756", "0.552", "7.411", "64.31"],
        ["C2PNet (Zheng et al., 2023)†", "2.050", "0.531", "7.168", "54.18", "0.720", "0.551", "7.399", "65.00"],
        ["DEA-Net (Chen et al., 2024)†", "1.781", "0.541", "7.196", "53.13", "0.696", "0.557", "7.423", "64.23"],
        ["PromptIR (Potlapalli et al., 2023)†", "1.765", "0.546", "7.189", "53.80", "0.668", "0.558", "7.498", "64.96"],
        ["DiffUIR (Zheng et al., 2024)†", "2.132", "0.531", "7.172", "54.81", "0.871", "0.537", "7.373", "64.95"],
        ["MoCE-IR (Zamfir et al., 2025)†", "1.922", "0.539", "7.191", "54.21", "0.678", "0.557", "7.409", "65.01"],
        ["FoundIR (Li et al., 2025)†", "1.760", "0.553", "7.275", "54.96", "0.759", "0.560", "7.376", "65.83"],
        ["DAD (Shao et al., 2020)†", "1.131", "0.561", "7.413", "49.34", "0.487", "0.589", "7.487", "59.38"],
        ["PSD (Chen et al., 2021)†", "1.143", "0.524", "7.276", "52.81", "0.438", "0.554", "7.463", "63.80"],
        ["D4 (Yang et al., 2022)†", "1.404", "0.556", "7.179", "53.57", "0.457", "0.537", "7.372", "64.14"],
        ["RIDCP (Wu et al., 2023)†", "0.955", "0.600", "7.541", "59.14", "0.396", "0.604", "7.468", "66.10"],
        ["KANet (Feng et al., 2024)†", "0.870", "0.583", "7.517", "54.54", "0.338", "0.560", "7.527", "65.35"],
        ["CoA (Ma et al., 2025)†", "0.859", "0.593", "__7.579__", "53.43", "0.314", "0.618", "7.585", "63.38"],
        ["IPC (Fu et al., 2025)†", "1.105", "0.592", "7.469", "59.61", "0.368", "0.593", "7.471", "67.58"],
        ["PHATNet (Tsai et al., 2025)†", "0.845", "0.585", "7.349", "56.43", "0.331", "0.604", "7.498", "66.87"],
        ["BiLaLoRA 论文 (Zhang et al., 2026)†", "0.752", "0.611", "7.576", "__61.77__", "__0.281__", "0.607", "7.541", "67.92"],
        ["BiLaLoRA 官方权重 (本地实测)", "0.7544", "0.6107", "7.5757", "__61.77__", "0.3070", "0.6206", "**7.6124**", "**68.59**"],
        ["参考位置 {1,7,8} + Conv-EVA", "__0.7172±0.0078__", "__0.6112±0.0008__", "7.5785", "61.65±0.14", "0.2847±0.0009", "__0.6217±0.0004__", "__7.5992__", "67.71±0.09"],
        ["**PRISM (本文, {6,7,8})**", "**0.6241±0.0055**", "**0.6191±0.0002**", "**7.5799**", "**61.79±0.05**", "**0.2714±0.0019**", "**0.6248±0.0003**", "7.5942", "__68.02±0.01__"],
    ]
    add_three_line_table(
        doc,
        caption="**表 3：RTTS（n=4,322）与 Fattal（n=31）基准上与 17 种代表性去雾方法的完整定量对比。** † 表示引自 BiLaLoRA 原文，最后两行为三种子（3407/3408/3409）均值±标准差。",
        headers=t3_headers,
        rows=t3_rows,
        col_widths_dxa=[1970, 1080, 1080, 710, 865, 1080, 1080, 710, 835],
        alignments=["left"] + ["center"] * 8,
        font_size_pt=8.0,
        midrules_after={8, 17, 18},
        nowrap_cols=set(range(9)),
        keep_together=True,
        cell_pad_v_dxa=14,
    )

    add_h2(doc, "6.2  配对统计显著性与去重灵敏度验证")
    add_body_p(
        doc,
        "为排除随机波动对实验结论的干扰，我们在每幅评测图像上先对三个随机种子的预测求均值，再计算 `PRISM {6,7,8}` 相对于同协议 `参考位置 {1,7,8}+EVA` 的逐图配对差值，并执行 10,000 次非参数配对自助法（Paired Bootstrap，种子 20260414）重采样估计 95% 置信区间（表 4）。同时，针对像素级哈希审计发现的无标签训练雾图 `0700912.jpg` 与 Fattal 测试图 `house_input.png` 重合问题，表 4 与表 5 同步报告了剔除 `house` 后的 `Fattal (n=30)` 去重核验结果。数据显示，FADE、BIQME 与 MUSIQ 的 95% 置信区间均严格不包含 0，证明 PRISM 的性能提升具有高度的统计显著性。",
    )

    # Table 4: Paired Bootstrap 95% CI & Per-image Win Rate (with Vertical Merge on Col 0)
    t4_headers = [["评测基准集合", "评价指标", "配对均值变化 (Δ)", "种子间标准差", "图像配对 95% 置信区间", "逐图改善比例"]]
    t4_rows = [
        ["RTTS 全集 (n = 4,322)", "FADE↓", "**-0.0931**", "0.0108", "[-0.0955, -0.0909]", "**99.7%**"],
        ["RTTS 全集 (n = 4,322)", "BIQME↑", "**+0.0079**", "0.0010", "[+0.0075, +0.0083]", "**78.9%**"],
        ["RTTS 全集 (n = 4,322)", "Entropy↑", "+0.0014", "0.0045", "[-0.0013, +0.0041]", "46.3%"],
        ["RTTS 全集 (n = 4,322)", "MUSIQ↑", "**+0.14**", "0.09", "[+0.10, +0.18]", "**59.9%**"],
        ["Fattal 原始集 (n = 31)", "FADE↓", "**-0.0133**", "0.0024", "[-0.0190, -0.0082]", "**91.4%**"],
        ["Fattal 原始集 (n = 31)", "BIQME↑", "**+0.0031**", "0.0001", "[+0.0002, +0.0057]", "**73.1%**"],
        ["Fattal 原始集 (n = 31)", "Entropy↑", "-0.0050", "0.0017", "[-0.0178, +0.0126]", "34.4%"],
        ["Fattal 原始集 (n = 31)", "MUSIQ↑", "**+0.31**", "0.08", "[+0.07, +0.52]", "**75.3%**"],
        ["Fattal 去重集 (n = 30)", "FADE↓", "**-0.0135**", "0.0024", "[-0.0193, -0.0081]", "**91.1%**"],
        ["Fattal 去重集 (n = 30)", "BIQME↑", "**+0.0033**", "0.0001", "[+0.0003, +0.0059]", "**75.6%**"],
        ["Fattal 去重集 (n = 30)", "Entropy↑", "-0.0044", "0.0017", "[-0.0179, +0.0136]", "35.6%"],
        ["Fattal 去重集 (n = 30)", "MUSIQ↑", "**+0.30**", "0.08", "[+0.05, +0.52]", "**74.4%**"],
    ]
    add_three_line_table(
        doc,
        caption="**表 4：PRISM {6,7,8} 相对于同协议参考位置 {1,7,8}+EVA 的逐图配对差值、10,000 次自助法 95% 置信区间及去重核验。**",
        headers=t4_headers,
        rows=t4_rows,
        col_widths_dxa=[2150, 1200, 1450, 1300, 1950, 1360],
        alignments=["center"] * 6,
        font_size_pt=8.5,
        merge_cols=[0],
        midrules_after={4, 8},
        nowrap_cols=set(range(6)),
    )

    # Table 5: Complete 3-Seed Breakdown Table (with Vertical Merge on Col 0)
    t5_headers = [["评测集合", "模型配置与随机种子", "FADE↓", "BIQME↑", "Entropy↑", "MUSIQ↑"]]
    t5_rows = [
        ["RTTS 全集 (n = 4,322)", "参考位置 {1,7,8}+EVA / Seed 3407", "0.7142", "0.6115", "7.5786", "61.75"],
        ["RTTS 全集 (n = 4,322)", "参考位置 {1,7,8}+EVA / Seed 3408", "0.7261", "0.6103", "7.5764", "61.49"],
        ["RTTS 全集 (n = 4,322)", "参考位置 {1,7,8}+EVA / Seed 3409", "0.7113", "0.6117", "7.5805", "61.71"],
        ["RTTS 全集 (n = 4,322)", "**PRISM (本文, {6,7,8}) / Seed 3407**", "**0.6304**", "**0.6190**", "7.5793", "**61.82**"],
        ["RTTS 全集 (n = 4,322)", "**PRISM (本文, {6,7,8}) / Seed 3408**", "**0.6212**", "**0.6193**", "**7.5826**", "**61.73**"],
        ["RTTS 全集 (n = 4,322)", "**PRISM (本文, {6,7,8}) / Seed 3409**", "**0.6207**", "**0.6189**", "7.5778", "**61.81**"],
        ["Fattal 原始集 (n = 31)", "参考位置 {1,7,8}+EVA / Seed 3407", "0.2840", "0.6215", "7.5965", "67.80"],
        ["Fattal 原始集 (n = 31)", "参考位置 {1,7,8}+EVA / Seed 3408", "0.2858", "0.6222", "7.6026", "67.63"],
        ["Fattal 原始集 (n = 31)", "参考位置 {1,7,8}+EVA / Seed 3409", "0.2844", "0.6215", "7.5985", "67.72"],
        ["Fattal 原始集 (n = 31)", "**PRISM (本文, {6,7,8}) / Seed 3407**", "**0.2735**", "**0.6247**", "7.5933", "**68.04**"],
        ["Fattal 原始集 (n = 31)", "**PRISM (本文, {6,7,8}) / Seed 3408**", "**0.2710**", "**0.6252**", "7.5973", "**68.01**"],
        ["Fattal 原始集 (n = 31)", "**PRISM (本文, {6,7,8}) / Seed 3409**", "**0.2698**", "**0.6246**", "7.5919", "**68.01**"],
        ["Fattal 去重集 (n = 30)", "参考位置 {1,7,8}+EVA (3种子均值±标准差)", "0.2891±0.0010", "0.6213±0.0004", "7.5907±0.0032", "67.60±0.09"],
        ["Fattal 去重集 (n = 30)", "**PRISM (本文, {6,7,8}, 3种子均值±标准差)**", "**0.2756±0.0019**", "**0.6246±0.0003**", "7.5863±0.0028", "**67.90±0.01**"],
    ]
    add_three_line_table(
        doc,
        caption="**表 5：三个独立随机种子（3407、3408、3409）下的逐模型分种子得分与 Fattal（n=30）去重核验均值表。**",
        headers=t5_headers,
        rows=t5_rows,
        col_widths_dxa=[1850, 2960, 1150, 1150, 1150, 1150],
        alignments=["center", "left", "center", "center", "center", "center"],
        font_size_pt=8.4,
        merge_cols=[0],
        midrules_after={6, 12},
        nowrap_cols=set(range(6)),
    )

    add_h2(doc, "6.3  选层规则受控消融与效率-稳定性权衡")
    add_body_p(
        doc,
        "为严格隔离选层准则本身的贡献，表 6 在完全相同的第 5 轮 Conv-EVA 共享快照与后续 25 轮纯 H2C 训练协议下，对比了四种等容量（110,595 参数）选层规则：（1）预置参考位置 `{1,7,8}`；（2）无约束 H2C 最优组合 `{5,6,8}`；（3）连续双层门控排序（种子 3407/3408/3409 分别选出 `{2,6,8}`、`{1,6,7}`、`{1,4,8}`）；（4）本文 PRISM 选出的 `{6,7,8}`。表 6 验证了两项核心结论：第一，纯最小化 H2C 的 `{5,6,8}` 虽取得低 FADE，但在 RTTS-128 与 Dense-Haze 上的整图 MUSIQ 分别跌至 61.18 与 52.65，证实了浅层过增强诱发的“语义—感知失配”；第二，连续门控排序在 RTTS-128 上的 FADE 标准差达 0.0474、MUSIQ 标准差达 0.54，而 PRISM 将二者标准差分别收敛至 **0.0053** 与 **0.05**（缩小近一个数量级），同时取得最高 MUSIQ（**62.57±0.05**）。图 9 进一步以三维气泡权衡图呈现 PRISM 在计算效率—精度与选择稳定性上的帕累托支配优势。",
    )

    # Table 6: Equal-budget Layer Selection Ablation Table
    t6_headers = [
        ["等容量选层规则（5+25轮）", "选中瓶颈层索引", "RTTS-128 FADE↓", "RTTS-128 MUSIQ↑", "RTTS-128 NIQE↓", "Dense-Haze FADE↓", "Dense-Haze MUSIQ↑", "Dense-Haze NIQE↓"]
    ]
    t6_rows = [
        ["预置参考位置 (Reference)", "固定 {1,7,8}", "0.6991±0.0074", "62.26±0.15", "4.90±0.02", "0.7475±0.0025", "**54.80±0.06**", "5.41±0.02"],
        ["纯 H2C 最小化 (H2C-min)", "固定 {5,6,8}", "**0.6122±0.0145**", "61.18±0.14", "**4.64±0.02**", "**0.5822±0.0213**", "52.65±0.25", "**4.72±0.04**"],
        ["连续双层门控排序 (Gate)", "{2,6,8}/{1,6,7}/{1,4,8}", "0.6435±0.0474", "__62.28±0.54__", "__4.88±0.02__", "0.6544±0.0339", "53.24±0.58", "__4.99±0.18__"],
        ["**PRISM (本文, 可行域筛选)**", "**一致锁定 {6,7,8}**", "__0.6147±0.0053__", "**62.57±0.05**", "__4.88±0.02__", "__0.6343±0.0036__", "__53.58±0.05__", "5.09±0.05"],
    ]
    add_three_line_table(
        doc,
        caption="**表 6：相同第 5 轮共享快照、相同训练预算（5+25 轮）与相同参数量（110,595）下的选层规则受控消融实验。**",
        headers=t6_headers,
        rows=t6_rows,
        col_widths_dxa=[1710, 1740, 1140, 960, 860, 1140, 960, 900],
        alignments=["left", "center", "center", "center", "center", "center", "center", "center"],
        font_size_pt=7.8,
        midrules_after={3},
        nowrap_cols=set(range(8)),
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig_efficiency_stability_bubble.png",
        width_cm=16.2,
        caption_text=(
            "**图 9：推理计算效率与选层稳定性的三维气泡综合权衡图。** "
            "(a) RTTS 全集计算开销—残雾消除权衡：气泡面积编码参数量（3.65M–132.4M），PRISM 经重参数化无损融合后仅需 3.65M 参数与 32.2G FLOPs（零额外开销），同时取得最低 FADE 0.6241；"
            "(b) RTTS-128 选层精度—跨种子稳定性权衡：气泡面积编码跨三随机种子标准差，PRISM 以最低 FADE 方差（Std = 0.0053，较连续双层门控 0.0474 降低 8.9 倍）同步登顶最高整图 MUSIQ（62.57）。"
        ),
    )

    add_h2(doc, "6.4  自然场景（Fattal）定性局部放大与 8 倍空间差分对比")
    add_body_p(
        doc,
        "图 10 展示了高分辨率自然雾天基准 Fattal 中 `swan`、`snow` 与 `buildings` 三组经典场景在种子 3407 下的 2:1 局部放大及 `8× |PRISM − 参考位置|` 绝对差分对比。在近景水面、雪地反光与多层建筑外墙纹理区域，PRISM 在较参考位置 `{1,7,8}` 显著加重去雾彻底度的同时（FADE 稳定下降），整图感知自然度 MUSIQ 同步净提升 +0.62 至 +1.13 分，且局部天空与高光边界平滑自然、无过增强振铃。",
    )
    add_figure_block(
        doc,
        ASSETS_DIR / "fig6_fattal_qualitative.png",
        width_cm=16.2,
        caption_text=(
            "**图 10：Fattal 基准典型场景（swan、snow、buildings）2:1 局部放大与 8 倍空间差分对比（种子 3407）。** "
            "各列下方标注整图 FADE↓ / MUSIQ↑（最优项加粗），最右列展示 8× |PRISM − 参考位置| 绝对差异及整图 MUSIQ 净增益。"
        ),
    )

    add_h2(doc, "6.5  评训分离机理验证、Dense-Haze 泛化与边界诊断")
    add_body_p(
        doc,
        "表 7(a) 进一步检验了将不同分辨率 MUSIQ 直接作为反向传播损失项加入训练的效果：在 128×128 裁剪块上反传 MUSIQ 会因感受野缺失整图构图而导致 FADE 恶化（0.5493→0.5788）；在 224×224 缩放整图上反传甚至使 MUSIQ 倒退至 61.48；在 1024 上限高分辨率整图上反传虽能提升 MUSIQ，但代价是 FADE 反弹至 0.5899、NIQE 恶化至 4.71 且显存开销剧增，有力验证了**评训分离**机制的必要性。表 7(b) 与图 11 给出了极端浓雾 **Dense-Haze（55 对）** 与 Fattal `train` 边界案例诊断：PRISM 在 Dense-Haze 上将全参考 **PSNR↑** 提升至 **12.2210 dB**（+0.2914 dB），FADE 降低 15.1%；在 `train` 样例中，PRISM 将残雾 FADE 降至最低（0.1722），但深层除雾显露的铁轨碎石高频颗粒使偏好平滑背景的 MUSIQ 产生 -0.95 分波动，客观揭示了极端高频纹理场景的评价边界。",
    )

    # Table 7: Supervision Objective Diagnostic & Dense-Haze Generalization (with Vertical Merge on Col 0)
    t7_headers = [["实验阶段与评测基准", "训练目标或选层规则配置", "FADE↓", "MUSIQ↑", "NIQE↓", "PSNR (dB)↑", "SSIM↑"]]
    t7_rows = [
        ["(a) 评分反传诊断: RTTS-128", "仅原版 H2C 损失 (30+5 轮纯 H2C 对照)", "**0.5493±0.0077**", "61.60±0.21", "__4.66±0.03__", "—", "—"],
        ["(a) 评分反传诊断: RTTS-128", "H2C + 128×128 局部裁剪块 MUSIQ 反传", "0.5788±0.0060", "__61.86±0.18__", "4.69±0.03", "—", "—"],
        ["(a) 评分反传诊断: RTTS-128", "H2C + 224×224 缩放整图 MUSIQ 反传", "__0.5572±0.0056__", "61.48±0.15", "**4.64±0.02**", "—", "—"],
        ["(a) 评分反传诊断: RTTS-128", "H2C + 1024 上限高分辨率整图 MUSIQ 反传", "0.5899±0.0088", "**62.23±0.25**", "4.71±0.04", "—", "—"],
        ["(a) 评分反传诊断: Dense-Haze (55)", "仅原版 H2C 损失 (30+5 轮纯 H2C 对照)", "**0.5341±0.0184**", "53.51±0.29", "**4.55±0.11**", "__12.3085±0.1565__", "0.4994±0.0016"],
        ["(a) 评分反传诊断: Dense-Haze (55)", "H2C + 128×128 局部裁剪块 MUSIQ 反传", "0.5701±0.0193", "__53.67±0.11__", "4.71±0.13", "12.2809±0.1532", "__0.5038±0.0020__"],
        ["(a) 评分反传诊断: Dense-Haze (55)", "H2C + 224×224 缩放整图 MUSIQ 反传", "__0.5410±0.0128__", "52.94±0.17", "__4.62±0.09__", "**12.4143±0.1174**", "0.5014±0.0012"],
        ["(a) 评分反传诊断: Dense-Haze (55)", "H2C + 1024 上限高分辨率整图 MUSIQ 反传", "0.5863±0.0254", "**54.44±0.24**", "4.80±0.18", "12.2664±0.1815", "**0.5053±0.0027**"],
        ["(b) 主方案泛化: Dense-Haze (55)", "参考位置 {1,7,8} + Conv-EVA (5+25 轮)", "0.7475±0.0025", "**54.80±0.06**", "5.41±0.02", "11.9296±0.0409", "**0.5101±0.0014**"],
        ["(b) 主方案泛化: Dense-Haze (55)", "**PRISM (本文, {6,7,8}, 5+25 轮)**", "**0.6343±0.0036**", "53.58±0.05", "**5.09±0.05**", "**12.2210±0.0366**", "0.5063±0.0011"],
    ]
    add_three_line_table(
        doc,
        caption="**表 7：评训分离机理验证（直接反传不同分辨率 MUSIQ 梯度的副作用）与 Dense-Haze（n=55）成对浓雾基准综合诊断。**",
        headers=t7_headers,
        rows=t7_rows,
        col_widths_dxa=[1980, 2240, 1140, 920, 820, 1160, 1150],
        alignments=["center", "left", "center", "center", "center", "center", "center"],
        font_size_pt=7.9,
        merge_cols=[0],
        midrules_after={4, 8},
        nowrap_cols=set(range(7)),
    )

    add_figure_block(
        doc,
        ASSETS_DIR / "fig7_train_boundary.png",
        width_cm=16.2,
        caption_text=(
            "**图 11：Fattal 基准中 train 边界案例的局部放大与 8 倍差分诊断。** "
            "PRISM 将残雾指标 FADE 降至全场最低（0.1722 vs. 参考位置 0.1814），但彻底剥离前景浓重蒸汽后显露的铁轨枕木与碎石高频纹理引发了无参考 MUSIQ 的小幅回落（-0.95 分），揭示了深层除雾与背景平滑偏好的权衡边界。"
        ),
    )

    # --------------------------------------------------------------------------
    # SECTION 7: 7 结论与展望 + 参考文献
    # --------------------------------------------------------------------------
    add_h1(doc, "7  结论与展望")
    add_body_p(
        doc,
        "本文针对无配对真实图像去雾低秩适配中长期被忽视的“语义—感知失配”与双层连续门控高方差问题，提出了感知可行域引导的等容量快照筛选框架 **PRISM**。通过将卷积激活解释方差对齐（Conv-EVA）与第 5 轮共享快照上的无梯度 56 组合前向可行域筛选相结合，PRISM 将离散层位置决策与反向传播彻底解耦：在不改动原版 H2C 语义训练目标、不引入额外训练轮数的前提下，在全部三个独立随机种子下一致锁定帕累托最优深层瓶颈子集 `{6,7,8}`，将跨种子性能标准差缩小近一个数量级。微调后的低秩旁路仅占主干参数的 3.03%，并在推理期通过线性重参数化无损合并回原 `DEA-Net` 卷积核（零额外参数与零额外 FLOPs）。在涵盖 30,471 条逐图记录的四大基准（RTTS 4,322 图、Fattal 31/30 图、Dense-Haze 55 对及课程指定 Course-10 10 图）上，PRISM 均展现出优异的残雾消除能力与感知自然度。未来工作将进一步探索图像自适应的动态秩分配以及面向超高清视频实时去雾的时序一致性低秩子空间约束。本项目完整代码、权重与英文论文已同步发布于 GitHub（[NEUMelon/PRISM](https://github.com/NEUMelon/PRISM.git) | 英文论文：[PRISM_Submission_EN.pdf](https://github.com/NEUMelon/PRISM/blob/main/paper/PRISM_Submission_EN.pdf)）。",
    )

    add_h1(doc, "参考文献")
    references_list = [
        "[1] Ancuti, C. O., Ancuti, C., Sbert, M., and Timofte, R. (2019). Dense-Haze: A Benchmark for Image Dehazing with Dense-Haze and Haze-Free Images. In IEEE International Conference on Image Processing (ICIP), pages 1014–1018.",
        "[2] Blau, Y. and Michaeli, T. (2018). The Perception-Distortion Tradeoff. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 6228–6237.",
        "[3] Chen, Z., Wang, Y., Yang, Y., and Liu, D. (2021). PSD: Principled Synthetic-to-Real Dehazing Guided by Physical Priors. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 7180–7189.",
        "[4] Chen, Z., He, Z., and Lu, Z.-M. (2024). DEA-Net: Single Image Dehazing Based on Detail-Enhanced Convolution and Content-Guided Attention. IEEE Transactions on Image Processing, 33:1002–1015.",
        "[5] Choi, L. K., You, J., and Bovik, A. C. (2015). Referenceless Prediction of Perceptual Fog Density and Perceptual Image Defogging. IEEE Transactions on Image Processing, 24(11):3888–3901.",
        "[6] Dong, H., Pan, J., Xiang, L., Hu, Z., Zhang, X., Wang, F., and Yang, M.-H. (2020). Multi-Scale Boosted Dehazing Network With Dense Feature Fusion. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 2157–2167.",
        "[7] Fang, C., He, C., Xiao, F., Zhang, Y., Tang, L., Zhang, Y., Li, K., and Li, X. (2024). Real-world Image Dehazing with Coherence-based Pseudo Labeling and Cooperative Unfolding Network. In Advances in Neural Information Processing Systems (NeurIPS).",
        "[8] Fattal, R. (2014). Dehazing Using Color-Lines. ACM Transactions on Graphics (TOG), 34(1):13:1–13:14.",
        "[9] Feng, Y., Ma, L., Meng, X., Zhou, F., Liu, R., and Su, Z. (2024). Advancing Real-World Image Dehazing: Perspective, Modules, and Training. IEEE Transactions on Pattern Analysis and Machine Intelligence, 46(12):9303–9320.",
        "[10] Fu, J., Liu, S., Liu, Z., Guo, C.-L., Park, H., Wu, R., Wang, G., and Li, C. (2025). Iterative Predictor-Critic Code Decoding for Real-World Image Dehazing. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 12700–12709.",
        "[11] Gu, K., Tao, D., Qiao, J.-F., and Lin, W. (2018). Learning a No-Reference Quality Assessment Model of Enhanced Images With Big Data. IEEE Transactions on Neural Networks and Learning Systems, 29(4):1301–1313.",
        "[12] Guo, C.-L., Yan, Q., Anwar, S., Cong, R., Ren, W., and Li, C. (2022). Image Dehazing Transformer With Transmission-Aware 3D Position Embedding. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 5812–5820.",
        "[13] He, K., Sun, J., and Tang, X. (2009). Single Image Haze Removal Using Dark Channel Prior. In Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR), pages 1956–1963.",
        "[14] Hosu, V., Lin, H., Sziranyi, T., and Saupe, D. (2020). KonIQ-10k: An Ecologically Valid Database for Deep Learning of Blind Image Quality Assessment. IEEE Transactions on Image Processing, 29:4041–4056.",
        "[15] Hu, E. J., Shen, Y., Wallis, P., Allen-Zhu, Z., Li, Y., Wang, S., Wang, L., and Chen, W. (2022). LoRA: Low-Rank Adaptation of Large Language Models. In International Conference on Learning Representations (ICLR).",
        "[16] Ke, J., Wang, Q., Wang, Y., Milanfar, P., and Yang, F. (2021). MUSIQ: Multi-Scale Image Quality Transformer. In Proceedings of the IEEE/CVF International Conference on Computer Vision (ICCV), pages 5148–5157.",
        "[17] Li, B., Ren, W., Fu, D., Tao, D., Feng, D., Zeng, W., and Wang, Z. (2019). Benchmarking Single-Image Dehazing and Beyond. IEEE Transactions on Image Processing, 28(1):492–505.",
        "[18] Li, H., Chen, X., Dong, J., Tang, J., and Pan, J. (2025). FoundIR: Unleashing Million-Scale Training Data to Advance Foundation Models for Image Restoration. In Proceedings of the IEEE/CVF International Conference on Computer Vision (ICCV), pages 12626–12636.",
        "[19] Ma, L., Feng, Y., Zhang, Y., Liu, J., Wang, W., Chen, G., Xu, C., and Su, Z. (2025). CoA: Towards Real Image Dehazing via Compression-and-Adaptation. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 11197–11206.",
        "[20] Meng, F., Wang, Z., and Zhang, M. (2024). PiSSA: Principal Singular Values and Singular Vectors Adaptation of Large Language Models. In Advances in Neural Information Processing Systems (NeurIPS).",
        "[21] Mittal, A., Soundararajan, R., and Bovik, A. C. (2013). Making a \"Completely Blind\" Image Quality Analyzer. IEEE Signal Processing Letters, 20(3):209–212.",
        "[22] Paischer, F., Hauzenberger, L., Schmied, T., Alkin, B., Deisenroth, M. P., and Hochreiter, S. (2025). Parameter Efficient Fine-tuning via Explained Variance Adaptation. In Advances in Neural Information Processing Systems (NeurIPS).",
        "[23] Potlapalli, V., Zamir, S. W., Khan, S. H., and Khan, F. S. (2023). PromptIR: Prompting for All-in-One Image Restoration. In Advances in Neural Information Processing Systems (NeurIPS).",
        "[24] Radford, A., Kim, J. W., Hallacy, C., Ramesh, A., Goh, G., Agarwal, S., Sastry, G., Askell, A., Mishkin, P., Clark, J., Krueger, G., and Sutskever, I. (2021). Learning Transferable Visual Models From Natural Language Supervision. In Proceedings of the 38th International Conference on Machine Learning (ICML), pages 8748–8763.",
        "[25] Shao, Y., Li, L., Ren, W., Gao, C., and Sang, N. (2020). Domain Adaptation for Image Dehazing. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 2808–2817.",
        "[26] Tan, M., Chen, B., Pang, R., Vasudevan, V., Sandler, M., Howard, A., and Le, Q. V. (2019). MnasNet: Platform-Aware Neural Architecture Search for Mobile. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 2820–2828.",
        "[27] Tsai, F., Peng, Y., Lin, Y., and Lin, C. (2025). PHATNet: A Physics-Guided Haze Transfer Network for Domain-Adaptive Real-World Image Dehazing. In Proceedings of the IEEE/CVF International Conference on Computer Vision (ICCV), pages 5591–5600.",
        "[28] Wang, S., Yu, L., and Li, J. (2024a). LoRA-GA: Low-Rank Adaptation with Gradient Approximation. In Advances in Neural Information Processing Systems (NeurIPS).",
        "[29] Wang, Z., Bovik, A. C., Sheikh, H. R., and Simoncelli, E. P. (2004). Image Quality Assessment: From Error Visibility to Structural Similarity. IEEE Transactions on Image Processing, 13(4):600–612.",
        "[30] Wang, Z., Zhao, H., Peng, J., Yao, L., and Zhao, K. (2024b). ODCR: Orthogonal Decoupling Contrastive Regularization for Unpaired Image Dehazing. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR).",
        "[31] Wu, R.-Q., Duan, Z.-P., Guo, C.-L., Chai, Z., and Li, C. (2023). RIDCP: Revitalizing Real Image Dehazing via High-Quality Codebook Priors. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 22282–22291.",
        "[32] Yang, Y., Wang, C., Liu, R., Zhang, L., Guo, X., and Tao, D. (2022). Self-Augmented Unpaired Image Dehazing via Density and Depth Decomposition. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 2037–2046.",
        "[33] Yang, Y., Li, X., Zhou, Z., Song, S. L., Wu, J., Nie, L., and Ghanem, B. (2024). CorDA: Context-Oriented Decomposition Adaptation of Large Language Models for Task-Aware Parameter-Efficient Fine-tuning. In Advances in Neural Information Processing Systems (NeurIPS).",
        "[34] Zamfir, E., Wu, Z., Mehta, N., Tan, Y., Paudel, D. P., Zhang, Y., and Timofte, R. (2025). Complexity Experts Are Task-Discriminative Learners for Any Image Restoration. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 12753–12763.",
        "[35] Zhang, W., Liu, Y., Dong, C., and Qiao, Y. (2019). RankSRGAN: Generative Adversarial Networks With Ranker for Image Super-Resolution. In Proceedings of the IEEE/CVF International Conference on Computer Vision (ICCV), pages 3096–3105.",
        "[36] Zhang, Y., Ma, L., Feng, Y., Huang, Z., Zhou, F., and Su, Z. (2026). Bilevel Layer-Positioning LoRA for Real Image Dehazing. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR).",
        "[37] Zheng, D., Wu, X., Yang, S., Zhang, J., Hu, J., and Zheng, W. (2024). Selective Hourglass Mapping for Universal Image Restoration Based on Diffusion Model. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 25445–25455.",
        "[38] Zheng, Y., Zhan, J., He, S., Dong, J., and Du, Y. (2023). Curricular Contrastive Regularization for Physics-Aware Single Image Dehazing. In Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition (CVPR), pages 5785–5794.",
    ]

    for ref_txt in references_list:
        p_ref = doc.add_paragraph(style="参考文献")
        p_ref.paragraph_format.space_before = Pt(0)
        p_ref.paragraph_format.space_after = Pt(0)
        p_ref.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_ref.paragraph_format.line_spacing = Pt(9.3)
        r = p_ref.add_run(ref_txt)
        set_run_fonts(r, cn_font="宋体", en_font="Times New Roman", size_pt=7.8)

    doc.save(str(OUT_REFINED_LOCAL))

    # Build PowerShell commands to embed true vector SVGs & update live Word TOC
    svg_ps_lines = []
    for tok, svg_p, w_cm in PENDING_SVG_INSERTS:
        svg_esc = svg_p.replace("'", "''")
        tok_esc = tok.replace("'", "''")
        svg_ps_lines.append(
            f"$rng = $d.Content; "
            f"if ($rng.Find.Execute('{tok_esc}')) {{ "
            f"$rng.Text = ''; "
            f"$shp = $rng.InlineShapes.AddPicture('{svg_esc}'); "
            f"$shp.LockAspectRatio = -1; "
            f"$shp.Width = $word.CentimetersToPoints({w_cm}); "
            f"}}"
        )
    svg_ps_block = "\n".join(svg_ps_lines)

    ps_toc_script = f'''
$p = "{str(OUT_REFINED_LOCAL)}"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$d = $word.Documents.Open($p)
{svg_ps_block}
foreach ($sName in @("TOC 1", "TOC 2", "目录 1", "目录 2")) {{
    try {{
        $st = $d.Styles.Item($sName)
        $st.Font.NameFarEast = "宋体"
        $st.Font.NameAscii = "Times New Roman"
        $st.Font.NameOther = "Times New Roman"
        $st.Font.Color = 986895
        if ($sName -match "1") {{
            $st.Font.Size = 10.5
            $st.Font.Bold = 1
            $st.ParagraphFormat.LeftIndent = 0
            $st.ParagraphFormat.SpaceBefore = 4
            $st.ParagraphFormat.SpaceAfter = 2
            $st.ParagraphFormat.LineSpacingRule = 4
            $st.ParagraphFormat.LineSpacing = 17
        }} else {{
            $st.Font.Size = 10.0
            $st.Font.Bold = 0
            $st.ParagraphFormat.LeftIndent = 18
            $st.ParagraphFormat.SpaceBefore = 1
            $st.ParagraphFormat.SpaceAfter = 1
            $st.ParagraphFormat.LineSpacingRule = 4
            $st.ParagraphFormat.LineSpacing = 15.5
        }}
        $st.ParagraphFormat.TabStops.ClearAll()
        [void]$st.ParagraphFormat.TabStops.Add($word.CentimetersToPoints(16.6), 2, 1)
    }} catch {{}}
}}
if ($d.TablesOfContents.Count -ge 1) {{
    $toc = $d.TablesOfContents.Item(1)
    $toc.UseHeadingStyles = $false
    $toc.UseOutlineLevels = $true
    $toc.UpperHeadingLevel = 1
    $toc.LowerHeadingLevel = 2
    $toc.RightAlignPageNumbers = $true
    $toc.UseHyperlinks = $true
    $toc.Update()
}}
$d.Fields.Update() | Out-Null
$d.Repaginate()
$pages = $d.ComputeStatistics(2)
Write-Host "WORD_COM_UPDATED_PAGES=$pages"
$pdfCheck = "{str((ROOT / 'preview/word_report_v5_final.pdf').resolve())}"
$d.ExportAsFixedFormat($pdfCheck, 17)
$d.Save()
$d.Close()
$word.Quit()
'''
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_toc_script],
            capture_output=True, text=True, timeout=60
        )
        print(res.stdout.strip())
        if res.returncode != 0 and res.stderr.strip():
            print("Word COM TOC update stderr:", res.stderr.strip())
    except Exception as exc:
        print("Word COM TOC update warning:", exc)

    try:
        shutil.copy2(OUT_REFINED_LOCAL, OUT_REFINED_WS)
        print("Saved refined Word report to:", OUT_REFINED_WS)
    except PermissionError:
        alt_ws = WORKSPACE / "PRISM_图像去雾大作业报告_精排终稿_新版.docx"
        shutil.copy2(OUT_REFINED_LOCAL, alt_ws)
        print("Workspace file locked by Word; saved to:", alt_ws)
    try:
        shutil.copy2(OUT_REFINED_LOCAL, OUT_FINAL_LOCAL)
        shutil.copy2(OUT_REFINED_LOCAL, OUT_FINAL_WS)
    except PermissionError:
        pass
    size_mb = OUT_REFINED_LOCAL.stat().st_size / (1024 * 1024)
    print(f"Saved refined Word report to: {OUT_REFINED_LOCAL} ({size_mb:.2f} MB)")


if __name__ == "__main__":
    build_document()
