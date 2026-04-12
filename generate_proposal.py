#!/usr/bin/env python3
"""
Xcelerate Growth Partners - Proposal Deck Generator

Produces TWO files for every proposal:
  1. A PDF  (same filename, .pdf extension)
  2. A PPTX (same filename, .pptx extension)

Usage:
    python generate_proposal.py \
        --company   "Thornburg Investments" \
        --contact   "Sarah Chen" \
        --date      "September 15, 2025" \
        --base-pdf  "/path/to/base_proposal_template.pdf" \
        --services  "Team Development and Optimization|Succession Planning|Keynote Speeches" \
        --costs     "Program Fee: $25,000|Travel & Expenses: At cost" \
        --notes     "Optional custom paragraph for the services slide." \
        --output    "/path/to/output/Thornburg_Proposal.pdf"

Services and costs use | as a separator between line items.
The .pptx file is saved automatically next to the .pdf output.
"""

import argparse
import os
import io

# ── reportlab (PDF) ───────────────────────────────────────────────────────────
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, white
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from pypdf import PdfReader, PdfWriter

# ── python-pptx (PPTX) ───────────────────────────────────────────────────────
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ── Shared brand constants ────────────────────────────────────────────────────
# ReportLab colours
RL_NAVY      = HexColor("#1D3461")
RL_GREEN     = HexColor("#5CB85C")
RL_GREEN_DK  = HexColor("#3A7A3A")
RL_WHITE     = white
RL_BODY      = HexColor("#2C2C2C")
RL_LGRAY     = HexColor("#F5F5F5")
RL_MGRAY     = HexColor("#CCCCCC")

# python-pptx colours
PT_NAVY      = RGBColor(0x1D, 0x34, 0x61)
PT_GREEN     = RGBColor(0x5C, 0xB8, 0x5C)
PT_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
PT_BODY      = RGBColor(0x2C, 0x2C, 0x2C)
PT_LGRAY     = RGBColor(0xF5, 0xF5, 0xF5)
PT_MGRAY     = RGBColor(0xCC, 0xCC, 0xCC)

# PDF page size — matches the base Edward Jones proposal PDF exactly
PDF_W, PDF_H = 720.0, 405.0

# PPTX slide size — standard 16:9 widescreen (13.33" x 7.5")
PPTX_W = Inches(13.33)
PPTX_H = Inches(7.5)


# ═══════════════════════════════════════════════════════════════════════════════
#  Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _logo_path(script_dir: str) -> str:
    return os.environ.get(
        "XCELERATE_LOGO",
        os.path.join(script_dir, "..", "assets", "xcelerate_logo.png")
    )


def _pptx_output_path(pdf_output: str) -> str:
    """Derive the .pptx path from the .pdf output path."""
    base = os.path.splitext(pdf_output)[0]
    return base + ".pptx"


# ── Boilerplate slide content (mirrors the Edward Jones proposal template) ────
STANDARD_SLIDES = [
    {
        "title": "Who We Are",
        "bullets": [
            "We are a team with extensive knowledge in Wealth Management. Our expertise is broad and diverse, with a unique understanding of clients, advisors and investment products.",
            "We are keenly focused on growth and business development. Each engagement is custom, and we partner with clients to drive specific desired outcomes.",
            "Our capabilities are unique and differentiated. We merge the talents of our core team with a roster of strategic alliances to ensure best-in-class support and project execution.",
            "Our process and strategy is to excel at leading edge thinking, converting fresh new ideas into measurable results and positive transformation.",
        ],
    },
    {
        "title": "Our Competitive Advantage",
        "bullets": [
            "Team Development and Optimization",
            "Succession Planning and Practice Transitions",
            "The Future of Wealth Management",
            "Developing Leadership",
            "Support Staff Development and Career Pathing",
            "The Wealth Management Process",
            "  – Client Management",
            "  – Financial Planning",
            "  – The Investment Process",
            "  – Exceeding Client Expectations",
        ],
    },
    {
        "title": "Why Xcelerate?",
        "subtitle": "Asset Manager Opportunities",
        "bullets": [
            "Over 150 years of significant Wealth Management experience.",
            "Team's track record of success and industry credibility.",
            "Access to cutting edge and relevant content.",
            "Firsthand knowledge of Financial Advisor and client preferences.",
            "Access to talented presenters and thought-provoking content development specialists.",
            "Compelling conversations about The Future of Wealth Management.",
            "Creative strategies to improve engagement and outcomes.",
        ],
    },
    {
        "title": "Specific Partnership Ideas",
        "bullets": [
            "Working with Wealth Management firms on strategy and growth around specific themes such as retirement planning, succession planning or acting on key business initiatives.",
            "Content development and value-added white papers.",
            "Interactive and innovative idea development.",
            "Cutting edge Practice Management training by talented industry professionals.",
            "Keynote speeches.",
            "Conference and sponsorship strategy to enhance relationships and outcomes.",
            "Action plan to align annual spend and year-long marketing planning.",
            "Advisor relationship development.",
            '"Train the trainer" workshops to improve wholesaler presentations.',
            "Access to best practice thinking, curated articles and creative training initiatives.",
            "Production of webinars, podcasts, and access to sought-after programs.",
        ],
    },
    {
        "title": "Our Team of Experts",
        "team": [
            ("Jim Tracy", "Co-Founder & CEO",
             "As a 40-year veteran of the wealth management industry, Jim is a recognized leader. He has achieved considerable success in growing and building innovative solutions that benefit advisors and their clients."),
            ("Mary Deatherage", "Co-Founder & President",
             "Mary is a respected industry leader, who previously led a multi-billion dollar team at Morgan Stanley. Her strategic vision and expertise are invaluable assets to the firm."),
            ("Tara Forrest", "Head of Business Management",
             "Tara brings a wealth of expertise in developing and implementing effective business strategies. Her background in operations and risk management add structure and additional capabilities to the group."),
        ],
    },
]


# ═══════════════════════════════════════════════════════════════════════════════
#  PDF GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

def _pdf_footer(c, page_num):
    c.setFillColor(RL_NAVY)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, 18, "Xcelerate Growth Partners")
    c.setFillColor(RL_GREEN)
    c.rect(PDF_W - 50, 8, 36, 24, fill=1, stroke=0)
    c.setFillColor(RL_WHITE)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(PDF_W - 32, 16, str(page_num))


def _pdf_cover(c, company, contact, date, logo_path):
    c.setFillColor(RL_NAVY)
    c.rect(0, 0, PDF_W, PDF_H, fill=1, stroke=0)
    c.setFillColor(RL_GREEN)
    c.rect(0, PDF_H - 140, PDF_W, 6, fill=1, stroke=0)
    if logo_path and os.path.exists(logo_path):
        c.drawImage(ImageReader(logo_path),
                    PDF_W / 2 - 100, PDF_H - 130, 200, 90,
                    mask='auto', preserveAspectRatio=True, anchor='c')
    c.setFillColor(RL_WHITE)
    c.setFont("Helvetica-Bold", 42)
    c.drawCentredString(PDF_W / 2, PDF_H / 2 + 80, contact or company)
    if contact:
        c.drawCentredString(PDF_W / 2, PDF_H / 2 + 20, company)
    c.setFillColor(RL_GREEN)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(PDF_W / 2, PDF_H / 2 - 50, date)
    c.setFillColor(RL_WHITE)
    c.setFont("Helvetica", 9)
    c.drawString(PDF_W - 200, 100, "Presented By:")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(PDF_W - 200, 85,  "Jim Tracy, Co-Founder & CEO")
    c.drawString(PDF_W - 200, 70,  "Mary Deatherage, Co-Founder & President")
    c.setFont("Helvetica", 10)
    c.drawString(PDF_W - 200, 55,  "Xcelerate Growth Partners")
    c.setFillColor(RL_GREEN)
    c.rect(0, 0, PDF_W, 8, fill=1, stroke=0)
    c.showPage()


def _pdf_services(c, company, services, notes, page_num):
    c.setFillColor(RL_WHITE)
    c.rect(0, 0, PDF_W, PDF_H, fill=1, stroke=0)
    c.setFillColor(RL_NAVY)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(36, PDF_H - 44, "Recommended Services")
    c.setFillColor(RL_BODY)
    c.setFont("Helvetica", 10)
    c.drawString(36, PDF_H - 60, f"Tailored engagement plan for {company}")
    c.setFillColor(RL_GREEN)
    c.rect(36, PDF_H - 68, PDF_W - 72, 2, fill=1, stroke=0)
    bullet_style = ParagraphStyle('b', fontName='Helvetica-Bold', fontSize=10,
                                  leading=14, textColor=RL_NAVY, leftIndent=12)
    mid = PDF_W / 2
    col_w = mid - 60
    y_l = y_r = PDF_H - 84
    for i, svc in enumerate(services):
        p = Paragraph(f"• {svc}", bullet_style)
        pw, ph = p.wrap(col_w, PDF_H)
        if i < (len(services) + 1) // 2:
            p.drawOn(c, 36, y_l - ph);  y_l -= ph + 6
        else:
            p.drawOn(c, mid + 10, y_r - ph);  y_r -= ph + 6
    if notes:
        ns = ParagraphStyle('n', fontName='Helvetica-Oblique', fontSize=9,
                            leading=13, textColor=RL_BODY)
        p = Paragraph(notes, ns)
        pw, ph = p.wrap(PDF_W - 72, PDF_H)
        p.drawOn(c, 36, min(y_l, y_r) - 10 - ph)
    _pdf_footer(c, page_num)
    c.showPage()


def _pdf_investment(c, company, costs, page_num):
    c.setFillColor(RL_WHITE)
    c.rect(0, 0, PDF_W, PDF_H, fill=1, stroke=0)
    c.setFillColor(RL_NAVY)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(36, PDF_H - 44, "Proposed Investment")
    c.setFillColor(RL_BODY)
    c.setFont("Helvetica", 10)
    c.drawString(36, PDF_H - 60, f"Fee structure for {company}")
    c.setFillColor(RL_GREEN)
    c.rect(36, PDF_H - 68, PDF_W - 72, 2, fill=1, stroke=0)
    row_h, row_y = 30, PDF_H - 84
    # Separate total line from itemized lines
    item_costs  = [l for l in costs if not l.lower().startswith("total")]
    total_lines = [l for l in costs if l.lower().startswith("total")]
    for i, line in enumerate(item_costs):
        label, amount = (line.split(":", 1) if ":" in line else (line, ""))
        if i % 2 == 0:
            c.setFillColor(RL_LGRAY)
            c.rect(36, row_y - row_h + 8, PDF_W - 72, row_h, fill=1, stroke=0)
        c.setFillColor(RL_NAVY);  c.setFont("Helvetica-Bold", 11)
        c.drawString(48, row_y - 6, label.strip())
        c.setFillColor(RL_GREEN_DK); c.setFont("Helvetica-Bold", 11)
        c.drawRightString(PDF_W - 48, row_y - 6, amount.strip())
        row_y -= row_h
    # Total row — navy background, larger text
    if total_lines:
        label, amount = (total_lines[0].split(":", 1) if ":" in total_lines[0] else (total_lines[0], ""))
        row_y -= 6
        c.setFillColor(RL_NAVY)
        c.rect(36, row_y - row_h + 8, PDF_W - 72, row_h + 4, fill=1, stroke=0)
        c.setFillColor(RL_WHITE); c.setFont("Helvetica-Bold", 13)
        c.drawString(48, row_y - 5, label.strip())
        c.setFillColor(RL_GREEN); c.setFont("Helvetica-Bold", 13)
        c.drawRightString(PDF_W - 48, row_y - 5, amount.strip())
    c.setFillColor(RL_MGRAY);  c.setFont("Helvetica-Oblique", 8)
    c.drawString(36, 48, "All fees are subject to final scope confirmation. "
                 "Travel expenses billed at cost.")
    _pdf_footer(c, page_num)
    c.showPage()


def build_proposal_pdf(company, contact, date, base_pdf_path,
                       services, costs, notes, output_path, logo_path):
    writer = PdfWriter()

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PDF_W, PDF_H))
    _pdf_cover(c, company, contact, date, logo_path)
    c.save();  buf.seek(0)
    writer.add_page(PdfReader(buf).pages[0])

    base_count = 0
    if base_pdf_path and os.path.exists(base_pdf_path):
        base_reader = PdfReader(base_pdf_path)
        for page in base_reader.pages[1:]:
            writer.add_page(page)
        base_count = len(base_reader.pages) - 1
    else:
        print("Warning: base PDF not found — skipping template pages")

    if services:
        buf2 = io.BytesIO()
        c2 = canvas.Canvas(buf2, pagesize=(PDF_W, PDF_H))
        _pdf_services(c2, company, services, notes, base_count + 2)
        c2.save();  buf2.seek(0)
        writer.add_page(PdfReader(buf2).pages[0])

    if costs:
        buf3 = io.BytesIO()
        c3 = canvas.Canvas(buf3, pagesize=(PDF_W, PDF_H))
        _pdf_investment(c3, company, costs, base_count + 3)
        c3.save();  buf3.seek(0)
        writer.add_page(PdfReader(buf3).pages[0])

    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"PDF  saved → {output_path}")


# ═══════════════════════════════════════════════════════════════════════════════
#  PPTX GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

def _pptx_blank_slide(prs):
    """Add a completely blank slide (no placeholders)."""
    blank_layout = prs.slide_layouts[6]   # index 6 = Blank
    return prs.slides.add_slide(blank_layout)


def _pptx_bg(slide, color: RGBColor):
    """Fill slide background with a solid colour."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _pptx_rect(slide, left, top, width, height, fill_color: RGBColor,
               line_color=None):
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def _pptx_textbox(slide, left, top, width, height, text, font_size,
                  bold=False, color: RGBColor = None, align=PP_ALIGN.LEFT,
                  italic=False, wrap=True):
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    return txb


def _pptx_footer(slide, page_num):
    W, H = PPTX_W, PPTX_H
    # "Xcelerate Growth Partners" text
    _pptx_textbox(slide, Inches(0.4), H - Inches(0.45),
                  Inches(4), Inches(0.35),
                  "Xcelerate Growth Partners", 9,
                  bold=True, color=PT_NAVY)
    # Green page-number box
    box = _pptx_rect(slide,
                     W - Inches(0.6), H - Inches(0.45),
                     Inches(0.45), Inches(0.35), PT_GREEN)
    _pptx_textbox(slide, W - Inches(0.6), H - Inches(0.45),
                  Inches(0.45), Inches(0.35),
                  str(page_num), 11,
                  bold=True, color=PT_WHITE, align=PP_ALIGN.CENTER)


def _pptx_cover(prs, company, contact, date, logo_path):
    slide = _pptx_blank_slide(prs)
    W, H = PPTX_W, PPTX_H
    _pptx_bg(slide, PT_NAVY)

    # Green stripe across top
    _pptx_rect(slide, 0, Inches(1.6), W, Inches(0.07), PT_GREEN)

    # Logo
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path,
                                 W / 2 - Inches(1.5), Inches(0.15),
                                 Inches(3.0), Inches(1.35))

    # Contact name / company name
    name_top = H / 2 - Inches(1.0)
    _pptx_textbox(slide, Inches(1), name_top,
                  W - Inches(2), Inches(0.8),
                  contact or company, 40,
                  bold=True, color=PT_WHITE, align=PP_ALIGN.CENTER)
    if contact:
        _pptx_textbox(slide, Inches(1), name_top + Inches(0.85),
                      W - Inches(2), Inches(0.8),
                      company, 40,
                      bold=True, color=PT_WHITE, align=PP_ALIGN.CENTER)

    # Date (green)
    _pptx_textbox(slide,
                  Inches(1), H / 2 + Inches(0.6),
                  W - Inches(2), Inches(0.5),
                  date, 22,
                  bold=True, color=PT_GREEN, align=PP_ALIGN.CENTER)

    # Presented By block (bottom-right)
    by_left = W - Inches(3.2)
    by_top  = H - Inches(1.6)
    _pptx_textbox(slide, by_left, by_top,
                  Inches(3.0), Inches(0.25),
                  "Presented By:", 9, color=PT_WHITE)
    lines = [
        ("Jim Tracy, Co-Founder & CEO",          True),
        ("Mary Deatherage, Co-Founder & President", True),
        ("Xcelerate Growth Partners",             False),
    ]
    y = by_top + Inches(0.28)
    for text, bold in lines:
        _pptx_textbox(slide, by_left, y, Inches(3.0), Inches(0.25),
                      text, 10, bold=bold, color=PT_WHITE)
        y += Inches(0.26)

    # Bottom green bar
    _pptx_rect(slide, 0, H - Inches(0.1), W, Inches(0.1), PT_GREEN)


def _pptx_content_slide(prs, title, bullets, subtitle=None,
                        page_num=None, two_col=False):
    slide = _pptx_blank_slide(prs)
    W, H = PPTX_W, PPTX_H
    _pptx_bg(slide, PT_WHITE)

    # Title
    _pptx_textbox(slide, Inches(0.5), Inches(0.3),
                  W - Inches(1), Inches(0.65),
                  title, 24, bold=True, color=PT_NAVY)

    # Optional subtitle
    y_start = Inches(1.05)
    if subtitle:
        _pptx_textbox(slide, Inches(0.5), Inches(0.95),
                      W - Inches(1), Inches(0.35),
                      subtitle, 12, color=PT_BODY)
        y_start = Inches(1.35)

    # Green divider line
    _pptx_rect(slide, Inches(0.5), y_start - Inches(0.05),
               W - Inches(1), Inches(0.04), PT_GREEN)

    # Bullets (single or two-column)
    bullet_font = 11
    line_h = Inches(0.38)
    content_h = H - y_start - Inches(0.6)

    if two_col:
        half = len(bullets) // 2 + len(bullets) % 2
        col_w = (W - Inches(1.2)) / 2
        for col_idx, col_bullets in enumerate([bullets[:half], bullets[half:]]):
            x = Inches(0.5) + col_idx * (col_w + Inches(0.2))
            y = y_start + Inches(0.12)
            for b in col_bullets:
                _pptx_textbox(slide, x, y, col_w, line_h,
                              f"• {b}", bullet_font,
                              bold=True, color=PT_NAVY)
                y += line_h
    else:
        y = y_start + Inches(0.12)
        for b in bullets:
            indent = b.startswith("  –")
            _pptx_textbox(slide,
                          Inches(0.5) + (Inches(0.3) if indent else 0),
                          y,
                          W - Inches(1.0 if not indent else 1.3),
                          line_h,
                          ("  " if indent else "• ") + b.lstrip(),
                          bullet_font,
                          bold=not indent, color=PT_NAVY if not indent else PT_BODY)
            y += line_h

    if page_num:
        _pptx_footer(slide, page_num)


def _pptx_team_slide(prs, team_members, page_num):
    slide = _pptx_blank_slide(prs)
    W, H = PPTX_W, PPTX_H
    _pptx_bg(slide, PT_WHITE)

    _pptx_textbox(slide, Inches(0.5), Inches(0.3),
                  W - Inches(1), Inches(0.65),
                  "Our Team of Experts", 24, bold=True, color=PT_NAVY)
    _pptx_rect(slide, Inches(0.5), Inches(1.0),
               W - Inches(1), Inches(0.04), PT_GREEN)

    row_h = (H - Inches(1.5)) / len(team_members)
    for i, (name, title, bio) in enumerate(team_members):
        y = Inches(1.1) + i * row_h
        # Name + title
        _pptx_textbox(slide, Inches(0.5), y,
                      Inches(4), Inches(0.35),
                      name, 13, bold=True, color=PT_NAVY)
        _pptx_textbox(slide, Inches(0.5), y + Inches(0.35),
                      Inches(4), Inches(0.3),
                      title, 10, bold=False, color=PT_GREEN)
        # Bio
        _pptx_textbox(slide, Inches(4.8), y,
                      W - Inches(5.3), row_h - Inches(0.1),
                      bio, 10, color=PT_BODY, wrap=True)

    _pptx_footer(slide, page_num)


def _pptx_services_slide(prs, company, services, notes, page_num):
    slide = _pptx_blank_slide(prs)
    W, H = PPTX_W, PPTX_H
    _pptx_bg(slide, PT_WHITE)

    _pptx_textbox(slide, Inches(0.5), Inches(0.3),
                  W - Inches(1), Inches(0.65),
                  "Recommended Services", 24, bold=True, color=PT_NAVY)
    _pptx_textbox(slide, Inches(0.5), Inches(0.92),
                  W - Inches(1), Inches(0.35),
                  f"Tailored engagement plan for {company}", 11, color=PT_BODY)
    _pptx_rect(slide, Inches(0.5), Inches(1.28),
               W - Inches(1), Inches(0.04), PT_GREEN)

    half = (len(services) + 1) // 2
    col_w = (W - Inches(1.2)) / 2
    line_h = Inches(0.42)
    for col_idx, col_svcs in enumerate([services[:half], services[half:]]):
        x = Inches(0.5) + col_idx * (col_w + Inches(0.2))
        y = Inches(1.4)
        for svc in col_svcs:
            _pptx_textbox(slide, x, y, col_w, line_h,
                          f"• {svc}", 11, bold=True, color=PT_NAVY)
            y += line_h

    if notes:
        note_y = Inches(1.4) + half * line_h + Inches(0.1)
        _pptx_textbox(slide, Inches(0.5), note_y,
                      W - Inches(1), Inches(0.5),
                      notes, 10, italic=True, color=PT_BODY, wrap=True)

    _pptx_footer(slide, page_num)


def _pptx_investment_slide(prs, company, costs, page_num):
    slide = _pptx_blank_slide(prs)
    W, H = PPTX_W, PPTX_H
    _pptx_bg(slide, PT_WHITE)

    _pptx_textbox(slide, Inches(0.5), Inches(0.3),
                  W - Inches(1), Inches(0.65),
                  "Proposed Investment", 24, bold=True, color=PT_NAVY)
    _pptx_textbox(slide, Inches(0.5), Inches(0.92),
                  W - Inches(1), Inches(0.35),
                  f"Fee structure for {company}", 11, color=PT_BODY)
    _pptx_rect(slide, Inches(0.5), Inches(1.28),
               W - Inches(1), Inches(0.04), PT_GREEN)

    row_h = Inches(0.55)
    y = Inches(1.4)
    item_costs  = [l for l in costs if not l.lower().startswith("total")]
    total_lines = [l for l in costs if l.lower().startswith("total")]
    for i, line in enumerate(item_costs):
        label, amount = (line.split(":", 1) if ":" in line else (line, ""))
        if i % 2 == 0:
            _pptx_rect(slide, Inches(0.5), y,
                       W - Inches(1), row_h, PT_LGRAY)
        _pptx_textbox(slide, Inches(0.65), y + Inches(0.1),
                      Inches(6), row_h - Inches(0.1),
                      label.strip(), 13, bold=True, color=PT_NAVY)
        _pptx_textbox(slide, W - Inches(3.5), y + Inches(0.1),
                      Inches(3.0), row_h - Inches(0.1),
                      amount.strip(), 13, bold=True,
                      color=RGBColor(0x3A, 0x7A, 0x3A),
                      align=PP_ALIGN.RIGHT)
        y += row_h
    # Total row — navy background
    if total_lines:
        label, amount = (total_lines[0].split(":", 1) if ":" in total_lines[0] else (total_lines[0], ""))
        y += Inches(0.05)
        _pptx_rect(slide, Inches(0.5), y, W - Inches(1), row_h + Inches(0.05), PT_NAVY)
        _pptx_textbox(slide, Inches(0.65), y + Inches(0.1),
                      Inches(6), row_h,
                      label.strip(), 14, bold=True, color=PT_WHITE)
        _pptx_textbox(slide, W - Inches(3.5), y + Inches(0.1),
                      Inches(3.0), row_h,
                      amount.strip(), 14, bold=True,
                      color=PT_GREEN, align=PP_ALIGN.RIGHT)
        y += row_h

    _pptx_textbox(slide, Inches(0.5), H - Inches(0.75),
                  W - Inches(1), Inches(0.3),
                  "All fees are subject to final scope confirmation. "
                  "Travel expenses billed at cost.",
                  8, italic=True, color=PT_MGRAY)

    _pptx_footer(slide, page_num)


def build_proposal_pptx(company, contact, date,
                        services, costs, notes,
                        output_path, logo_path):
    prs = Presentation()
    prs.slide_width  = PPTX_W
    prs.slide_height = PPTX_H

    # 1. Cover
    _pptx_cover(prs, company, contact, date, logo_path)

    # 2. Standard content slides
    page = 2
    for slide_def in STANDARD_SLIDES:
        if "team" in slide_def:
            _pptx_team_slide(prs, slide_def["team"], page)
        else:
            two_col = len(slide_def.get("bullets", [])) > 6
            _pptx_content_slide(
                prs,
                title    = slide_def["title"],
                bullets  = slide_def.get("bullets", []),
                subtitle = slide_def.get("subtitle"),
                page_num = page,
                two_col  = two_col,
            )
        page += 1

    # 3. Custom services slide
    if services:
        _pptx_services_slide(prs, company, services, notes, page)
        page += 1

    # 4. Investment slide
    if costs:
        _pptx_investment_slide(prs, company, costs, page)

    prs.save(output_path)
    print(f"PPTX saved → {output_path}")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Generate Xcelerate proposal — outputs both PDF and PPTX")
    parser.add_argument("--company",  default="[COMPANY NAME]")
    parser.add_argument("--contact",  default="")
    parser.add_argument("--date",     default="[DATE]")
    parser.add_argument("--base-pdf", default="",
                        help="Path to the base proposal template PDF")
    parser.add_argument("--services", default="",
                        help="Pipe-separated services: 'Service A|Service B'")
    parser.add_argument("--costs",    default="",
                        help="Pipe-separated cost lines: 'Label: $amount|...'")
    parser.add_argument("--notes",    default="",
                        help="Optional custom paragraph for the services slide")
    parser.add_argument("--output",   default="Xcelerate_Proposal.pdf",
                        help="Output PDF path (.pptx is saved alongside automatically)")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    logo       = _logo_path(script_dir)

    services = [s.strip() for s in args.services.split("|") if s.strip()] if args.services else []
    costs    = [c.strip() for c in args.costs.split("|")    if c.strip()] if args.costs    else []

    # ── PDF ──────────────────────────────────────────────────────────────────
    build_proposal_pdf(
        company       = args.company,
        contact       = args.contact,
        date          = args.date,
        base_pdf_path = args.base_pdf,
        services      = services,
        costs         = costs,
        notes         = args.notes,
        output_path   = args.output,
        logo_path     = logo,
    )

    # ── PPTX ─────────────────────────────────────────────────────────────────
    pptx_out = _pptx_output_path(args.output)
    build_proposal_pptx(
        company     = args.company,
        contact     = args.contact,
        date        = args.date,
        services    = services,
        costs       = costs,
        notes       = args.notes,
        output_path = pptx_out,
        logo_path   = logo,
    )


if __name__ == "__main__":
    main()
