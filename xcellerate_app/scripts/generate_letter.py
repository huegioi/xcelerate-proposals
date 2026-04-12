#!/usr/bin/env python3
"""
Xcelerate Growth Partners - Intro Letter PDF Generator

Usage:
    python generate_letter.py \
        --company "ACME CORP" \
        --contact "Jane Smith" \
        --date "September 15, 2025" \
        --services "practice management training, advisor development, keynote speakers" \
        --body "Optional custom body paragraph to replace the default text." \
        --output "/path/to/output/ACME_Corp_Intro_Letter.pdf"

All arguments except --output are optional; defaults are shown in the script.
"""

import argparse
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, white
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ── Brand Colours ─────────────────────────────────────────────────────────────
NAVY   = HexColor("#1D3461")   # dark navy header/footer background
GREEN  = HexColor("#5CB85C")   # thin green accent bar
WHITE  = white
BODY_DARK = HexColor("#2C2C2C")  # body text

# ── Layout Constants (points, 1pt = 1/72 inch) ────────────────────────────────
PAGE_W, PAGE_H = letter  # 612 x 792

HEADER_H     = 110        # height of dark navy top band
FOOTER_H     = 50         # height of navy footer
GREEN_BAR_H  = 6          # thin green stripe beneath header

LOGO_X       = PAGE_W / 2 - 90   # centred logo
LOGO_Y       = PAGE_H - HEADER_H + 15
LOGO_W       = 180
LOGO_H       = 80

CONTENT_TOP  = PAGE_H - HEADER_H - GREEN_BAR_H - 60  # start of white area
CONTENT_MID  = PAGE_W / 2


def build_letter(company: str, contact: str, date: str,
                 services: str, body_override: str, output_path: str,
                 logo_path: str):
    c = canvas.Canvas(output_path, pagesize=letter)

    # ── Header band ──────────────────────────────────────────────────────────
    c.setFillColor(NAVY)
    c.rect(0, PAGE_H - HEADER_H, PAGE_W, HEADER_H, fill=1, stroke=0)

    # Logo
    if logo_path and os.path.exists(logo_path):
        c.drawImage(ImageReader(logo_path),
                    LOGO_X, LOGO_Y, LOGO_W, LOGO_H,
                    mask='auto', preserveAspectRatio=True, anchor='c')

    # ── Thin green accent bar ─────────────────────────────────────────────────
    c.setFillColor(GREEN)
    c.rect(0, PAGE_H - HEADER_H - GREEN_BAR_H, PAGE_W, GREEN_BAR_H,
           fill=1, stroke=0)

    # ── Company name (large centred heading) ──────────────────────────────────
    company_y = CONTENT_TOP - 10
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 36)
    c.drawCentredString(CONTENT_MID, company_y, company.upper())

    # Contact name (if provided)
    if contact:
        c.setFont("Helvetica", 16)
        c.setFillColor(BODY_DARK)
        c.drawCentredString(CONTENT_MID, company_y - 30, contact)

    # Delivery Launch Date
    date_label_y = company_y - (55 if contact else 38)
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(NAVY)
    c.drawCentredString(CONTENT_MID, date_label_y, "Delivery Launch Date:")
    c.setFont("Helvetica", 11)
    c.setFillColor(GREEN)
    c.drawCentredString(CONTENT_MID, date_label_y - 16, date)

    # ── Body paragraphs ───────────────────────────────────────────────────────
    # Build default body text using the variable fields
    if body_override:
        paragraphs = [body_override]
    else:
        paragraphs = [
            f"We're looking forward to learning from each other and how we can "
            f"best meet your needs when we connect on {date}.",

            f"As we developed our training themes for {company}, we knew there "
            f"could be many ways to deliver our ideas — in person, via digital "
            f"learning, customized to your specific goals. Areas we're excited "
            f"to explore with you include: {services}.",

            "This is not a CPE program! Throughout, our goal has been to focus "
            "on open-ended issues that real advisory teams face each day. We "
            "offer provocative questions with no canned solutions, but with a "
            "lot of thoughtful possible answers, most offered by the participants.",

            "Let's see how we can work together!",
        ]

    style = ParagraphStyle(
        name='body',
        fontName='Helvetica',
        fontSize=12,
        leading=18,
        textColor=BODY_DARK,
        alignment=TA_LEFT,
        spaceAfter=16,
    )

    body_top = company_y - (110 if contact else 90)
    for para_text in paragraphs:
        para = Paragraph(para_text, style)
        w, h = para.wrap(PAGE_W - 120, PAGE_H)
        para.drawOn(c, 60, body_top - h)
        body_top -= h + 20

    # ── Footer band ───────────────────────────────────────────────────────────
    c.setFillColor(NAVY)
    c.rect(0, 0, PAGE_W, FOOTER_H, fill=1, stroke=0)

    # Thin green bar above footer
    c.setFillColor(GREEN)
    c.rect(0, FOOTER_H, PAGE_W, GREEN_BAR_H, fill=1, stroke=0)

    # Website text in footer
    c.setFillColor(WHITE)
    c.setFont("Helvetica", 10)
    c.drawCentredString(CONTENT_MID, FOOTER_H / 2 - 4,
                        "www.xcelerategrowthpartners.com")

    c.save()
    print(f"Letter saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Generate Xcelerate intro letter PDF")
    parser.add_argument("--company",  default="[COMPANY NAME]")
    parser.add_argument("--contact",  default="")
    parser.add_argument("--date",     default="[DATE]")
    parser.add_argument("--services", default="practice management, advisor development, and keynote presentations")
    parser.add_argument("--body",     default="",
                        help="Full custom body text (replaces all default paragraphs)")
    parser.add_argument("--output",   default="Xcelerate_Intro_Letter.pdf")
    args = parser.parse_args()

    # Logo: env var override (from web app), else default assets folder
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logo_path  = os.environ.get(
        "XCELERATE_LOGO",
        os.path.join(script_dir, "..", "assets", "xcelerate_logo.png")
    )

    build_letter(
        company       = args.company,
        contact       = args.contact,
        date          = args.date,
        services      = args.services,
        body_override = args.body,
        output_path   = args.output,
        logo_path     = logo_path,
    )


if __name__ == "__main__":
    main()
