#!/usr/bin/env python3
"""
Xcelerate Growth Partners — Proposal Package Generator
Flask web app: generates Intro Letter PDF + Proposal PDF + Proposal PPTX
from a single form submission.
"""

import os
import sys
import uuid
import subprocess
from pathlib import Path
from flask import (Flask, render_template, request, send_file,
                   jsonify)
from werkzeug.utils import secure_filename

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
SCRIPTS_DIR   = BASE_DIR / "scripts"
ASSETS_DIR    = BASE_DIR / "assets"
OUTPUTS_DIR   = BASE_DIR / "outputs"

# Bundled defaults — always available, no upload needed
DEFAULT_LOGO      = ASSETS_DIR / "xcelerate_logo.png"
DEFAULT_BASE_PDF  = ASSETS_DIR / "base_proposal_template.pdf"

ALLOWED_IMAGE_EXT = {"png", "jpg", "jpeg"}
ALLOWED_PDF_EXT   = {"pdf"}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24))
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


def allowed_file(filename, exts):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in exts


def safe_prefix(company: str) -> str:
    return "".join(
        c if c.isalnum() or c in (" ", "-", "_") else ""
        for c in company
    ).strip().replace(" ", "_")


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    # ── Collect form data ────────────────────────────────────────────────────
    company       = request.form.get("company", "").strip()
    contact       = request.form.get("contact", "").strip()
    date          = request.form.get("date", "").strip()
    services      = request.form.getlist("services")
    fee           = request.form.get("fee", "").strip()
    extra_fees    = request.form.getlist("extra_fees")
    notes         = request.form.get("notes", "").strip()
    body_override = request.form.get("body_override", "").strip()

    if not company or not date:
        return jsonify({"error": "Company name and date are required."}), 400

    # ── Build cost lines ─────────────────────────────────────────────────────
    cost_lines = []
    if fee:
        cost_lines.append(f"Program Fee: {fee}")
    for ef in extra_fees:
        ef = ef.strip()
        if ef:
            cost_lines.append(ef)

    # ── Set up per-job output directory ──────────────────────────────────────
    job_id  = uuid.uuid4().hex[:8]
    job_dir = OUTPUTS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    # ── Resolve logo: uploaded > bundled default ──────────────────────────────
    logo_path = str(DEFAULT_LOGO)
    if "logo" in request.files and request.files["logo"].filename:
        f = request.files["logo"]
        if allowed_file(f.filename, ALLOWED_IMAGE_EXT):
            dest = job_dir / "logo.png"
            f.save(dest)
            logo_path = str(dest)

    # ── Resolve base PDF: uploaded > bundled default ──────────────────────────
    base_pdf_path = str(DEFAULT_BASE_PDF) if DEFAULT_BASE_PDF.exists() else ""
    if "base_pdf" in request.files and request.files["base_pdf"].filename:
        f = request.files["base_pdf"]
        if allowed_file(f.filename, ALLOWED_PDF_EXT):
            dest = job_dir / "base_template.pdf"
            f.save(dest)
            base_pdf_path = str(dest)

    # ── Output paths ─────────────────────────────────────────────────────────
    prefix        = safe_prefix(company)
    letter_pdf    = str(job_dir / f"{prefix}_Intro_Letter.pdf")
    proposal_pdf  = str(job_dir / f"{prefix}_Proposal.pdf")
    proposal_pptx = str(job_dir / f"{prefix}_Proposal.pptx")

    env = {**os.environ, "XCELERATE_LOGO": logo_path}
    errors = []

    # ── Generate Intro Letter PDF ─────────────────────────────────────────────
    letter_cmd = [
        sys.executable, str(SCRIPTS_DIR / "generate_letter.py"),
        "--company",  company,
        "--date",     date,
        "--services", ", ".join(services) if services else "our full suite of services",
        "--output",   letter_pdf,
    ]
    if contact:       letter_cmd += ["--contact", contact]
    if body_override: letter_cmd += ["--body", body_override]

    r = subprocess.run(letter_cmd, capture_output=True, text=True, env=env)
    if r.returncode != 0:
        errors.append(f"Letter: {r.stderr.strip()}")

    # ── Generate Proposal PDF + PPTX ─────────────────────────────────────────
    proposal_cmd = [
        sys.executable, str(SCRIPTS_DIR / "generate_proposal.py"),
        "--company", company,
        "--date",    date,
        "--output",  proposal_pdf,
    ]
    if contact:        proposal_cmd += ["--contact",  contact]
    if base_pdf_path:  proposal_cmd += ["--base-pdf", base_pdf_path]
    if services:       proposal_cmd += ["--services", "|".join(services)]
    if cost_lines:     proposal_cmd += ["--costs",    "|".join(cost_lines)]
    if notes:          proposal_cmd += ["--notes",    notes]

    r2 = subprocess.run(proposal_cmd, capture_output=True, text=True, env=env)
    if r2.returncode != 0:
        errors.append(f"Proposal: {r2.stderr.strip()}")

    # ── Check outputs ─────────────────────────────────────────────────────────
    produced = {
        "letter_pdf":    os.path.exists(letter_pdf),
        "proposal_pdf":  os.path.exists(proposal_pdf),
        "proposal_pptx": os.path.exists(proposal_pptx),
    }

    if not any(produced.values()):
        return jsonify({"error": "Generation failed. " + " | ".join(errors)}), 500

    return jsonify({
        "job_id":             job_id,
        "company":            company,
        "date":               date,
        "produced":           produced,
        "errors":             errors,
        "letter_name":        f"{prefix}_Intro_Letter.pdf",
        "proposal_pdf_name":  f"{prefix}_Proposal.pdf",
        "proposal_pptx_name": f"{prefix}_Proposal.pptx",
    })


@app.route("/download/<job_id>/<filename>")
def download(job_id, filename):
    safe_job  = secure_filename(job_id)
    safe_name = secure_filename(filename)
    file_path = OUTPUTS_DIR / safe_job / safe_name
    if not file_path.exists():
        return "File not found", 404
    return send_file(str(file_path), as_attachment=True, download_name=safe_name)


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    OUTPUTS_DIR.mkdir(exist_ok=True)
    port = int(os.environ.get("PORT", 5050))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print(f"\n✅  Xcelerate Proposal Generator running at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=debug)
