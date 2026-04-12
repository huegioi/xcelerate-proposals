#!/usr/bin/env python3
"""
Xcelerate Growth Partners — Proposal Package Generator
"""

import os, sys, uuid, json, subprocess
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

try:
    import anthropic as _anthropic
    _ANTHROPIC_AVAILABLE = True
except ImportError:
    _ANTHROPIC_AVAILABLE = False

# ── Services catalog (must match index.html) ──────────────────────────────────
SERVICES_CATALOG = [
    "Team Development and Optimization",
    "Succession Planning and Practice Transitions",
    "The Future of Wealth Management (keynotes / workshops)",
    "Developing Leadership",
    "Support Staff Development and Career Pathing",
    "The Wealth Management Process",
    "Content Development and White Papers",
    "Conference and Sponsorship Strategy",
    '"Train the Trainer" Workshops',
    "Webinars and Podcasts",
    "Advisor Relationship Development",
]

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
SCRIPTS_DIR   = BASE_DIR / "scripts"
ASSETS_DIR    = BASE_DIR / "assets"
OUTPUTS_DIR   = BASE_DIR / "outputs"
DATA_DIR      = BASE_DIR / "data"
PROPOSALS_FILE = DATA_DIR / "proposals.json"

DEFAULT_LOGO      = ASSETS_DIR / "xcelerate_logo.png"
DEFAULT_BASE_PDF  = ASSETS_DIR / "base_proposal_template.pdf"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24))
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


def allowed_file(filename, exts):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in exts

def safe_prefix(company):
    return "".join(c if c.isalnum() or c in (" ","-","_") else "" for c in company).strip().replace(" ","_")


# ── Proposal storage ──────────────────────────────────────────────────────────

def load_proposals():
    DATA_DIR.mkdir(exist_ok=True)
    if not PROPOSALS_FILE.exists():
        return []
    try:
        with open(PROPOSALS_FILE) as f:
            return json.load(f).get("proposals", [])
    except Exception:
        return []

def save_proposals(proposals):
    DATA_DIR.mkdir(exist_ok=True)
    with open(PROPOSALS_FILE, "w") as f:
        json.dump({"proposals": proposals}, f, indent=2)

def add_proposal(record):
    proposals = load_proposals()
    # Keep max 50 saved proposals
    proposals = [p for p in proposals if p["id"] != record["id"]]
    proposals.insert(0, record)
    proposals = proposals[:50]
    save_proposals(proposals)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/proposals")
def api_proposals():
    proposals = load_proposals()
    # Annotate which files still exist on disk
    for p in proposals:
        job_dir = OUTPUTS_DIR / p["id"]
        prefix  = safe_prefix(p["company"])
        p["files"] = {
            "letter_pdf":    (job_dir / f"{prefix}_Intro_Letter.pdf").exists(),
            "proposal_pdf":  (job_dir / f"{prefix}_Proposal.pdf").exists(),
            "proposal_pptx": (job_dir / f"{prefix}_Proposal.pptx").exists(),
        }
    return jsonify(proposals)


@app.route("/api/proposals/<proposal_id>", methods=["DELETE"])
def delete_proposal(proposal_id):
    proposals = load_proposals()
    proposals = [p for p in proposals if p["id"] != proposal_id]
    save_proposals(proposals)
    return jsonify({"ok": True})


@app.route("/generate", methods=["POST"])
def generate():
    company       = request.form.get("company", "").strip()
    contact       = request.form.get("contact", "").strip()
    date          = request.form.get("date", "").strip()
    services      = request.form.getlist("services")
    cost_lines    = [c.strip() for c in request.form.getlist("cost_lines") if c.strip()]
    notes         = request.form.get("notes", "").strip()
    body_override = request.form.get("body_override", "").strip()

    if not company or not date:
        return jsonify({"error": "Company name and date are required."}), 400

    job_id  = uuid.uuid4().hex[:8]
    job_dir = OUTPUTS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    logo_path     = str(DEFAULT_LOGO)
    base_pdf_path = str(DEFAULT_BASE_PDF) if DEFAULT_BASE_PDF.exists() else ""

    prefix        = safe_prefix(company)
    letter_pdf    = str(job_dir / f"{prefix}_Intro_Letter.pdf")
    proposal_pdf  = str(job_dir / f"{prefix}_Proposal.pdf")
    proposal_pptx = str(job_dir / f"{prefix}_Proposal.pptx")
    env           = {**os.environ, "XCELERATE_LOGO": logo_path}
    errors        = []

    # Generate intro letter
    letter_cmd = [
        sys.executable, str(SCRIPTS_DIR / "generate_letter.py"),
        "--company", company, "--date", date,
        "--services", ", ".join(services) if services else "our full suite of services",
        "--output", letter_pdf,
    ]
    if contact:       letter_cmd += ["--contact", contact]
    if body_override: letter_cmd += ["--body", body_override]
    r = subprocess.run(letter_cmd, capture_output=True, text=True, env=env)
    if r.returncode != 0: errors.append(f"Letter: {r.stderr.strip()}")

    # Generate proposal PDF + PPTX
    proposal_cmd = [
        sys.executable, str(SCRIPTS_DIR / "generate_proposal.py"),
        "--company", company, "--date", date, "--output", proposal_pdf,
    ]
    if contact:        proposal_cmd += ["--contact",  contact]
    if base_pdf_path:  proposal_cmd += ["--base-pdf", base_pdf_path]
    if services:       proposal_cmd += ["--services", "|".join(services)]
    if cost_lines:     proposal_cmd += ["--costs",    "|".join(cost_lines)]
    if notes:          proposal_cmd += ["--notes",    notes]
    r2 = subprocess.run(proposal_cmd, capture_output=True, text=True, env=env)
    if r2.returncode != 0: errors.append(f"Proposal: {r2.stderr.strip()}")

    produced = {
        "letter_pdf":    os.path.exists(letter_pdf),
        "proposal_pdf":  os.path.exists(proposal_pdf),
        "proposal_pptx": os.path.exists(proposal_pptx),
    }

    if not any(produced.values()):
        return jsonify({"error": "Generation failed. " + " | ".join(errors)}), 500

    # Save proposal metadata
    add_proposal({
        "id":            job_id,
        "company":       company,
        "contact":       contact,
        "date":          date,
        "services":      services,
        "cost_lines":    cost_lines,
        "notes":         notes,
        "body_override": body_override,
        "created_at":    datetime.now().strftime("%B %d, %Y at %I:%M %p"),
        "produced":      produced,
    })

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
        return "File not found — please regenerate this proposal.", 404
    return send_file(str(file_path), as_attachment=True, download_name=safe_name)


@app.route("/regenerate/<proposal_id>", methods=["POST"])
def regenerate(proposal_id):
    proposals = load_proposals()
    saved = next((p for p in proposals if p["id"] == proposal_id), None)
    if not saved:
        return jsonify({"error": "Proposal not found"}), 404

    # Replay the generation with saved data
    job_id  = uuid.uuid4().hex[:8]
    job_dir = OUTPUTS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    company       = saved["company"]
    contact       = saved.get("contact", "")
    date          = saved["date"]
    services      = saved.get("services", [])
    cost_lines    = saved.get("cost_lines", [])
    notes         = saved.get("notes", "")
    body_override = saved.get("body_override", "")

    logo_path     = str(DEFAULT_LOGO)
    base_pdf_path = str(DEFAULT_BASE_PDF) if DEFAULT_BASE_PDF.exists() else ""
    prefix        = safe_prefix(company)
    letter_pdf    = str(job_dir / f"{prefix}_Intro_Letter.pdf")
    proposal_pdf  = str(job_dir / f"{prefix}_Proposal.pdf")
    proposal_pptx = str(job_dir / f"{prefix}_Proposal.pptx")
    env           = {**os.environ, "XCELERATE_LOGO": logo_path}

    letter_cmd = [
        sys.executable, str(SCRIPTS_DIR / "generate_letter.py"),
        "--company", company, "--date", date,
        "--services", ", ".join(services) if services else "our full suite of services",
        "--output", letter_pdf,
    ]
    if contact:       letter_cmd += ["--contact", contact]
    if body_override: letter_cmd += ["--body", body_override]
    subprocess.run(letter_cmd, capture_output=True, env=env)

    proposal_cmd = [
        sys.executable, str(SCRIPTS_DIR / "generate_proposal.py"),
        "--company", company, "--date", date, "--output", proposal_pdf,
    ]
    if contact:       proposal_cmd += ["--contact",  contact]
    if base_pdf_path: proposal_cmd += ["--base-pdf", base_pdf_path]
    if services:      proposal_cmd += ["--services", "|".join(services)]
    if cost_lines:    proposal_cmd += ["--costs",    "|".join(cost_lines)]
    if notes:         proposal_cmd += ["--notes",    notes]
    subprocess.run(proposal_cmd, capture_output=True, env=env)

    produced = {
        "letter_pdf":    os.path.exists(letter_pdf),
        "proposal_pdf":  os.path.exists(proposal_pdf),
        "proposal_pptx": os.path.exists(proposal_pptx),
    }

    return jsonify({
        "job_id":             job_id,
        "company":            company,
        "date":               date,
        "produced":           produced,
        "errors":             [],
        "letter_name":        f"{prefix}_Intro_Letter.pdf",
        "proposal_pdf_name":  f"{prefix}_Proposal.pdf",
        "proposal_pptx_name": f"{prefix}_Proposal.pptx",
    })


@app.route("/parse-transcript", methods=["POST"])
def parse_transcript():
    """
    Accept a raw sales-call transcript and return structured proposal data
    by calling the Anthropic API.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"error": "ANTHROPIC_API_KEY is not configured on this server. Add it in Railway → Variables."}), 503

    if not _ANTHROPIC_AVAILABLE:
        return jsonify({"error": "The 'anthropic' Python package is not installed. Add it to requirements.txt and redeploy."}), 503

    data = request.get_json(force=True, silent=True) or {}
    transcript = (data.get("transcript") or "").strip()
    if not transcript:
        return jsonify({"error": "No transcript text provided."}), 400

    catalog_str = "\n".join(f"- {s}" for s in SERVICES_CATALOG)

    prompt = f"""You are an expert assistant helping Xcelerate Growth Partners prepare a proposal after a sales call.

Read the transcript below and extract the following information. Return ONLY a single valid JSON object — no markdown, no explanation, no code fences.

JSON schema:
{{
  "company":        "Prospect company name (string)",
  "contact":        "Prospect contact person full name (string, empty string if unknown)",
  "date":           "Proposed meeting or presentation date formatted as 'Month DD, YYYY' (string, empty string if not mentioned)",
  "services":       ["Exact service names from catalog only (array of strings)"],
  "service_costs":  {{"Service Name": "cost string e.g. $3,500/mo — only include if discussed, otherwise omit"}},
  "notes":          "1–2 sentence note for the services slide summarizing the recommended engagement approach (string)",
  "body_override":  "3–4 paragraph custom intro letter body written in a warm, professional tone. Reference the company name, key pain points or goals mentioned in the call, and why Xcelerate is uniquely positioned to help. Do NOT include salutation or signature — body paragraphs only (string)"
}}

Service catalog (use ONLY these exact strings for the 'services' array):
{catalog_str}

Transcript:
{transcript}"""

    try:
        client = _anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()

        # Strip markdown code fences if model wrapped the JSON
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.strip()

        parsed = json.loads(raw)
        return jsonify({"ok": True, "data": parsed})

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Could not parse Claude's response as JSON: {e}", "raw": raw[:500]}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    OUTPUTS_DIR.mkdir(exist_ok=True)
    DATA_DIR.mkdir(exist_ok=True)
    port  = int(os.environ.get("PORT", 5050))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print(f"\n✅  Xcelerate Proposal Generator running at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=debug)
