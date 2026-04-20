#!/usr/bin/env python3
"""
Xcelerate Growth Partners — Proposal Package Generator
"""

import os, sys, uuid, json, re, subprocess
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
# DATA_DIR and OUTPUTS_DIR can be overridden via env vars so that Railway
# persistent volumes (or any external mount) survive redeploys.
# Example Railway setup:
#   DATA_DIR    = /data
#   OUTPUTS_DIR = /data/outputs
#   Then add a Railway Volume mounted at /data
BASE_DIR      = Path(__file__).parent
SCRIPTS_DIR   = BASE_DIR / "scripts"
ASSETS_DIR    = BASE_DIR / "assets"
OUTPUTS_DIR   = Path(os.environ.get("OUTPUTS_DIR", str(BASE_DIR / "outputs")))
DATA_DIR      = Path(os.environ.get("DATA_DIR",    str(BASE_DIR / "data")))
PROPOSALS_FILE  = DATA_DIR / "proposals.json"
ANALYSES_FILE   = DATA_DIR / "analyses.json"

DEFAULT_LOGO      = ASSETS_DIR / "xcelerate_logo.png"
DEFAULT_BASE_PDF  = ASSETS_DIR / "base_proposal_template.pdf"


def extract_json(raw: str) -> dict:
    """
    Robustly extract a JSON object from a Claude response that may contain
    extra text before/after the JSON, or markdown code fences.
    """
    # Strip markdown fences
    text = raw.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()

    # Try direct parse first (fast path)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Find the outermost { ... } block and parse that
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        return json.loads(match.group())

    raise ValueError(f"No JSON object found in response: {text[:200]}")

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


# ── Analyses storage ──────────────────────────────────────────────────────────

def load_analyses():
    DATA_DIR.mkdir(exist_ok=True)
    if not ANALYSES_FILE.exists():
        return []
    try:
        with open(ANALYSES_FILE) as f:
            return json.load(f).get("analyses", [])
    except Exception:
        return []

def save_analyses(analyses):
    DATA_DIR.mkdir(exist_ok=True)
    with open(ANALYSES_FILE, "w") as f:
        json.dump({"analyses": analyses}, f, indent=2)

def add_analysis(record):
    analyses = load_analyses()
    analyses = [a for a in analyses if a["id"] != record["id"]]
    analyses.insert(0, record)
    analyses = analyses[:100]
    save_analyses(analyses)


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


# ── Analyses API ──────────────────────────────────────────────────────────────

@app.route("/api/analyses")
def api_analyses():
    return jsonify(load_analyses())

@app.route("/api/analyses/<analysis_id>")
def get_analysis(analysis_id):
    analyses = load_analyses()
    match = next((a for a in analyses if a.get("id") == analysis_id), None)
    if not match:
        return jsonify({"error": "Not found"}), 404
    return jsonify(match)

@app.route("/api/analyses", methods=["POST"])
def save_analysis():
    data = request.get_json(force=True, silent=True) or {}
    if not data:
        return jsonify({"error": "No data"}), 400
    record = dict(data)
    record["id"]         = str(uuid.uuid4())
    record["created_at"] = datetime.now().strftime("%b %d, %Y %I:%M %p")
    add_analysis(record)
    return jsonify({"ok": True, "id": record["id"]})

@app.route("/api/analyses/<analysis_id>", methods=["DELETE"])
def delete_analysis(analysis_id):
    analyses = load_analyses()
    analyses = [a for a in analyses if a.get("id") != analysis_id]
    save_analyses(analyses)
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
  "company":         "Prospect company name (string)",
  "contact":         "Prospect contact person full name (string, empty string if unknown)",
  "date":            "Proposed meeting or presentation date formatted as 'Month DD, YYYY' (string, empty string if not mentioned)",
  "services":        ["Exact service names from catalog that were discussed (array of strings — use ONLY the exact strings listed below)"],
  "service_costs":   {{"Service Name": "cost string e.g. $3,500/mo — only include if explicitly discussed, otherwise omit"}},
  "custom_services": [{{"name": "Short descriptive service name", "cost": "cost string if mentioned, otherwise empty string"}}],
  "notes":           "1–2 sentence note for the services slide summarizing the recommended engagement approach (string)",
  "body_override":   "3–4 paragraph custom intro letter body written in a warm, professional tone. Reference the company name, key pain points or goals mentioned in the call, and why Xcelerate is uniquely positioned to help. Do NOT include salutation or signature — body paragraphs only (string)"
}}

IMPORTANT — for "custom_services": if the transcript mentions any services, needs, or deliverables that do NOT appear in the catalog below, add each one as an object in the "custom_services" array with a short, clear name and any cost mentioned. Do NOT force them into the catalog list. Examples of custom services: "Executive Coaching Program", "Monthly Market Briefings", "Custom Advisor Scorecard", "Onboarding Training Series".

Service catalog (use ONLY these exact strings in the "services" array):
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
        parsed = extract_json(raw)
        return jsonify({"ok": True, "data": parsed})

    except (json.JSONDecodeError, ValueError) as e:
        return jsonify({"error": f"Could not parse Claude's response as JSON: {e}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/sd-clarify", methods=["POST"])
def sd_clarify():
    """
    Step 1 of the conversational Sales Director: quickly read the transcript/email
    and return a brief summary + 2 targeted clarifying questions.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"error": "ANTHROPIC_API_KEY is not configured. Add it in Railway → Variables."}), 503
    if not _ANTHROPIC_AVAILABLE:
        return jsonify({"error": "The 'anthropic' package is not installed."}), 503

    data       = request.get_json(force=True, silent=True) or {}
    transcript = (data.get("transcript") or "").strip()
    email      = (data.get("email") or "").strip()

    if not transcript and not email:
        return jsonify({"error": "Please provide a transcript or email."}), 400

    input_text = ""
    if transcript: input_text += f"TRANSCRIPT:\n{transcript}\n\n"
    if email:      input_text += f"EMAIL:\n{email}\n\n"

    catalog_str = "\n".join(f"- {s}" for s in SERVICES_CATALOG)

    prompt = f"""You are the Sales Director at Xcelerate Growth Partners reviewing a sales opportunity.

Read the following transcript/email and do TWO things:
1. Write a 1–2 sentence summary of what you understand about this prospect and their needs.
2. Ask exactly 2 targeted clarifying questions — the 2 questions whose answers would MOST improve your package recommendation. Make them specific to THIS prospect, not generic.

Return ONLY valid JSON — no markdown, no code fences:
{{
  "summary": "1-2 sentence understanding of the prospect and their situation",
  "questions": [
    "First specific clarifying question",
    "Second specific clarifying question"
  ]
}}

Xcelerate service catalog (for context):
{catalog_str}

{input_text}"""

    try:
        client  = _anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()
        parsed = extract_json(raw)
        return jsonify({"ok": True, "data": parsed})
    except (json.JSONDecodeError, ValueError) as e:
        return jsonify({"error": f"Could not parse response: {e}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/analyze-opportunity", methods=["POST"])
def analyze_opportunity():
    """
    Sales Director analysis: accepts a transcript and/or email, references past
    proposals for pricing context, and returns structured package recommendations,
    pricing ranges, a strategic rationale, and deal-coaching notes.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"error": "ANTHROPIC_API_KEY is not configured on this server. Add it in Railway → Variables."}), 503
    if not _ANTHROPIC_AVAILABLE:
        return jsonify({"error": "The 'anthropic' Python package is not installed."}), 503

    data              = request.get_json(force=True, silent=True) or {}
    transcript        = (data.get("transcript") or "").strip()
    email             = (data.get("email") or "").strip()
    clarifications    = data.get("clarifications") or {}   # {question: answer} pairs
    followup_question = (data.get("followup_question") or "").strip()

    if not transcript and not email:
        return jsonify({"error": "Please provide a transcript and/or email thread."}), 400

    # ── Follow-up chat mode ───────────────────────────────────────────────────
    if followup_question:
        context = ""
        if transcript: context += f"TRANSCRIPT:\n{transcript}\n\n"
        if email:      context += f"EMAIL:\n{email}\n\n"
        followup_prompt = f"""{context}A sales rep is asking a follow-up question about this deal.
Answer concisely and directly as an experienced Sales Director at Xcelerate Growth Partners.
Be specific to THIS prospect — don't give generic advice.

Question: {followup_question}"""
        try:
            client  = _anthropic.Anthropic(api_key=api_key)
            message = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=600,
                messages=[{"role": "user", "content": followup_prompt}],
            )
            answer = message.content[0].text.strip()
            return jsonify({"ok": True, "data": {"followup_answer": answer}})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    # ── Build past-proposals context (last 8 for pricing reference) ───────────
    past = load_proposals()[:8]
    past_context = ""
    if past:
        lines = []
        for p in past:
            svcs  = ", ".join(p.get("services", [])) or "N/A"
            costs = " | ".join(p.get("costs", [])) or "N/A"
            lines.append(f"  • {p.get('company','?')} — Services: {svcs} — Fees: {costs}")
        past_context = "PAST PROPOSAL REFERENCE (use for pricing calibration):\n" + "\n".join(lines)
    else:
        past_context = "PAST PROPOSAL REFERENCE: No past proposals on file yet."

    catalog_str = "\n".join(f"- {s}" for s in SERVICES_CATALOG)

    input_text = ""
    if transcript:
        input_text += f"SALES CALL TRANSCRIPT:\n{transcript}\n\n"
    if email:
        input_text += f"EMAIL THREAD:\n{email}\n\n"
    if clarifications:
        claras = "\n".join(f"  Q: {q}\n  A: {a}" for q, a in clarifications.items())
        input_text += f"CLARIFYING ANSWERS FROM SALES REP:\n{claras}\n\n"

    system_prompt = f"""You are the Sales Director at Xcelerate Growth Partners — a firm that delivers
custom growth, leadership, and practice management programs to wealth management firms,
asset managers, and financial advisor teams.

Your job is to analyze every sales opportunity and give sharp, actionable recommendations.
You speak like an experienced deal-maker: direct, decisive, and specific.

{past_context}

XCELERATE'S SERVICE CATALOG (with typical price ranges):
{catalog_str}

Typical fee context:
- Asset managers / large wirehouses: $8,000–$25,000/month retainer
- Regional broker-dealers: $4,000–$12,000/month
- Individual advisor teams / smaller firms: $2,500–$6,000/month
- One-time projects (succession, workshops, keynotes): $5,000–$15,000
- Bundle 3+ services: 10–15% discount is common

Return ONLY a single valid JSON object with NO markdown, NO code fences, NO explanation.

JSON schema:
{{
  "company": "Prospect company name extracted from the input",
  "contact": "Contact person full name",
  "contact_email": "Contact email if mentioned, else empty string",
  "contact_phone": "Contact phone if mentioned, else empty string",
  "date": "Meeting or follow-up date if mentioned (e.g. 'April 22, 2026'), else empty string",
  "opportunity_summary": "2-3 sentence read on this prospect — who they are, what they really need, and your confidence level on closing",
  "recommended_services": [
    {{
      "name": "Exact service name from catalog",
      "rationale": "One sentence on why this fits THIS specific prospect",
      "price_range": "Use '/mo' suffix for monthly retainers (e.g. '$3,500-$5,000/mo') and 'one-time' label for project fees (e.g. '$8,000-$12,000 one-time')"
    }}
  ],
  "custom_services": [
    {{
      "name": "Any service not in catalog but clearly needed",
      "rationale": "Why",
      "price_range": "Use '/mo' or 'one-time' suffix as appropriate",
      "is_custom": true
    }}
  ],
  "total_range": "Describe clearly, e.g. '$5,500-$8,000/mo retainer + $12,000 one-time setup' or '$7,500-$11,000/mo'",
  "strategic_rationale": "2-3 sentences on the overall package strategy and why this combination wins the deal",
  "objections": [
    {{
      "objection": "Specific concern the prospect raised or is likely to raise",
      "rebuttal": "Specific, confident response — tie it to something they said"
    }}
  ],
  "competitive_angle": "2-3 sentences on what makes Xcelerate uniquely positioned to win this vs. any alternative",
  "follow_up_email": {{
    "subject": "Email subject line",
    "body": "Full email body — ready to send. Reference the call. End with a clear ask. Sign off as Jim Tracy."
  }},
  "next_steps": ["Action 1", "Action 2", "Action 3"],
  "intro_notes": "1-2 sentences for the intro letter — the angle or hook that will resonate most with this prospect",
  "body_override": "If you can write a compelling personalized 2-paragraph intro letter body for this prospect, include it here. Otherwise empty string."
}}"""

    try:
        client  = _anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=3000,
            system=system_prompt,
            messages=[{"role": "user", "content": input_text}],
        )
        raw = message.content[0].text.strip()
        parsed = extract_json(raw)
        return jsonify({"ok": True, "data": parsed})

    except (json.JSONDecodeError, ValueError) as e:
        return jsonify({"error": f"Could not parse response as JSON: {e}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    OUTPUTS_DIR.mkdir(exist_ok=True)
    DATA_DIR.mkdir(exist_ok=True)
    port  = int(os.environ.get("PORT", 5050))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print(f"\n✅  Xcelerate Proposal Generator running at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=debug)
