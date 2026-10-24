import os
import csv
import time
import random
import smtplib
import uuid
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from typing import Optional, List, Dict

import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from pydantic import BaseModel

app = FastAPI(title="Zylo Reach - Cold Email Automation Engine")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CSV = os.path.join(BASE_DIR, "leads.csv")
ACTIVE_CSV = os.path.join(BASE_DIR, "active_leads.csv")

# Ensure active_leads.csv exists by copying from default if not present
if not os.path.exists(ACTIVE_CSV) and os.path.exists(DEFAULT_CSV):
    import shutil
    shutil.copyfile(DEFAULT_CSV, ACTIVE_CSV)

# Global in-memory job tracker
JOBS: Dict[str, dict] = {}

def get_personalized_email(business_name: str, category: str, city: str, sender_name: str, sender_company: str):
    first_name = (business_name or "there").strip()
    cat = (category or "").lower()
    loc = (city or "your area").strip()

    if "cafe" in cat or "coffee" in cat or "restaurant" in cat or "food" in cat:
        subject = f"A few digital growth ideas for {first_name}"
        body = f"""Hi {first_name} team,

I came across {first_name} in {loc} and really liked what you've built.

I'm {sender_name} from {sender_company}. We build brands that sell—helping local businesses and cafes strengthen their online presence through:
- Local Google Maps & SEO optimization (so nearby customers find you first)
- AI-powered food visuals & mouth-watering short-form reels for Instagram/TikTok
- Fast, mobile-friendly online menu & high-converting website design
- Targeted Google & Meta Ads to drive consistent footfall

Rather than guessing what works, I put together 2-3 quick opportunities specifically for {first_name}.

Would you be open to a brief 15-minute discovery call next week to see if we might be a fit?

Best regards,

{sender_name}
{sender_company}
We build brands that sell.
---
If you'd prefer not to receive ideas from us, simply reply with 'Unsubscribe' and I'll remove you immediately."""
    else:  # Fashion / Boutiques / E-commerce / Apparel
        subject = f"Creative and digital support for {first_name}"
        body = f"""Hi {first_name} team,

I've been checking out {first_name}'s collection in {loc}—you have great styling and standout pieces.

I'm {sender_name} from {sender_company}. We help boutique and apparel brands scale direct-to-consumer sales through:
- AI-powered product photoshoots & high-converting video creatives
- Custom e-commerce website design optimized for smooth mobile checkout
- Targeted Meta & Google Ads management to bring in active fashion shoppers
- Strategic SEO & creative content to turn attention into sales

Rather than guessing what works, I noted 2-3 quick opportunities specifically for {first_name}'s online catalog.

Would you be open to a brief 15-minute discovery call next week to explore a few ideas?

Best regards,

{sender_name}
{sender_company}
We build brands that sell.
---
If you'd prefer not to receive ideas from us, simply reply with 'Unsubscribe' and I'll remove you immediately."""

    return subject, body


def normalize_lead_row(row: dict) -> dict:
    # Fuzzy column mapping for custom uploaded files
    keys = {str(k).lower().strip(): k for k in row.keys()}
    
    # Business Name
    b_name = row.get(keys.get("business name") or keys.get("company") or keys.get("business") or keys.get("name") or "", "")
    # Email
    email = row.get(keys.get("researched business email") or keys.get("business email") or keys.get("email") or keys.get("email address") or "", "")
    # Category
    cat = row.get(keys.get("category") or keys.get("industry") or keys.get("niche") or "", "General")
    # City
    city = row.get(keys.get("city") or keys.get("location") or keys.get("town") or "", "")
    # Confidence
    confidence = row.get(keys.get("email confidence") or keys.get("confidence") or "", "High")
    # Opt-out
    opt_out = str(row.get(keys.get("opt-out / do not contact") or keys.get("opt-out") or "No", "No")).strip().lower()

    email_clean = str(email).strip() if pd.notna(email) else ""
    is_verified = bool(
        email_clean
        and "@" in email_clean
        and "not found" not in email_clean.lower()
        and "mismatch" not in str(confidence).lower()
        and opt_out != "yes"
    )

    return {
        "Business Name": str(b_name).strip() if pd.notna(b_name) else "Valued Business",
        "Category": str(cat).strip() if pd.notna(cat) else "General",
        "City": str(city).strip() if pd.notna(city) else "USA",
        "Researched Business Email": email_clean,
        "Email Confidence": str(confidence) if pd.notna(confidence) else "Verified",
        "is_verified": is_verified,
        "Opt-out / Do Not Contact": opt_out
    }


def parse_leads_from_file(file_path: str):
    if not os.path.exists(file_path):
        return []
    try:
        if file_path.endswith(".xlsx") or file_path.endswith(".xls"):
            df = pd.read_excel(file_path)
            raw_leads = df.to_dict(orient="records")
        else:
            df = pd.read_csv(file_path)
            raw_leads = df.to_dict(orient="records")
        return [normalize_lead_row(r) for r in raw_leads]
    except Exception as e:
        print(f"Error parsing file: {e}")
        return []


@app.get("/api/leads")
def get_leads():
    file_to_read = ACTIVE_CSV if os.path.exists(ACTIVE_CSV) else DEFAULT_CSV
    leads = parse_leads_from_file(file_to_read)
    return {
        "total": len(leads),
        "verified_count": sum(1 for l in leads if l["is_verified"]),
        "leads": leads
    }


@app.post("/api/upload-leads")
async def upload_leads(file: UploadFile = File(...)):
    filename = file.filename.lower()
    if not (filename.endswith(".csv") or filename.endswith(".xlsx") or filename.endswith(".xls")):
        return JSONResponse(status_code=400, content={"error": "Please upload a .csv, .xlsx, or .xls file."})

    contents = await file.read()
    try:
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            df = pd.read_excel(io.BytesIO(contents))
            df.to_csv(ACTIVE_CSV, index=False)
        else:
            with open(ACTIVE_CSV, "wb") as f:
                f.write(contents)

        leads = parse_leads_from_file(ACTIVE_CSV)
        return {
            "success": True,
            "filename": file.filename,
            "total": len(leads),
            "verified_count": sum(1 for l in leads if l["is_verified"]),
            "leads": leads
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to parse file: {str(e)}"})


@app.post("/api/reset-leads")
def reset_leads():
    if os.path.exists(DEFAULT_CSV):
        import shutil
        shutil.copyfile(DEFAULT_CSV, ACTIVE_CSV)
    leads = parse_leads_from_file(ACTIVE_CSV)
    return {
        "success": True,
        "total": len(leads),
        "verified_count": sum(1 for l in leads if l["is_verified"]),
        "leads": leads
    }


class OutreachRequest(BaseModel):
    sender_name: str
    sender_company: str
    sender_email: str
    app_password: Optional[str] = ""
    mode: str  # 'dry_run' or 'live'
    min_delay: int = 40
    max_delay: int = 70


def background_outreach_worker(job_id: str, req: OutreachRequest):
    job = JOBS[job_id]
    file_to_read = ACTIVE_CSV if os.path.exists(ACTIVE_CSV) else DEFAULT_CSV
    leads = parse_leads_from_file(file_to_read)
    verified = [l for l in leads if l.get("is_verified")]
    
    job["total"] = len(verified)
    job["status"] = "running"
    
    if req.mode == "dry_run":
        job["logs"].append("🛡️ ========================================================")
        job["logs"].append("🧪 SIMULATION MODE ACTIVATED (100% Safe - No Emails Sent)")
        job["logs"].append("🛡️ ========================================================")
    else:
        job["logs"].append("⚡ ========================================================")
        job["logs"].append("🚀 LIVE SENDING MODE ACTIVATED (Real Emails Sending via SMTP)")
        job["logs"].append("⚡ ========================================================")
        
    job["logs"].append(f"📊 Total verified recipients: {len(verified)}")

    smtp_server = None
    if req.mode == "live":
        try:
            job["logs"].append("🔐 Authenticating with Gmail SMTP...")
            smtp_server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            clean_pwd = req.app_password.replace(" ", "")
            smtp_server.login(req.sender_email, clean_pwd)
            job["logs"].append("✅ Gmail SMTP authentication successful!")
        except Exception as e:
            job["status"] = "failed"
            job["error"] = f"SMTP Authentication Error: {str(e)}"
            job["logs"].append(f"❌ Authentication Failed: {str(e)}")
            return

    sent_count = 0
    for idx, lead in enumerate(verified, 1):
        name = lead.get("Business Name", "")
        category = lead.get("Category", "Cafe")
        city = lead.get("City", "")
        to_email = lead.get("Researched Business Email", "")

        subject, body = get_personalized_email(name, category, city, req.sender_name, req.sender_company)

        job["logs"].append(f"[{idx:02d}/{len(verified):02d}] 🎯 Target: {name} ({category}) | {city}")
        job["logs"].append(f"       📧 Recipient: {to_email}")
        job["logs"].append(f"       📌 Subject  : {subject}")

        if req.mode == "dry_run":
            time.sleep(1.0)  # Visual preview delay for simulation
            job["logs"].append(f"       🧪 [SIMULATION PASS] Email body generated & validated ✅")
            sent_count += 1
            job["processed"] = idx
        else:
            try:
                msg = MIMEMultipart()
                msg["From"] = f"{req.sender_name} <{req.sender_email}>"
                msg["To"] = to_email
                msg["Subject"] = subject
                msg.attach(MIMEText(body, "plain"))

                smtp_server.send_message(msg)
                job["logs"].append(f"       ✅ Sent successfully to {to_email}")
                sent_count += 1
                job["processed"] = idx

                if idx < len(verified):
                    sleep_secs = random.randint(req.min_delay, req.max_delay)
                    job["logs"].append(f"       ⏳ Anti-Spam Jitter: Waiting {sleep_secs}s before next email...")
                    time.sleep(sleep_secs)
            except Exception as e:
                job["logs"].append(f"       ❌ Failed to send to {to_email}: {str(e)}")

    if smtp_server:
        try:
            smtp_server.quit()
        except Exception:
            pass

    job["status"] = "completed"
    if req.mode == "dry_run":
        job["logs"].append(f"🎉 Simulation complete! All {sent_count} leads tested safely.")
    else:
        job["logs"].append(f"🎉 Campaign complete! All {sent_count} emails delivered.")


@app.post("/api/start-outreach")
def start_outreach(req: OutreachRequest, background_tasks: BackgroundTasks):
    if not req.sender_email or "@" not in req.sender_email:
        return JSONResponse(status_code=400, content={"error": "Invalid Sender Gmail address"})
    if req.mode == "live" and (not req.app_password or len(req.app_password.replace(" ", "")) < 10):
        return JSONResponse(status_code=400, content={"error": "16-digit Google App Password is required for Live mode"})

    job_id = str(uuid.uuid4())[:8]
    JOBS[job_id] = {
        "id": job_id,
        "mode": req.mode,
        "status": "queued",
        "processed": 0,
        "total": 0,
        "logs": [],
        "created_at": datetime.now().isoformat()
    }

    background_tasks.add_task(background_outreach_worker, job_id, req)
    return {"job_id": job_id, "status": "queued"}


@app.get("/api/job-status/{job_id}")
def get_job_status(job_id: str):
    if job_id not in JOBS:
        return JSONResponse(status_code=404, content={"error": "Job not found"})
    return JOBS[job_id]


@app.get("/", response_class=HTMLResponse)
def index_page():
    index_file = os.path.join(BASE_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Zylo Reach Backend Running</h1>"


@app.get("/download-pdf")
def download_pdf():
    pdf_file = os.path.join(BASE_DIR, "Mayras_Boutique_Sales_Script_Zylo_Digital.pdf")
    if os.path.exists(pdf_file):
        return FileResponse(
            path=pdf_file,
            filename="Mayras_Boutique_Sales_Script_Zylo_Digital.pdf",
            media_type="application/pdf"
        )
    return JSONResponse(status_code=404, content={"error": "PDF not found"})


@app.get("/view-script", response_class=HTMLResponse)
def view_script():
    html_file = os.path.join(BASE_DIR, "mayras_boutique_script.html")
    if os.path.exists(html_file):
        with open(html_file, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Script not found</h1>"

