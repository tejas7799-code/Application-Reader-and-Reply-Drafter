#!/usr/bin/env python3
"""Application Reader and Reply Drafter (rules-based, zero API calls).

Usage:  python reader.py [applications.csv]
Output: output.csv, output.xlsx (if openpyxl is installed), review.html
"""
import csv
import html
import json
import re
import sys
from datetime import datetime

INPUT = sys.argv[1] if len(sys.argv) > 1 else "applications.csv"

# ---------------------------------------------------------------- cleaning
MISSING = {"", "n/a", "na", "none", "null", "not provided", "-"}

COLLEGE_ALIASES = {
    "Mohan Babu University": ["mohan babu university", "mohan babu univeristy", "mbu"],
    "Sri Venkateswara University College of Engineering": [
        "sri venkateswara university college of engineering",
        "sri venkateswara univ. college of engineering",
        "sv university college of engg"],
    "Annamacharya Institute of Technology and Sciences": [
        "annamacharya institute of technology and sciences",
        "annamacharya inst. of tech & sciences"],
    "VIT Vellore": ["vit vellore", "vit", "vellore institute of technology"],
    "JNTU Anantapur": ["jntu anantapur", "jntua"],
    "SRM Institute of Science and Technology": [
        "srm institute of science and technology", "srm institute of sci & tech"],
    "Sree Vidyanikethan Engineering College": [
        "sree vidyanikethan engineering college", "sree vidyanikethan engg college", "svec"],
    "Chadalawada Ramanamma Engineering College": [
        "chadalawada ramanamma engineering college", "chadalawada ramanamma engg college"],
    "Anna University": ["anna university"],
}
COLLEGE_LOOKUP = {a: canon for canon, al in COLLEGE_ALIASES.items() for a in al}

DOMAIN_INFO = {
    "Software Development and Full-Stack Engineering":
        "In this track you would build and maintain web applications and have your code reviewed by engineers.",
    "Artificial Intelligence and Automation":
        "In this track you would build AI-assisted tools and automations that solve real business problems.",
    "Data Engineering, Cloud and Analytics":
        "In this track you would work on data pipelines, cloud services and dashboards.",
    "Software Testing and Quality Assurance":
        "In this track you would find bugs, write clear reports and build automated tests.",
}

# statement category -> (keywords, matching domain or None, tailored sentence)
STATEMENT_RULES = [
    ("ml", ["kaggle", "ml mini"], "Artificial Intelligence and Automation",
     "Your Kaggle ML projects are a good starting point."),
    ("data", ["patterns in data", "data pipelines"], "Data Engineering, Cloud and Analytics",
     "Your interest in data pipelines and monitoring fits well with what we do."),
    ("testing", ["breaking things", "bug reports", "test automation"], "Software Testing and Quality Assurance",
     "Your attention to bug reports and test automation is exactly what a QA team values."),
    ("project", ["fest registration"], "Software Development and Full-Stack Engineering",
     "Building a fest registration site is a solid first real-world project."),
    ("deadlines", ["deadlines and code reviews"], None,
     "Wanting to move from tutorials to production-quality work with code reviews is the right goal."),
]
GENERIC = ["hardworking and dedicated", "to gain knowledge", "want to learn new skills",
           "some more skills", "interested"]


def clean_missing(v):
    v = (v or "").strip()
    return "" if v.lower() in MISSING else v


def clean_name(n):
    return re.sub(r"\s+", " ", (n or "").strip()).title()


def clean_college(c):
    c = re.sub(r"\s+", " ", (c or "").strip())
    return COLLEGE_LOOKUP.get(c.lower(), c)


def clean_phone(p):
    p = (p or "").strip()
    digits = re.sub(r"\D", "", p)
    if p.startswith("+91") or (len(digits) == 12 and digits.startswith("91")):
        digits = digits[2:]
    ok = len(digits) == 10 and digits[0] in "6789"
    return (digits if ok else digits), ok


def valid_email(e):
    return bool(re.fullmatch(r"[\w.+-]+@[\w-]+(\.[\w-]+)+", (e or "").strip()))


def parse_date(s):
    for fmt in ("%Y-%m-%d %H:%M", "%d/%m/%Y %I:%M %p"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            pass
    return datetime.min


def classify_statement(s):
    s_low = s.lower()
    if not s:
        return "none", None, None
    for cat, kws, dom, line in STATEMENT_RULES:
        if any(k in s_low for k in kws):
            return cat, dom, line
    if len(s.split()) < 8 or any(g in s_low for g in GENERIC):
        return "generic", None, None
    return "specific", None, None


# ---------------------------------------------------------------- scoring
def score(row):
    pts, reasons = 10, []
    if not row["resume"]:
        pts -= 3; reasons.append("resume missing (-3)")
    if row["stmt_cat"] == "none":
        pts -= 4; reasons.append("statement missing (-4)")
    elif row["stmt_cat"] == "generic":
        pts -= 3; reasons.append("statement too short/generic (-3)")
    if not row["email_ok"]:
        pts -= 3; reasons.append("invalid email (-3)")
    if not row["phone_ok"]:
        pts -= 2; reasons.append("invalid phone (-2)")
    return max(pts, 0), "; ".join(reasons) or "complete: resume, valid contact details, specific statement"


# ---------------------------------------------------------------- reply
def draft_reply(r):
    first = r["name"].split()[0]
    asks = []
    if not r["resume"]:
        asks.append("a working link to your resume (Google Drive or PDF)")
    if r["stmt_cat"] in ("none", "generic"):
        asks.append("3-4 lines about a project you have done or something specific you want to build or learn")
    if not r["phone_ok"]:
        asks.append("your correct 10-digit mobile number (the one we have looks incomplete)")
    if r["mismatch"]:
        asks.append(f"confirmation of your preferred domain (you chose {r['domain']}, "
                    f"but your note points to {r['stmt_dom']})")
    lines = [f"Hi {first},", "",
             f"Thank you for applying to the InternFactory Internship Program for the "
             f"{r['domain']} track from {r['college']}."]
    if r["stmt_line"] and not r["mismatch"]:
        lines.append(r["stmt_line"])
    lines.append(DOMAIN_INFO.get(r["domain"], ""))
    if asks:
        lines += ["", "To complete your application, please reply with:"]
        lines += [f"  - {a}" for a in asks]
    else:
        lines += ["", "Your application is complete and we have it on file."]
    lines += ["", "Warm regards,", "InternFactory Hiring Team", "info@internfactory.in"]
    subject = ("InternFactory application: a few details needed" if asks
               else "InternFactory application received")
    return subject, "\n".join(l for l in lines if l is not None)


# ---------------------------------------------------------------- main
def load(path):
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for raw in csv.DictReader(f):
            phone, phone_ok = clean_phone(raw["mobile"])
            stmt = clean_missing(raw["statement_of_interest"])
            cat, sdom, sline = classify_statement(stmt)
            domain = raw["preferred_domain"].strip()
            rows.append({
                "id": raw["application_id"].strip(),
                "submitted": parse_date(raw["submitted_at"]),
                "name": clean_name(raw["full_name"]),
                "email": raw["email"].strip().lower(),
                "email_ok": valid_email(raw["email"]),
                "phone": phone, "phone_ok": phone_ok,
                "college": clean_college(raw["college"]),
                "domain": domain,
                "resume": clean_missing(raw["resume_link"]),
                "statement": stmt, "stmt_cat": cat, "stmt_dom": sdom, "stmt_line": sline,
                "mismatch": bool(sdom and sdom != domain),
            })
    return rows


def mark_duplicates(rows):
    groups = {}
    for r in rows:
        groups.setdefault(r["email"] if r["email_ok"] else r["id"], []).append(r)
    for g in groups.values():
        g.sort(key=lambda r: r["submitted"])
        keep = g[-1]  # latest submission wins
        for r in g:
            r["dup_of"] = "" if r is keep else keep["id"]
            r["has_dups"] = len(g) > 1


def main():
    rows = load(INPUT)
    mark_duplicates(rows)
    out = []
    for r in sorted(rows, key=lambda r: r["id"]):
        s, reason = score(r)
        flags = []
        if r["dup_of"]:
            flags.append(f"duplicate of {r['dup_of']} - do not send")
        elif r["has_dups"]:
            flags.append("has duplicate submissions - this is the latest")
        if not r["email_ok"]:
            flags.append("invalid email - cannot send, contact by phone")
        if s <= 4:
            flags.append("low score")
        if r["mismatch"]:
            flags.append("statement does not match chosen domain")
        subject, body = ("", "") if r["dup_of"] else draft_reply(r)
        out.append({
            "application_id": r["id"], "full_name": r["name"], "email": r["email"],
            "college": r["college"], "preferred_domain": r["domain"],
            "completeness_score": s, "score_reason": reason,
            "needs_human_check": "YES" if flags else "no", "check_reason": "; ".join(flags),
            "duplicate_of": r["dup_of"], "reply_subject": subject, "drafted_reply": body,
        })

    cols = list(out[0].keys())
    with open("output.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); w.writerows(out)

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        wb = Workbook(); ws = wb.active; ws.title = "Replies"
        ws.append(cols)
        for row in out:
            ws.append([row[c] for c in cols])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0B1020")
        widths = {"score_reason": 40, "check_reason": 40, "drafted_reply": 80, "full_name": 20,
                  "email": 32, "college": 38, "preferred_domain": 36, "reply_subject": 38}
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = widths.get(c, 16)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "C2"
        wb.save("output.xlsx")
        print("wrote output.xlsx")
    except ImportError:
        print("openpyxl not installed: skipped output.xlsx (output.csv has everything)")

    write_review(out)
    flagged = sum(1 for o in out if o["needs_human_check"] == "YES")
    print(f"{len(out)} applications | avg score {sum(o['completeness_score'] for o in out)/len(out):.1f} "
          f"| {flagged} need human check | 0 API calls, 0 cost")


def write_review(out):
    data = json.dumps(out)
    page = """<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>Review replies</title>
<style>body{font:15px system-ui;margin:0;background:#f5f6fa;color:#111}header{background:#0b1020;color:#fff;padding:14px 18px;position:sticky;top:0}
select,button,input{font:inherit;padding:6px 10px;margin-right:6px;border-radius:6px;border:1px solid #ccc}button{background:#2f5bff;color:#fff;border:0;cursor:pointer}
.card{background:#fff;margin:12px;padding:14px;border-radius:10px;border:1px solid #e3e5ec}.card.flag{border-left:5px solid #e8590c}
.meta{color:#555;font-size:13px}.score{font-weight:700}textarea{width:100%;box-sizing:border-box;height:210px;font:13px/1.45 ui-monospace,monospace;margin-top:8px}
.why{color:#b34700;font-size:13px}</style>
<header><b>Reply review</b> &nbsp;<select id=f><option value=all>All</option><option value=flag>Needs human check</option><option value=ok>Ready to send</option></select>
<input id=q placeholder="search name / id"><button onclick=dl()>Download edited CSV</button><span id=n></span></header><div id=app></div>
<script>const D=__DATA__;const app=document.getElementById('app');
function render(){const f=document.getElementById('f').value,q=document.getElementById('q').value.toLowerCase();app.innerHTML='';let c=0;
D.forEach((r,i)=>{const flag=r.needs_human_check==='YES';if(f==='flag'&&!flag||f==='ok'&&flag)return;
if(q&&!(r.full_name+r.application_id).toLowerCase().includes(q))return;c++;
const d=document.createElement('div');d.className='card'+(flag?' flag':'');
d.innerHTML=`<b></b> <span class=meta></span> <span class=score></span><div class=meta></div><div class=why></div><div class=meta></div><textarea></textarea><button>Copy reply</button>`;
const k=d.children;k[0].textContent=r.full_name;k[1].textContent=r.application_id+' | '+r.college;k[2].textContent='  Score '+r.completeness_score+'/10';
k[3].textContent=r.score_reason;k[4].textContent=flag?'Check: '+r.check_reason:'';k[5].textContent=r.reply_subject?'Subject: '+r.reply_subject:'';
const t=k[6];t.value=r.drafted_reply||'(no reply: duplicate)';t.oninput=()=>r.drafted_reply=t.value;
k[7].onclick=()=>navigator.clipboard.writeText(t.value);app.appendChild(d)});document.getElementById('n').textContent='  '+c+' shown'}
function dl(){const cols=Object.keys(D[0]);const e=v=>'"'+String(v).replace(/"/g,'""')+'"';
const s=[cols.join(',')].concat(D.map(r=>cols.map(c=>e(r[c])).join(','))).join('\\n');
const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([s],{type:'text/csv'}));a.download='output_edited.csv';a.click()}
f.onchange=q.oninput=render;render();</script>""".replace("__DATA__", data.replace("</", "<\\/"))
    with open("review.html", "w", encoding="utf-8") as f:
        f.write(page)
    print("wrote review.html")


if __name__ == "__main__":
    main()
