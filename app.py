import io
import os
import re
import base64
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# ----------------- AUTOMATED VISITOR COUNTER -----------------
VISITOR_FILE = "visitor_count.txt"

def get_and_increment_visitors() -> int:
    """Reads and increments the total visitor count once per user session."""
    count = 0
    if os.path.exists(VISITOR_FILE):
        try:
            with open(VISITOR_FILE, "r") as f:
                content = f.read().strip()
                count = int(content) if content.isdigit() else 0
        except Exception:
            count = 0
            
    # Increment only once when a new browser session starts
    if "user_counted" not in st.session_state:
        st.session_state["user_counted"] = True
        count += 1
        try:
            with open(VISITOR_FILE, "w") as f:
                f.write(str(count))
        except Exception:
            pass
            
    return count

total_visitors = get_and_increment_visitors()

# ----------------- PAGE CONFIGURATION & STYLES -----------------
st.set_page_config(page_title="Result Analyzer Portal", page_icon="🎓", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Poppins', sans-serif; }
    
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white; padding: 22px; border-radius: 12px; margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.15); text-align: center;
    }
    .visitor-badge {
        display: inline-block;
        background: rgba(255, 255, 255, 0.2);
        border: 1px solid rgba(255, 255, 255, 0.4);
        padding: 5px 15px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
        margin-top: 10px;
    }
    .metric-pass { 
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); 
        color: white; padding: 15px; border-radius: 10px; text-align: center; 
    }
    .metric-fail { 
        background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%); 
        color: white; padding: 15px; border-radius: 10px; text-align: center; 
    }
    .metric-info { 
        background: linear-gradient(135deg, #36D1DC 0%, #5B86E5 100%); 
        color: white; padding: 15px; border-radius: 10px; text-align: center; 
    }
    .subject-box { 
        background: white; border-left: 5px solid #2a5298; border-radius: 8px; 
        padding: 12px; margin-bottom: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); 
    }
    .inst-card {
        background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 10px;
        padding: 15px; margin-bottom: 20px; border-left: 5px solid #1e3c72;
    }
    .footer-box { 
        text-align: center; padding: 18px; margin-top: 35px; background: #f8f9fa; 
        border-top: 2px solid #dee2e6; font-size: 13px; color: #495057; border-radius: 8px; 
    }
    @media print {
        header, footer, [data-testid="stSidebar"], .stButton, .stDownloadButton {
            display: none !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# Main Banner with Visitor Badge
st.markdown(f"""
<div class="main-header">
    <h1 style="margin:0; font-weight:700;">🎓 Result Analyzer Portal</h1>
    <p style="margin:5px 0 0 0; font-size:14px; opacity: 0.95;">
        Automated Examination Result Analysis, Dynamic Paper List Mapping & Merit System
    </p>
    <div class="visitor-badge">
        👥 Total Portal Visitors: <b>{total_visitors}</b>
    </div>
</div>
""", unsafe_allow_html=True)

# ----------------- BUILT-IN FALLBACK SUBJECT TITLES -----------------
BUILTIN_TITLES = {
    # SPPU Zoology NEP & CBCS
    "ZOO-242-MN": "Animal Diversity (Minor)",
    "ZOO-244-MNP": "Zoology Practical (Minor)",
    "ZOO-292-MN": "Applied Zoology (Minor)",
    "ZOO-294-MNP": "Zoology Practical (Minor)",
    "ZOO-101": "Animal Systematics & Diversity I",
    "ZOO-102": "Animal Systematics & Diversity II",
    "ZOO-103": "Zoology Practical I",
    "ZOO-201": "Animal Diversity III",
    "ZOO-202": "Applied Zoology I",
    "ZOO-203": "Zoology Practical II",
    "35151": "Animal Diversity III",
    "35152": "Animal Diversity IV",
    "35153": "Applied Zoology I",
    "35154": "Applied Zoology II",
    "35155": "Cell & Molecular Biology",
    "35156": "Genetics",
    "35157": "Zoology Practical I",
    "35158": "Zoology Practical II",
    "35159": "Zoology Practical III",
    "36151": "Animal Physiology",
    "36152": "Molecular Biology",
    "36153": "Mammalian Histology",
    "36154": "Parasitology",
    "36155": "Evolutionary Biology",
    "36156": "General Entomology",
    # Common Arts / Commerce / General
    "CMA102T": "Financial Accounting",
    "FAC101T": "Financial Accounting I",
    "IBE109T": "Business Economics",
    "SEC111COM": "Skill Enhancement - Commerce",
    "AEC-101": "Developing Communicative Competence in English - I",
    "VEC-101-T": "Value Education Course",
    "VEC-102-T": "Indian Constitution",
    "IKS-101-T": "Indian Knowledge Systems (IKS)",
    "HIN101T": "Hindi Kavya Tatha Kahani Sahitya",
    "HIN102P": "Kavya Tatha Kahani Vachan Aur Prastuti",
    "MAR101T": "Sahityache Swarup (Nivedanatma Sahityaprakar)",
    "MAR102P": "Sahityaprakaranche Sadarikaran",
    "POL101T": "Introduction to Political Science - I",
    "POL102P": "Political Science Practical - I",
    "OE-102-BOT-T": "Plants and Human Welfare",
    "SEC-104-HIN": "Bhasha Shikshan Kaushalya"
}

def extract_paper_list_mapping_from_text(raw_text: str) -> dict:
    mapping = {}
    lines = raw_text.split('\n')
    is_paper_list = False

    for line in lines:
        line_clean = line.strip()
        if "Paper List" in line_clean:
            is_paper_list = True
            continue
        if is_paper_list:
            if "SEAT NO" in line_clean.upper() or "PRN:" in line_clean.upper():
                is_paper_list = False
                continue
            m = re.match(r'^([A-Z0-9_\-\(\)\s]{2,18})\s{2,}([A-Za-z0-9\s\-_:,\.\(\)\/\+\&]+)$', line_clean)
            if m:
                c_raw = m.group(1).strip()
                t_raw = m.group(2).strip()
                if any(w in t_raw.lower() for w in ["continued", "page", "savitribai"]):
                    continue
                norm = re.sub(r'[\s\-]+', '', c_raw).upper()
                mapping[norm] = t_raw.title()
                mapping[c_raw.upper()] = t_raw.title()
    return mapping

def lookup_subject_title(code_raw: str, dynamic_map: dict) -> str:
    clean = re.sub(r'[\s\-]+', '', code_raw).upper()
    if clean in dynamic_map:
        return dynamic_map[clean]
    if code_raw.strip().upper() in dynamic_map:
        return dynamic_map[code_raw.strip().upper()]
    for k, v in dynamic_map.items():
        if k == clean or k.startswith(clean) or clean.startswith(k):
            return v
    if clean in BUILTIN_TITLES:
        return BUILTIN_TITLES[clean]
    if code_raw.strip().upper() in BUILTIN_TITLES:
        return BUILTIN_TITLES[code_raw.strip().upper()]
    return ""

def format_subject_display(code_raw: str, dynamic_map: dict) -> str:
    title = lookup_subject_title(code_raw, dynamic_map)
    clean_code = code_raw.strip().upper()
    if title and len(title) >= 3:
        return f"{clean_code} ({title})"
    return clean_code

# ----------------- PARSER ENGINE -----------------
EXCLUDED = {
    "MIN", "MAX", "INT", "EXT", "TOT", "CRD", "ERN", "GRD", "PNT", "PAGE", 
    "SEAT", "NAME", "MOTHER", "SAVITRIBAI", "UNIVERSITY", "COLLEGE", "CENTRE",
    "PUNE", "RESULT", "BRANCH", "COURSE", "SESSION", "JUL", "JUNE", "APRIL", 
    "MAY", "OCT", "NOV", "SGPA", "CGPA", "TOTAL", "CREDIT", "SEM", "DATE"
}

VALID_GRADES = {"O", "A+", "A", "B+", "B", "C", "D", "P", "F", "FFF", "AAA", "AB", "FAIL", "PASS"}

def clean_marks(val):
    raw = str(val).strip()
    if any(k in raw.upper() for k in ["AAA", "AB", "ABSENT", "NIL", "---", "--"]):
        return 0.0, "AAA"
    if "FFF" in raw.upper():
        return 0.0, "F"
    clean = re.sub(r'[*$#&^!:]', '', raw).strip()
    if "/" in clean:
        clean = clean.split("/")[0].strip()
    try:
        return float(clean), None
    except ValueError:
        return 0.0, None

def extract_metadata(text):
    meta = {
        "institution": "ANNASAHEB WAGHIRE COLLEGE, OTUR",
        "puncode": "CAAP010180",
        "college_code": "0234",
        "faculty": "Faculty of Arts / Science / Commerce",
        "exam_session": "Exam Session 2026"
    }
    
    pun_match = re.search(r'\b(CA[A-Z0-9]{8,10})\b', text, re.IGNORECASE)
    if pun_match:
        meta["puncode"] = pun_match.group(1).upper()
        
    inst_match = re.search(r'(?:COLLEGE|CENTRE|INSTITUTE)\s*:\s*\[?([0-9A-Za-z]+)\]?\s*([A-Za-z0-9\s,.-]+)', text, re.IGNORECASE)
    if inst_match:
        meta["college_code"] = inst_match.group(1).strip()
        meta["institution"] = inst_match.group(2).split("\n")[0].strip()

    fac_match = re.search(r
