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
    count = 0
    if os.path.exists(VISITOR_FILE):
        try:
            with open(VISITOR_FILE, "r") as f:
                content = f.read().strip()
                count = int(content) if content.isdigit() else 0
        except Exception:
            count = 0
            
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

    fac_match = re.search(r'(?:FACULTY\s+OF\s+[A-Za-z\s&]+|RESULT\s+OF\s+[A-Za-z0-9\s().&,-]+|BACHELOR\s+OF\s+[A-Za-z0-9\s().&,-]+)', text, re.IGNORECASE)
    if fac_match:
        meta["faculty"] = fac_match.group(0).replace("RESULT OF", "").strip()

    ses_match = re.search(r'(OCT/NOV|APRIL/MAY|MARCH/APRIL|WINTER|SUMMER)\s*20\d{2}', text, re.IGNORECASE)
    if ses_match:
        meta["exam_session"] = ses_match.group(0).strip()

    return meta

def extract_text_from_file(file):
    ext = file.name.split('.')[-1].lower()
    full_text = ""
    if ext == "pdf" and pdfplumber:
        file.seek(0)
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                t = page.extract_text(layout=True) or page.extract_text()
                if t:
                    full_text += t + "\n"
    return full_text

def parse_ledger_text(raw_text, dynamic_paper_map):
    records = []
    current_student = None

    for line in raw_text.split('\n'):
        line_clean = line.strip()
        if not line_clean:
            continue

        parts = line_clean.split()
        if len(parts) >= 4 and parts[0].isdigit() and len(parts[0]) in (5, 6, 7):
            seat = parts[0]
            prn = next((p.replace(":", "").strip() for p in parts[1:] if len(p.replace(":", "").strip()) in (10, 11) and p.replace(":", "").strip()[:9].isdigit()), seat)
            sex_idx = next((i for i, p in enumerate(parts) if p in ('M', 'F') and i >= 2), -1)
            name = " ".join(parts[1:sex_idx - 1]) if sex_idx != -1 else " ".join(parts[1:3])
            current_student = {"Seat_No": seat, "PRN": prn, "Student_Name": name}
            continue

        if "SEAT" in line_clean.upper() and ("NAME" in line_clean.upper() or "PRN" in line_clean.upper()):
            s_m = re.search(r'(?:SEAT(?:\s*NO)?[\s.:]*)(\d{5,7})\b', line_clean, re.IGNORECASE)
            p_m = re.search(r'(?:PRN|P\.R\.N\.)[\s.:]*(\d{9,11}[A-Z]?)', line_clean, re.IGNORECASE)
            n_m = re.search(r'NAME(?:\s*OF\s*THE\s*CANDIDATE)?[\s.:]+([A-Za-z\s]+?)(?:\s+MOTHER|\s+PRN|$)', line_clean, re.IGNORECASE)
            if s_m:
                seat = s_m.group(1)
                prn = p_m.group(1) if p_m else seat
                name = n_m.group(1).strip() if n_m else "STUDENT"
                current_student = {"Seat_No": seat, "PRN": prn, "Student_Name": name}
                continue

        if current_student:
            m_subj = re.match(r'^([A-Z0-9_\-\(\)\s]{2,18}?)\s{2,}(.*(?:[OABCDEF][+]?|PASS|FAIL|FFF|AAA|AB).*)$', line_clean)
            if m_subj:
                code_raw = m_subj.group(1).strip()
                seg = m_subj.group(2).strip()
                code_norm = code_raw.replace(" ", "").upper()
                if code_norm in EXCLUDED or any(code_norm.startswith(ex) for ex in EXCLUDED):
                    continue

                tokens = seg.split()
                grade = None
                grade_idx = -1
                for idx in range(len(tokens) - 1, -1, -1):
                    t = re.sub(r'[*$#&^!:]', '', tokens[idx]).strip().upper()
                    if t in VALID_GRADES:
                        grade = "F" if t in ("FFF", "FAIL") else ("AAA" if t in ("AB", "AAA", "ABSENT") else ("P" if t == "PASS" else t))
                        grade_idx = idx
                        break

                if not grade:
                    continue

                valid_nums = []
                for tok in tokens[:grade_idx]:
                    if tok in ('---', '--', '-', ':', '!', '*', '$', '#', '&', '^'):
                        continue
                    m_val, _ = clean_marks(tok)
                    if tok.replace("*","").replace("$","").isdigit() or "/" in tok:
                        valid_nums.append(m_val)

                tot_marks = 0.0
                if valid_nums:
                    if len(valid_nums) >= 3 and valid_nums[-1] <= 10 and valid_nums[-2] <= 10:
                        tot_marks = valid_nums[-3]
                    elif len(valid_nums) >= 2 and valid_nums[-1] <= 4 and valid_nums[-2] > 4:
                        tot_marks = valid_nums[-2]
                    else:
                        tot_marks = valid_nums[-1]

                display_name = format_subject_display(code_raw, dynamic_paper_map)

                records.append({
                    "Seat_No": current_student["Seat_No"],
                    "PRN": current_student["PRN"],
                    "Student_Name": current_student["Student_Name"],
                    "Subject_Code": code_norm,
                    "Subject_Display": display_name,
                    "Total_Marks": tot_marks,
                    "Grade_or_Status": grade
                })

    return pd.DataFrame(records)

def parse_excel_ledger(file, dynamic_paper_map):
    file.seek(0)
    df_raw = pd.read_excel(file, header=None)
    meta = {
        "institution": str(df_raw.iloc[0, 0]).strip() if pd.notna(df_raw.iloc[0, 0]) else "ANNASAHEB WAGHIRE COLLEGE, OTUR",
        "puncode": "CAAP010180",
        "college_code": "0234",
        "faculty": str(df_raw.iloc[1, 0]).strip() if len(df_raw) > 1 and pd.notna(df_raw.iloc[1, 0]) else "Faculty Analysis",
        "exam_session": "Exam Session 2026"
    }

    hdr_idx = -1
    for idx, row in df_raw.iterrows():
        r_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        if "seat" in r_str and ("student" in r_str or "prn" in r_str):
            hdr_idx = idx
            break

    records = []
    if hdr_idx != -1:
        headers = [str(x).strip().replace('\n', ' ') for x in df_raw.iloc[hdr_idx]]
        data = df_raw.iloc[hdr_idx + 1:].copy()
        data.columns = headers
        data = data[data.iloc[:, 0].apply(lambda x: str(x).strip().isdigit() if pd.notna(x) else False)]

        seat_col = next((c for c in headers if "seat" in c.lower()), headers[1])
        prn_col = next((c for c in headers if "prn" in c.lower()), headers[2])
        name_col = next((c for c in headers if "student" in c.lower() or "name" in c.lower()), headers[3])
        subj_cols = [c for c in headers if "marks" in c.lower()]

        for _, row in data.iterrows():
            seat = str(row[seat_col]).strip()
            prn = str(row[prn_col]).strip() if pd.notna(row[prn_col]) else seat
            name = str(row[name_col]).strip()

            for m_col in subj_cols:
                raw_code = m_col.replace("Marks", "").replace("marks", "").strip().upper()
                g_col = next((c for c in headers if raw_code.lower() in c.lower() and "grade" in c.lower()), None)
                m_val, sp = clean_marks(row[m_col]) if pd.notna(row[m_col]) else (0.0, None)
                g_val = sp if sp else (str(row[g_col]).strip() if g_col and pd.notna(row[g_col]) else ("P" if m_val >= 20 else "F"))

                display_name = format_subject_display(raw_code, dynamic_paper_map)

                records.append({
                    "Seat_No": seat,
                    "PRN": prn,
                    "Student_Name": name,
                    "Subject_Code": raw_code,
                    "Subject_Display": display_name,
                    "Total_Marks": m_val,
                    "Grade_or_Status": g_val
                })

    return meta, pd.DataFrame(records)

# ----------------- ROBUST FILTERING & MATRIX BUILDER -----------------
def build_master_table(df_records, selected_subjs=None):
    if df_records.empty:
        return pd.DataFrame(), [], {}
        
    code_to_display = dict(zip(df_records['Subject_Code'], df_records['Subject_Display']))
    available_codes = list(df_records['Subject_Code'].unique())
    
    if selected_subjs:
        active_codes = []
        for s in selected_subjs:
            s_raw = s.strip()
            if not s_raw:
                continue
            s_norm = re.sub(r'[\s\-]+', '', s_raw).lower()
            s_lower = s_raw.lower()
            
            for code in available_codes:
                code_norm = re.sub(r'[\s\-]+', '', code).lower()
                disp_str = code_to_display.get(code, '').lower()
                disp_norm = re.sub(r'[\s\-]+', '', disp_str)
                
                # Robust match on normalized code, full title, or substring
                if (s_norm in code_norm or code_norm.startswith(s_norm) or 
                    s_lower in disp_str or s_norm in disp_norm):
                    if code not in active_codes:
                        active_codes.append(code)
    else:
        active_codes = available_codes

    # If no valid subjects match, return empty so user sees warning
    if not active_codes:
        return pd.DataFrame(), [], {}

    students = df_records[['Seat_No', 'PRN', 'Student_Name']].drop_duplicates(subset=['Seat_No']).reset_index(drop=True)
    students.insert(0, 'Sr. No.', range(1, len(students) + 1))

    wide = students.copy()
    for code in active_codes:
        disp = code_to_display.get(code, code)
        sub_df = df_records[df_records['Subject_Code'] == code].drop_duplicates(subset=['Seat_No']).set_index('Seat_No')
        wide[f"{disp} Marks (TOT)"] = wide['Seat_No'].map(sub_df['Total_Marks']).apply(lambda x: x if pd.notna(x) else "---")
        wide[f"{disp} Grade (GRD)"] = wide['Seat_No'].map(sub_df['Grade_or_Status']).apply(lambda x: x if pd.notna(x) else "---")
        
    # If specific filtering is requested, show only students who opted for at least one selected subject
    if selected_subjs:
        subj_grade_cols = [f"{code_to_display.get(c, c)} Grade (GRD)" for c in active_codes]
        opted_mask = wide[subj_grade_cols].apply(
            lambda row: any(str(g).strip() not in ('---', 'NOT OPTED', 'NAN', '') for g in row), 
            axis=1
        )
        wide = wide[opted_mask].copy().reset_index(drop=True)
        wide['Sr. No.'] = range(1, len(wide) + 1)

    return wide, active_codes, code_to_display

def get_toppers(wide_df, active_codes, code_to_display):
    rows = []
    for code in active_codes:
        disp = code_to_display.get(code, code)
        col_m = f"{disp} Marks (TOT)"
        col_g = f"{disp} Grade (GRD)"
        if col_m not in wide_df.columns:
            continue
        sub = wide_df[['Seat_No', 'PRN', 'Student_Name', col_m, col_g]].copy()
        sub['Marks_Num'] = pd.to_numeric(sub[col_m], errors='coerce').fillna(0)
        opted = sub[sub[col_g] != "---"]
        top3 = opted.sort_values(by="Marks_Num", ascending=False).head(3).reset_index(drop=True)
        for rank, r in top3.iterrows():
            rows.append({
                "Subject / Paper": disp,
                "Rank": f"Rank {rank + 1}",
                "Seat No": str(r["Seat_No"]),
                "PRN No": str(r["PRN"]),
                "Student Name": str(r["Student_Name"]),
                "Highest Marks (TOT)": float(r["Marks_Num"]),
                "Grade (GRD)": str(r[col_g])
            })
    return pd.DataFrame(rows)

def generate_excel(meta, wide_df, active_codes, code_to_display, subj_toppers, overall_toppers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master Marks & Summary"

    h_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    f_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    bdr = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

    num_cols = 4 + len(active_codes) * 2

    ws.cell(row=1, column=1, value=meta.get("institution", "").upper()).font = Font(name="Calibri", size=13, bold=True, color="1F497D")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    ws.cell(row=2, column=1, value=f"PunCode : {meta.get('puncode', 'N/A')}  |  College Code : {meta.get('college_code', 'N/A')}").font = Font(name="Calibri", size=10, bold=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    ws.cell(row=3, column=1, value=f"{meta.get('faculty', '')} - {meta.get('exam_session', '')}").font = Font(name="Calibri", size=10, italic=True)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=num_cols)

    base_headers = ["Sr. No.", "Seat No.", "PRN No.", "Student Name"]
    for c, h in enumerate(base_headers, 1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = h_fill

    c = 5
    for code in active_codes:
        disp = code_to_display.get(code, code)
        cell_m = ws.cell(row=5, column=c, value=f"{disp}\nMarks (TOT)")
        cell_m.font = Font(color="FFFFFF", bold=True)
        cell_m.fill = h_fill
        cell_g = ws.cell(row=5, column=c+1, value=f"{disp}\nGrade (GRD)")
        cell_g.font = Font(color="FFFFFF", bold=True)
        cell_g.fill = h_fill
        c += 2

    r_idx = 6
    for _, row in wide_df.iterrows():
        ws.cell(row=r_idx, column=1, value=row.get('Sr. No.', ''))
        ws.cell(row=r_idx, column=2, value=row.get('Seat_No', ''))
        ws.cell(row=r_idx, column=3, value=str(row.get('PRN', '')))
        ws.cell(row=r_idx, column=4, value=row.get('Student_Name', ''))
        c = 5
        for code in active_codes:
            disp = code_to_display.get(code, code)
            m = row.get(f"{disp} Marks (TOT)", "---")
            g = str(row.get(f"{disp} Grade (GRD)", "---")).strip()
            ws.cell(row=r_idx, column=c, value=m)
            c_g = ws.cell(row=r_idx, column=c+1, value=g)
            if g.upper() in ["F", "FAIL", "FFF"]:
                c_g.fill = f_fill
            c += 2
        for col_i in range(1, num_cols + 1):
            ws.cell(row=r_idx, column=col_i).border = bdr
        r_idx += 1

    ws.cell(row=r_idx, column=1, value="RESULT PASS %").font = Font(bold=True, color="1F497D")
    c = 5
    for code in active_codes:
        disp = code_to_display.get(code, code)
        g_col = wide_df[f"{disp} Grade (GRD)"].astype(str).str.upper()
        opted = (g_col != "---").sum()
        passed = g_col.isin(["O", "A+", "A", "B+", "B", "C", "D", "P"]).sum()
        pct = f"{round((passed / opted) * 100, 2)}%" if opted > 0 else "0%"
        ws.cell(row=r_idx, column=c, value=pct).alignment = Alignment(horizontal="center")
        c += 2

    if not subj_toppers.empty:
        ws_top = wb.create_sheet(title="Subject-Wise Toppers")
        for c_i, col_n in enumerate(subj_toppers.columns, 1):
            ws_top.cell(row=1, column=c_i, value=col_n).font = Font(color="FFFFFF", bold=True)
            ws_top.cell(row=1, column=c_i).fill = h_fill
        for r_i, (_, row) in enumerate(subj_toppers.iterrows(), 2):
            for c_i, val in enumerate(row.values, 1):
                ws_top.cell(row=r_i, column=c_i, value=val)

    if not overall_toppers.empty:
        ws_all = wb.create_sheet(title="Overall Toppers")
        for c_i, col_n in enumerate(overall_toppers.columns, 1):
            ws_all.cell(row=1, column=c_i, value=col_n).font = Font(color="FFFFFF", bold=True)
            ws_all.cell(row=1, column=c_i).fill = h_fill
        for r_i, (_, row) in enumerate(overall_toppers.iterrows(), 2):
            for c_i, val in enumerate(row.values, 1):
                ws_all.cell(row=r_i, column=c_i, value=val)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ----------------- SIDEBAR CONTROLS -----------------
st.sidebar.markdown("### 📁 Upload Ledger File")
uploaded_file = st.sidebar.file_uploader(
    "Choose PDF or Excel Result File",
    type=["pdf", "xlsx", "xls"]
)

st.sidebar.markdown(f"""
<div style="background:#eef2f7; padding:10px; border-radius:8px; text-align:center; margin-bottom:15px; border-left:4px solid #2a5298;">
    <span style="font-size:12px; color:#555;">Live Traffic / अभ्यागत</span><br>
    <b style="font-size:16px; color:#1e3c72;">👥 {total_visitors} Visitors</b>
</div>
""", unsafe_allow_html=True)

analysis_mode = "Complete Analysis (All Subjects)"
subject_filter_input = ""

if uploaded_file:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ⚙️ Analysis Filter")
    analysis_mode = st.sidebar.radio(
        "Analysis Scope:",
        ["Complete Analysis (All Subjects)", "Specific Subjects / Codes"]
    )
    if analysis_mode == "Specific Subjects / Codes":
        subject_filter_input = st.sidebar.text_area(
            "Enter Subject Code(s) or Name(s):",
            placeholder="e.g. HIN 101, MAR 101, 35151, ZOO-242"
        )
        st.sidebar.caption("ℹ️ Separate multiple codes/names with commas (,).")

# ----------------- DETAILED BILINGUAL INSTRUCTIONS -----------------
with st.expander("📖 User Guide & Instructions / पोर्टल कसे वापरावे (मार्गदर्शन)", expanded=False):
    t_en, t_mr = st.tabs(["English Instructions", "मराठी सूचना"])
    
    with t_en:
        st.markdown("""
        #### How to Use the Result Analyzer Portal:
        1. **Upload Result Ledger:** Click **Browse files** in the sidebar to upload your university result PDF or Excel copy.
        2. **Automatic Paper List Detection:** The system scans the first few pages of the PDF to extract the exact Paper Title for every course code.
        3. **Select Analysis Mode:**
           * **Complete Analysis (All Subjects):** Generates reports for all subjects with full Subject Name & Code.
           * **Specific Subjects / Codes:** Type desired subject codes or names separated by commas (e.g. `HIN 101, MAR 101, 35151`). Only these specific subjects will be displayed.
        4. **Generate Analysis:** Click the blue **🚀 Analyze Result & Generate Merit Report** button.
        5. **Review Results:** Check College PunCode, class pass percentage, subject-wise statistics, and top 3 scorers in each paper.
        6. **Download / Print:** Click **Download Master Analysis Excel** or press **Ctrl + P** to save as a clean PDF.
        """)
        
    with t_mr:
        st.markdown("""
        #### निकाल विश्लेषक पोर्टल वापरण्याची पद्धत:
        1. **निकाल पत्रक अपलोड करा:** डाव्या बाजूला दिलेल्या **Browse files** वर क्लिक करून निकाल पत्रक (PDF किंवा Excel) अपलोड करा.
        2. **विषयांची नावे आपोआप ओळखणे:** प्रणाली पहिल्या काही पानांवरून (Paper List) प्रत्येक कोडचे अचूक नाव आपोआप शोधून काढते.
        3. **विश्लेषण प्रकार निवडा:**
           * **Complete Analysis:** सर्व विषयांचे संपूर्ण निकाल विश्लेषण (विषयाचा कोड व अचूक नावासह) तयार होते.
           * **Specific Subjects / Codes:** ठरावीक विषयांचे विश्लेषण करण्यासाठी विषयांचे कोड किंवा नाव स्वल्पविरामाने (comma) वेगळे करून लिहा (उदा. `HIN 101, MAR 101, 35151`). फक्त तेच विषय अहवालात दिसतील.
        4. **विश्लेषण सुरू करा:** **🚀 Analyze Result & Generate Merit Report** या बटणावर क्लिक करा.
        5. **अहवाल तपासा:** महाविद्यालय PunCode, एकूण निकाल %, विषयानुसार निकाल टक्केवारी आणि प्रत्येक विषयातील प्रथम ३ गुणवंत विद्यार्थी पहा.
        6. **डाउनलोड व प्रिंट:** **Download Master Analysis Excel** बटणावर क्लिक करून एक्सेल फाइल मिळवा किंवा **Ctrl + P** दाबून अहवाल PDF स्वरूपात सेव्ह करा.
        """)

# Bilingual AI Assistant
st.sidebar.markdown("---")
st.sidebar.markdown("### 🤖 Assistant / सहाय्यक")
ai_lang = st.sidebar.radio("Language / भाषा:", ["English", "मराठी"], horizontal=True)
with st.sidebar.expander("💬 Ask Help / मदत", expanded=False):
    q = st.text_input("Ask question:" if ai_lang == "English" else "प्रश्न विचारा:")
    if q:
        if ai_lang == "English":
            st.info("💡 Upload your result ledger file. Use 'Specific Subjects' to filter departmental papers or 'Complete Analysis' to evaluate all students.")
        else:
            st.info("💡 निकाल पत्रक अपलोड करा. आपल्या विभागातील विषय पाहण्यासाठी 'Specific Subjects' चा वापर करा किंवा संपूर्ण वर्ग निकाल तपासा.")

# ----------------- MAIN SCREEN -----------------
if uploaded_file:
    file_ext = uploaded_file.name.split('.')[-1].lower()

    with st.expander("👁️ View Uploaded Document Preview", expanded=True):
        if file_ext == "pdf":
            uploaded_file.seek(0)
            b64 = base64.b64encode(uploaded_file.read()).decode('utf-8')
            st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="400" type="application/pdf"></iframe>', unsafe_allow_html=True)
            uploaded_file.seek(0)
        else:
            df_preview = pd.read_excel(uploaded_file)
            st.dataframe(df_preview.head(8), use_container_width=True)
            uploaded_file.seek(0)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Analyze Result & Generate Merit Report", type="primary", use_container_width=True):
        st.session_state['run_analysis'] = True

    if st.session_state.get('run_analysis', False):
        with st.spinner("Scanning Paper List from first pages, mapping subjects, and extracting marks..."):
            if file_ext in ["xlsx", "xls"]:
                dynamic_map = {}
                metadata, df_records = parse_excel_ledger(uploaded_file, dynamic_map)
            else:
                raw_text = extract_text_from_file(uploaded_file)
                dynamic_map = extract_paper_list_mapping_from_text(raw_text)
                metadata = extract_metadata(raw_text)
                df_records = parse_ledger_text(raw_text, dynamic_map)

        st.markdown(f"""
        <div class="inst-card">
            <h3 style="margin:0; color:#1e3c72; font-weight:700;">College : {metadata.get('institution', 'ANNASAHEB WAGHIRE COLLEGE, OTUR')}</h3>
            <p style="margin:6px 0 0 0; font-size:14px; color:#495057;">
                <b>PunCode :</b> {metadata.get('puncode', 'CAAP010180')} &nbsp;|&nbsp; 
                <b>College Code :</b> {metadata.get('college_code', '0234')} &nbsp;|&nbsp;
                <b>Faculty :</b> {metadata.get('faculty', 'Faculty Analysis')} &nbsp;|&nbsp;
                <b>Session :</b> {metadata.get('exam_session', '2026')}
            </p>
        </div>
        """, unsafe_allow_html=True)

        if df_records.empty:
            st.error("⚠️ No student records could be detected. Please verify the uploaded file contains legible SPPU tabular student entries.")
        else:
            # Parse filter input
            if analysis_mode == "Specific Subjects / Codes" and subject_filter_input.strip():
                selected_list = [s.strip() for s in subject_filter_input.split(",") if s.strip()]
            else:
                selected_list = None

            master_sheet, active_codes, code_to_display = build_master_table(df_records, selected_list)

            if master_sheet.empty or not active_codes:
                st.warning(f"⚠️ No matching subjects found for '{subject_filter_input}'. Please check the subject codes or names.")
            else:
                scope_label = f"Specific Subjects ({len(active_codes)} Selected)" if selected_list else f"Complete Analysis ({len(active_codes)} Subjects)"
                st.markdown("---")
                st.markdown(f"### 📋 Master Marks Ledger — {scope_label}")
                st.dataframe(master_sheet, use_container_width=True, height=430)

                # Subject Wise Cards with Code & Full Name
                st.markdown("### 📈 Subject-Wise Summary (विषयानुसार निकाल)")
                cols = st.columns(min(len(active_codes), 4) if active_codes else 1)
                for idx, code in enumerate(active_codes):
                    disp = code_to_display.get(code, code)
                    col = cols[idx % len(cols)]
                    g_col = master_sheet[f"{disp} Grade (GRD)"].astype(str).str.upper()
                    opted = int((g_col != "---").sum())
                    passed = int(g_col.isin(["O", "A+", "A", "B+", "B", "C", "D", "P"]).sum())
                    failed = int(g_col.isin(["F", "FAIL", "FFF"]).sum())
                    pct = round((passed / opted) * 100, 2) if opted > 0 else 0.0
                    with col:
                        st.markdown(f"""
                        <div class="subject-box">
                            <strong style="color: #1e3c72; font-size:13px;">{disp}</strong><br>
                            <small>Opted: <b>{opted}</b> | Passed: <b>{passed}</b> | Failed: <b>{failed}</b></small><br>
                            <span style="font-size: 16px; color: {'#11998e' if pct >= 50 else '#eb3349'}; font-weight: 700;">{pct}% Pass</span>
                        </div>
                        """, unsafe_allow_html=True)

                # Overall Stats for evaluated subjects
                total_students = len(master_sheet)
                all_grades = [f"{code_to_display.get(c, c)} Grade (GRD)" for c in active_codes]
                student_pass = master_sheet[all_grades].apply(lambda row: all(str(g).upper() in ["O", "A+", "A", "B+", "B", "C", "D", "P", "---"] for g in row) and any(str(g).upper() in ["O", "A+", "A", "B+", "B", "C", "D", "P"] for g in row), axis=1)
                overall_pass = int(student_pass.sum())
                overall_pct = round((overall_pass / total_students) * 100, 2) if total_students > 0 else 0.0

                st.markdown("---")
                st.markdown("### 📊 Overall Class Performance / एकूण वर्ग निकाल")
                m1, m2, m3, m4 = st.columns(4)
                m1.markdown(f'<div class="metric-info"><small>TOTAL APPEARED (एकूण विद्यार्थी)</small><h2>{total_students}</h2></div>', unsafe_allow_html=True)
                m2.markdown(f'<div class="metric-pass"><small>PASSED ALL PAPERS (सर्व विषयात उत्तीर्ण)</small><h2>{overall_pass}</h2></div>', unsafe_allow_html=True)
                m3.markdown(f'<div class="metric-fail"><small>FAILED / ATKT (अनुत्तीर्ण / एटीकेटी)</small><h2>{total_students - overall_pass}</h2></div>', unsafe_allow_html=True)
                m4.markdown(f'<div class="metric-pass"><small>OVERALL PASS % (एकूण निकाल %)</small><h2>{overall_pct}%</h2></div>', unsafe_allow_html=True)

                # Subject-Wise Toppers
                st.markdown("---")
                st.markdown("### 🏆 Top 3 Toppers in Respective Subjects / विषयानुसार प्रथम ३ गुणवंत विद्यार्थी")
                subj_toppers_df = get_toppers(master_sheet, active_codes, code_to_display)
                if not subj_toppers_df.empty:
                    st.dataframe(subj_toppers_df, use_container_width=True)

                # Overall Toppers across evaluated subjects
                st.markdown("### 🎖️ Overall Merit Rankers / सर्वसाधारण प्रथम ३ गुणवंत")
                all_marks_cols = [f"{code_to_display.get(c, c)} Marks (TOT)" for c in active_codes]
                numeric_marks = master_sheet[all_marks_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
                master_sheet['Grand_Total'] = numeric_marks.sum(axis=1)
                overall_top = master_sheet.sort_values(by="Grand_Total", ascending=False).head(3).reset_index(drop=True)
                overall_top_df = pd.DataFrame({
                    "Rank": [f"Rank {i+1}" for i in range(len(overall_top))],
                    "Seat No": overall_top['Seat_No'],
                    "PRN No": overall_top['PRN'],
                    "Student Name": overall_top['Student_Name'],
                    "Total Marks": overall_top['Grand_Total']
                })
                st.table(overall_top_df)

                # Export Excel
                st.markdown("---")
                excel_bytes = generate_excel(metadata, master_sheet, active_codes, code_to_display, subj_toppers_df, overall_top_df)
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    st.download_button(
                        label="📗 Download Master Analysis Excel (.xlsx)",
                        data=excel_bytes,
                        file_name="Master_Result_Analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with col_b2:
                    st.info("💡 **Save as PDF:** Press **Ctrl + P** on your keyboard to instantly print or save this report.")

else:
    st.info("👈 Please upload your result ledger from the sidebar to begin analysis.")

# Departmental Copyright Notice
st.markdown("""
<div class="footer-box">
    © Developed & Maintained by <b>Dr. Rahul Patil</b>, Department of Zoology, 
    PDAs Annasaheb Waghire College, Otur, Junnar, Dist-Pune.
</div>
""", unsafe_allow_html=True)
