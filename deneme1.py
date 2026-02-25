import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from openpyxl import load_workbook

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Bakım Paketi Oluşturucu", layout="wide")
st.title("✈️ Bakım Paketi Excel Oluşturucu (Production)")

pdf_file = st.file_uploader("PDF Yükle", type=["pdf"])
template_file = st.file_uploader("Excel Şablon Yükle", type=["xlsx"])
wo_number = st.text_input("W/O Numarası")

use_engineering = st.checkbox("Mühendislik değerlendirmesi kullan (opsiyonel)")
map_file = None
if use_engineering:
    map_file = st.file_uploader(
        "Mühendislik Değerlendirmesi Excel Yükle (DESCRIPTION + CMT/IMT/CDCCL/KOMPLEKS)",
        type=["xlsx"]
    )

# -----------------------------
# Session state (download persistence)
# -----------------------------
if "filled_xlsx" not in st.session_state:
    st.session_state["filled_xlsx"] = None
if "filled_tsv" not in st.session_state:
    st.session_state["filled_tsv"] = None
if "dl_aircraft" not in st.session_state:
    st.session_state["dl_aircraft"] = ""
if "dl_v" not in st.session_state:
    st.session_state["dl_v"] = 0

# -----------------------------
# Helpers
# -----------------------------
def get_location_from_package(package_name: str) -> str:
    package_name = (package_name or "").strip().upper()
    return package_name[-3:] if len(package_name) >= 3 else ""

def normalize_skill(skill_value):
    s = str(skill_value).strip().upper()
    return s if s in ("B1", "B2") else "B1"

def norm_header(s) -> str:
    return str(s).strip().lower() if s is not None else ""

def clean_text_key(text) -> str:
    if text is None:
        return ""
    s = str(text).upper().replace("İ", "I")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def clean_description(text) -> str:
    return clean_text_key(text)

def yn_from_any(val) -> str:
    if val is None:
        return "N"
    s = str(val).strip().upper()
    return "Y" if s in ("Y", "YES", "TRUE", "1", "T") else "N"

def parse_mh_and_skill(value):
    """
    Est. MH örnekleri:
      - 03:00/B1 -> ("3","B1")
      - 00:45/B1 -> ("0","B1")  (sonra 0->1)
      - 3/B1     -> ("3","B1")
      - 03:00    -> ("3","")
    """
    if value is None:
        return "", ""
    v = str(value).strip()
    if not v:
        return "", ""

    skill = ""
    mh_part = v

    if "/" in v:
        mh_part, skill = v.split("/", 1)
        mh_part = mh_part.strip()
        skill = normalize_skill(skill)

    if ":" in mh_part:
        try:
            h = int(float(mh_part.split(":")[0]))
            return str(h), skill
        except Exception:
            return mh_part, skill

    try:
        return str(int(float(mh_part))), skill
    except Exception:
        return mh_part, skill

def safe_cell_str(v) -> str:
    if v is None:
        return ""
    return str(v).replace("\t", " ").replace("\n", " ").replace("\r", " ")

# -----------------------------
# Cover info
# -----------------------------
def extract_cover_info(full_text: str):
    """Paket: Type Of Work, Tescil: A/C Type / Registration"""
    package_name = ""
    aircraft = ""

    m_type = re.search(r"Type\s*Of\s*Work\s*:?\s*(.+)", full_text, re.IGNORECASE)
    if m_type:
        package_name = m_type.group(1).strip()

    m_reg = re.search(r"A/C Type\s*/\s*Registration\s*(.+)", full_text, re.IGNORECASE)
    if m_reg:
        aircraft = m_reg.group(1).split("/")[-1].strip()

    return aircraft, package_name

def detect_aircraft_family_from_cover(pdf_bytes: bytes):
    """
    Cover Page'deki 'A/C Type / Registration' alanından ilk 4 haneye göre
    B737NG (B73N) / B737MAX (B73M) tespiti yapar.
    """
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    ac_match = re.search(r"A/C Type\s*/\s*Registration\s*(.+)", full_text, re.IGNORECASE)
    if not ac_match:
        return None, "⚠️ A/C Type / Registration bulunamadı."

    ac_info = ac_match.group(1).strip()
    prefix4 = ac_info[:4].upper()

    if prefix4 == "B73N":
        return "B737NG", "✈️ Uçak tipi: B737NG (B73N)"
    if prefix4 == "B73M":
        return "B737MAX", "✈️ Uçak tipi: B737MAX (B73M) ‼️YETKİ KONTROLÜ YAPILMASI GEREKİYOR‼️"

    return "UNKNOWN", f"⚠️ Uçak tipi tanınamadı (ilk 4 hane: {prefix4})"

# -----------------------------
# Summary table detection (robust)
# -----------------------------
def is_summary_page(page_text: str) -> bool:
    return "SUMMARY" in (page_text or "").upper()

def normalize_colname(c) -> str:
    return str(c).strip().upper() if c is not None else ""

def find_best_columns(df_cols):
    """Esnek kolon bulma: Description, Est.MH, W/O & Reference"""
    desc_col = next((c for c in df_cols if "DESC" in normalize_colname(c)), None)
    mh_col = next((c for c in df_cols if "MH" in normalize_colname(c)), None)

    ref_col = next((c for c in df_cols if ("W/O" in normalize_colname(c) and "REFER" in normalize_colname(c))), None)
    if ref_col is None:
        ref_col = next((c for c in df_cols if "REFER" in normalize_colname(c)), None)
    if ref_col is None:
        ref_col = next((c for c in df_cols if ("W/O" in normalize_colname(c) or "WO" in normalize_colname(c))), None)

    return desc_col, mh_col, ref_col

def table_looks_like_summary(header_row) -> bool:
    """Summary tablosunu ayır: DESC + MH + (REFER veya W/O)"""
    header = [normalize_colname(h) for h in header_row]
    has_desc = any("DESC" in h for h in header)
    has_mh = any("MH" in h for h in header)
    has_ref = (
        any("REFER" in h for h in header)
        or any("W/O" in h for h in header)
        or any("WO" in h.replace(" ", "") for h in header)
    )
    return has_desc and has_mh and has_ref

def extract_summary_tasks(pdf_bytes: bytes):
    """
    PDF’ten tasks:
    - final description kuralı:
        * İlk 20 karakterde '-' varsa: CAMO_PREFIX + raw_desc
        * Yoksa: CAMO_PREFIX + WO:xxxx  (raw_desc eklenmez)
          (WO bulunamazsa: CAMO_PREFIX + raw_desc fallback)
    - match_key: raw_desc (engineering eşleşme için)
    """
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                full_text += t + "\n"

    aircraft, package_name = extract_cover_info(full_text)

    camo_prefix = f"PLEASE PERFORM CAMO WP: {package_name} | "
    camo_prefix = camo_prefix.upper().replace("İ", "I")

    tasks = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if not is_summary_page(page_text):
                continue

            tables = page.extract_tables() or []
            for table in tables:
                if not table or len(table) < 2:
                    continue

                header = table[0]
                if not table_looks_like_summary(header):
                    continue

                df = pd.DataFrame(table[1:], columns=header)
                desc_col, mh_col, ref_col = find_best_columns(df.columns)

                if desc_col is None or mh_col is None:
                    continue

                for _, row in df.iterrows():
                    raw_desc = clean_description(row.get(desc_col, ""))
                    if not raw_desc or raw_desc.lower() == "none":
                        continue

                    mh, skill = parse_mh_and_skill(row.get(mh_col, ""))
                    if mh in ("", "0"):
                        mh = "1"

                    needs_wo = "-" not in raw_desc[:20]

                    wo_prefix = ""
                    if needs_wo and ref_col is not None:
                        ref_val = str(row.get(ref_col, "") or "")
                        m = re.search(r"\b\d{3,}\b", ref_val)
                        if m:
                            wo_prefix = f"WO:{m.group(0)} "

                    if needs_wo and wo_prefix:
                        final_desc = camo_prefix + wo_prefix
                    elif needs_wo and not wo_prefix:
                        final_desc = camo_prefix + raw_desc
                    else:
                        final_desc = camo_prefix + raw_desc

                    tasks.append({
                        "description": final_desc,
                        "match_key": raw_desc,
                        "man_hour": mh,
                        "skill": skill,
                        "rII": "N",
                        "critical_task": "N",
                        "cdccl": "N",
                    })

    return aircraft, package_name, tasks

# -----------------------------
# Engineering mapping
# -----------------------------
def load_engineering_mapping(uploaded_excel):
    df = pd.read_excel(uploaded_excel)
    cols = {str(c).strip().upper(): c for c in df.columns}

    desc_col_key = next((k for k in cols.keys() if k.strip().upper() == "DESCRIPTION"), None)
    if not desc_col_key:
        raise KeyError("Mühendislik değerlendirmesi Excel'inde DESCRIPTION sütunu yok.")

    desc_col = cols[desc_col_key]
    cmt_col = cols.get("CMT")
    imt_col = cols.get("IMT")
    cdccl_col = cols.get("CDCCL")
    kompleks_col = cols.get("KOMPLEKS")

    mapping = {}
    kompleks_any = False

    for _, r in df.iterrows():
        key = clean_text_key(r.get(desc_col))
        if not key:
            continue

        cmt = yn_from_any(r.get(cmt_col)) if cmt_col else "N"
        imt = yn_from_any(r.get(imt_col)) if imt_col else "N"
        cdccl = yn_from_any(r.get(cdccl_col)) if cdccl_col else "N"
        kompleks = yn_from_any(r.get(kompleks_col)) if kompleks_col else "N"

        if kompleks == "Y":
            kompleks_any = True

        prev = mapping.get(key, {"cmt": "N", "imt": "N", "cdccl": "N"})
        mapping[key] = {
            "cmt": "Y" if (prev["cmt"] == "Y" or cmt == "Y") else "N",
            "imt": "Y" if (prev["imt"] == "Y" or imt == "Y") else "N",
            "cdccl": "Y" if (prev["cdccl"] == "Y" or cdccl == "Y") else "N",
        }

    return mapping, kompleks_any

# -----------------------------
# Fill template
# -----------------------------
def fill_template_excel(template_bytes, aircraft, package_name, tasks, wo_number):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    header_map = {norm_header(c.value): c.column for c in ws[1] if c.value}

    def col(name: str):
        return header_map.get(norm_header(name))

    required = [
        "Aircraf", "Check", "wo", "chapter", "sectIon", "task_card_descrIptIon",
        "addItIon_work", "edItor_used", "source_code", "rII", "cdccl",
        "crItIcal_task", "etops", "mechanIc", "skIll", "man_hours", "men_requIred"
    ]
    missing = [k for k in required if col(k) is None]
    if missing:
        raise KeyError(f"Şablon Excel'de şu başlıklar eksik: {missing}")

    start_row = 2
    for i, t in enumerate(tasks):
        r = start_row + i

        ws.cell(r, col("Aircraf")).value = aircraft
        ws.cell(r, col("Check")).value = package_name
        ws.cell(r, col("wo")).value = wo_number
        ws.cell(r, col("chapter")).value = 5
        ws.cell(r, col("sectIon")).value = 0
        ws.cell(r, col("task_card_descrIptIon")).value = t["description"]
        ws.cell(r, col("addItIon_work")).value = "YES"
        ws.cell(r, col("edItor_used")).value = "STYLESHEET"
        ws.cell(r, col("source_code")).value = "R"

        ws.cell(r, col("rII")).value = t["rII"]
        ws.cell(r, col("cdccl")).value = t["cdccl"]
        ws.cell(r, col("crItIcal_task")).value = t["critical_task"]
        ws.cell(r, col("etops")).value = "N"

        ws.cell(r, col("mechanIc")).value = "Y"
        ws.cell(r, col("skIll")).value = t["skill"]
        ws.cell(r, col("man_hours")).value = t["man_hour"]
        ws.cell(r, col("men_requIred")).value = 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# -----------------------------
# Export TSV from filled workbook (includes hidden columns)
# -----------------------------
def workbook_bytes_to_tsv_bytes(xlsx_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    headers = [safe_cell_str(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
    lines = ["\t".join(headers)]

    for r in range(2, ws.max_row + 1):
        row_vals = []
        for c in range(1, ws.max_column + 1):
            row_vals.append(safe_cell_str(ws.cell(r, c).value))
        lines.append("\t".join(row_vals))

    return ("\n".join(lines)).encode("utf-8")

# -----------------------------
# Boeing interval control (ONLY NG)
# -----------------------------
def _to_int_num(s: str) -> int:
    s = (s or "").strip().replace(",", "")
    try:
        return int(float(s))
    except Exception:
        return 0

def scan_threshold_repeat_values(block_text: str):
    """
    Kart bloğunda THRESHOLD/REPEAT satırı ve sonraki 3 satırdan
    FH/FC/YR değerlerini toplar. (max değerleri döner)
    """
    if not block_text:
        return {"FH": 0, "FC": 0, "YR": 0, "found": []}

    lines = [ln.strip() for ln in block_text.splitlines() if ln.strip()]
    nu = re.compile(r"\b(\d{1,3}(?:,\d{3})|\d+)\s(FH|FC|YR)\b", re.IGNORECASE)

    found = []  # (value_str, unit, where)
    maxv = {"FH": 0, "FC": 0, "YR": 0}

    for i, ln in enumerate(lines):
        up = ln.upper()
        if ("THRESHOLD" in up) or ("REPEAT" in up):
            where = "THRESHOLD" if "THRESHOLD" in up else "REPEAT"
            chunk = " ".join(lines[i:i+4]).upper()

            for m in nu.finditer(chunk):
                v_str = m.group(1)
                unit = m.group(2).upper()
                v = _to_int_num(v_str)
                found.append((v_str, unit, where))
                if unit in maxv:
                    maxv[unit] = max(maxv[unit], v)

    return {"FH": maxv["FH"], "FC": maxv["FC"], "YR": maxv["YR"], "found": found}

def extract_boeing_interval_exceedances(pdf_bytes: bytes, debug: bool = False):
    """
    Kart bazında FH/FC/YR max değerlerini çıkarır.
    Limit aşan birim varsa (hangisi aşıyorsa), öncelik: FH>FC>YR (aşanlar içinde).
    """
    exceed = []
    debug_rows = []

    marker_re = re.compile(r"BOEING\s+CARD\s+NO\.?", re.IGNORECASE)

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            if not text:
                continue

            splits = marker_re.split(text)
            if len(splits) <= 1:
                continue

            for blk in splits[1:]:
                block = "BOEING CARD NO.\n" + blk

                # Kart adı: marker'dan sonraki ilk dolu satır
                card_name = "(CARD NAME NOT FOUND)"
                after_lines = block.splitlines()[1:12]
                for ln in after_lines:
                    if ln.strip():
                        card_name = ln.strip()
                        break

                vals = scan_threshold_repeat_values(block)

                over_fh = vals["FH"] > 7500
                over_fc = vals["FC"] > 4000
                over_yr = vals["YR"] > 3

                show_unit = None
                show_value = 0
                show_limit = 0

                if over_fh:
                    show_unit, show_value, show_limit = "FH", vals["FH"], 7500
                elif over_fc:
                    show_unit, show_value, show_limit = "FC", vals["FC"], 4000
                elif over_yr:
                    show_unit, show_value, show_limit = "YR", vals["YR"], 3

                if show_unit:
                    exceed.append({
                        "Page": pno,
                        "Card": card_name,
                        "Unit": show_unit,
                        "Value": show_value,
                        "Limit": show_limit,
                        "FH_max": vals["FH"],
                        "FC_max": vals["FC"],
                        "YR_max": vals["YR"],
                    })

                if debug:
                    debug_rows.append({
                        "Page": pno,
                        "Card": card_name,
                        "FH_max": vals["FH"],
                        "FC_max": vals["FC"],
                        "YR_max": vals["YR"],
                        "FoundCount": len(vals["found"]),
                        "FoundSample": ", ".join([f"{v} {u} ({w})" for v, u, w in vals["found"][:6]])
                    })

    return exceed, debug_rows

# -----------------------------
# Main action: compute & store
# -----------------------------
if st.button("Excel Oluştur"):
    if not (pdf_file and template_file and wo_number):
        st.error("PDF, Excel şablon ve W/O numarasını girmen gerekiyor.")
    elif use_engineering and map_file is None:
        st.error("Mühendislik değerlendirmesi seçildi ama Excel yüklenmedi.")
    else:
        try:
            pdf_bytes = pdf_file.getvalue()

            family, msg = detect_aircraft_family_from_cover(pdf_bytes)
            if family == "B737MAX":
                st.info(msg)
            elif family == "B737NG":
                st.success(msg)
            else:
                st.warning(msg)

            aircraft, package_name, tasks = extract_summary_tasks(pdf_bytes)

            # Lokasyon + AYT uyarısı
            location = get_location_from_package(package_name)
            target = "38-070-00-01"
            has_target = any((t.get("match_key", "") or "")[:12].upper() == target for t in tasks)
            if location == "AYT" and has_target:
                st.warning("WATER DISINFECTION KARTI TOOL SORUNU VAR")

            # Boeing interval kontrolü (SADECE NG)
            if family == "B737NG":
                st.subheader("Boeing Task Card Interval Kontrolü (B737NG)")
                exceed, debug_rows = extract_boeing_interval_exceedances(pdf_bytes, debug=True)

                df_ex = pd.DataFrame(exceed)
                if not df_ex.empty:
                    st.warning(f"⚠️ Limit aşan kart sayısı: {len(df_ex)}")
                    st.dataframe(df_ex, use_container_width=True)
                else:
                    st.info("✅ Interval limit aşımı bulunmadı.")

                with st.expander("Interval Debug (kart başına yakalanan değerler)"):
                    st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)
            else:
                st.info("Boeing interval kontrolü yalnızca B737NG için çalışır.")

            # Engineering mapping (optional)
            if use_engineering and map_file is not None:
                mapping, kompleks_any = load_engineering_mapping(map_file)

                matched = 0
                for t in tasks:
                    k = clean_text_key(t["match_key"])
                    if k in mapping:
                        matched += 1
                        t["rII"] = mapping[k]["cmt"]
                        t["critical_task"] = mapping[k]["imt"]
                        t["cdccl"] = mapping[k]["cdccl"]
                    else:
                        t["rII"] = "N"
                        t["critical_task"] = "N"
                        t["cdccl"] = "N"

                st.info(f"Eşleştirme: {matched} eşleşti | {len(tasks) - matched} eşleşmedi")
                if kompleks_any:
                    st.warning("⚠️ Kompleks iş var (KOMPLEKS=Y/YES bulundu)")

            # Summary metrics
            total_mh = 0
            for t in tasks:
                try:
                    total_mh += int(str(t["man_hour"]).strip())
                except Exception:
                    pass

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Toplam İş", len(tasks))
            c2.metric("Toplam Man Hour", total_mh)
            c3.metric("rII=Y", sum(1 for t in tasks if t["rII"] == "Y"))
            c4.metric("Critical=Y", sum(1 for t in tasks if t["critical_task"] == "Y"))
            c5.metric("Lokasyon", location or "-")

            if len(tasks) == 0:
                st.error("Summary tablosundan hiç iş çekilemedi.")
            else:
                filled_xlsx = fill_template_excel(
                    template_bytes=template_file.read(),
                    aircraft=aircraft,
                    package_name=package_name,
                    tasks=tasks,
                    wo_number=wo_number
                )
                tsv_bytes = workbook_bytes_to_tsv_bytes(filled_xlsx)

                st.session_state["filled_xlsx"] = filled_xlsx
                st.session_state["filled_tsv"] = tsv_bytes
                st.session_state["dl_aircraft"] = aircraft
                st.session_state["dl_v"] += 1

                st.success("Dosyalar hazır. Aşağıdan indirebilirsin ✅")

        except Exception as e:
            st.error(f"Hata: {e}")

# -----------------------------
# Persistent download buttons
# -----------------------------
if st.session_state["filled_xlsx"] is not None:
    aircraft = st.session_state["dl_aircraft"] or "IMPORT"
    v = st.session_state["dl_v"]

    colA, colB = st.columns(2)
    with colA:
        st.download_button(
            label="IMPORT Excel (.xlsx)",
            data=st.session_state["filled_xlsx"],
            file_name=f"{aircraft} IMPORT EXCELI_v{v}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_xlsx_persist_{v}",
        )
    with colB:
        st.download_button(
            label="IMPORT Text (Tab Delimited .txt)",
            data=st.session_state["filled_tsv"],
            file_name=f"{aircraft} IMPORT EXCELI_v{v}.txt",
            mime="text/plain",
            key=f"dl_txt_persist_{v}",
        )
