# -*- coding: utf-8 -*-
"""
=======================================================================================
 CAFFE BENE — Өдрийн тооллого, Борлуулалтын систем тулгалтын веб апп
=======================================================================================
Ашиглах сангууд:
    pip install streamlit pandas numpy fuzzywuzzy python-Levenshtein openpyxl

Ажиллуулах:
    streamlit run app.py
=======================================================================================

Энэ хувилбарт орсон гол засвар / сайжруулалтууд:
  1. clean_code: тэргүүн тэгтэй код (жишээ нь "007") устгагдахгүй болсон.
  2. parse_system_excel: Нэр багана олдоогүй тохиолдолд бүх мөрийг санамсаргүй
     хаяхгүй болсон (зөвхөн Нэр багана байгаа үед л хоосон нэрийг шүүнэ).
  3. Хайлтын талбарууд (search) regex мэт тэмдэгтээс болж crash хийхгүй болсон
     (regex=False).
  4. pandas-ийн шинэ хувилбарт .applymap deprecated тул .map руу шилжиж,
     хуучин хувилбартай нийцтэй байлгах fallback нэмэгдсэн.
  5. JSON файл бичихдээ түр файл руу бичээд атомаар сольж хадгалдаг болсон
     (os.replace) — апп гэнэт унасан ч өгөгдөл эвдрэхгүй.
  6. Устгах/цэвэрлэх мэт эргэлт буцалтгүй үйлдлүүдэд 2 дахь баталгаажуулах
     алхам (checkbox) нэмэгдсэн.
  7. Тулгалтын үр дүнд "Тоолоогүй ч Системд зарагдсан" бараануудыг тусад нь
     харуулдаг болсон.
  8. Код баганыг чөлөөт бичих Text талбар болгож, мастер жагсаалтаас
     автоматаар нэр бөглөх санал (autofill) хэвээр үлдсэн — гэхдээ
     сонголтоор хязгаарлагдахгүй.
  9. Fuzzy тулгалтын босго (threshold) sidebar-аас тохируулах боломжтой болсон.
  10. load_master/load_json-д хөнгөн кэш нэмэгдсэн, хадгалах бүрд цэвэрлэдэг.
=======================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import io
import re
import uuid
import tempfile
import base64
from datetime import datetime, date

try:
    from fuzzywuzzy import fuzz, process
except ImportError:
    st.error("fuzzywuzzy сан суугаагүй байна. Терминал дээр: pip install fuzzywuzzy python-Levenshtein")
    st.stop()


# =======================================================================================
# 0. ЕРӨНХИЙ ТОХИРГОО
# =======================================================================================
st.set_page_config(
    page_title="Caffe Bene | Тооллого & Тулгалт",
    page_icon="☕",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = "cafe_bene_data"
os.makedirs(DATA_DIR, exist_ok=True)

PATH_MASTER = os.path.join(DATA_DIR, "master_items.json")
PATH_CURRENT = os.path.join(DATA_DIR, "inventory_current.json")
PATH_HISTORY = os.path.join(DATA_DIR, "inventory_history.json")
PATH_DELETED = os.path.join(DATA_DIR, "inventory_deleted.json")

DEFAULT_FUZZY_THRESHOLD = 70  # Нэрээр тулгах босго оноо (0-100)

# ---- Responsive / хөнгөн загвар (CSS) ----
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans:wght@400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Noto Sans', sans-serif; }
    .main .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1200px;}
    .cb-header {
        background: linear-gradient(90deg,#4b2e19,#7a4a24);
        padding: 18px 22px; border-radius: 14px; margin-bottom: 14px;
        color: white;
    }
    .cb-header h1 {margin:0; font-size: 1.5rem;}
    .cb-header p {margin:0; opacity:.85; font-size:.9rem;}
    .cb-card {
        background:#fff; border:1px solid #eee; border-radius:12px;
        padding:14px 16px; margin-bottom:10px; box-shadow:0 1px 3px rgba(0,0,0,.05);
    }
    div[data-testid="stMetric"] {
        background:#faf6f2; border-radius:10px; padding:10px 6px; border:1px solid #eee;
    }
    @media (max-width: 640px){
        .cb-header h1 {font-size:1.15rem;}
        .main .block-container {padding-left:.6rem; padding-right:.6rem;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =======================================================================================
# 1. JSON DB ТУСЛАХ ФУНКЦУУД (UTF-8 бүрэн дэмжинэ, атомаар бичнэ)
# =======================================================================================
def load_json(path: str, default):
    if not os.path.exists(path):
        save_json(path, default)
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return default
            return json.loads(content)
    except (json.JSONDecodeError, UnicodeDecodeError):
        return default


def save_json(path: str, data):
    """
    Атомаар бичих: эхлээд түр файл руу бичээд, амжилттай бол os.replace-ээр
    жинхэнэ файлыг сольно. Апп ажиллаж байх үед унтарсан/crash хийсэн ч
    хуучин файл эвдрэхгүй байх нөхцлийг хангана.
    """
    folder = os.path.dirname(path) or "."
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


@st.cache_data(show_spinner=False)
def _load_master_cached(mtime):
    """mtime-г кэш түлхүүр болгон ашигладаг тул файл өөрчлөгдвөл автоматаар дахин уншина."""
    data = load_json(PATH_MASTER, [])
    if not data:
        return pd.DataFrame(columns=["Код", "Нэр"])
    df = pd.DataFrame(data)
    df = df.rename(columns={"code": "Код", "name": "Нэр"})
    if "Код" not in df.columns:
        df["Код"] = ""
    if "Нэр" not in df.columns:
        df["Нэр"] = ""
    return df[["Код", "Нэр"]].astype(str)


def load_master() -> pd.DataFrame:
    mtime = os.path.getmtime(PATH_MASTER) if os.path.exists(PATH_MASTER) else 0
    return _load_master_cached(mtime).copy()


def save_master(df: pd.DataFrame):
    records = [{"code": str(r["Код"]).strip(), "name": str(r["Нэр"]).strip()} for _, r in df.iterrows()
               if str(r["Код"]).strip() != ""]
    save_json(PATH_MASTER, records)
    _load_master_cached.clear()


def empty_count_row():
    return {"Код": "", "Нэр": "", "Өглөө": 0.0, "Хүргэлт": 0.0, "Орой": 0.0, "Тайлбар": ""}


MAX_IMAGE_MB = 5  # Нэг зургийн дээд хэмжээ (JSON файл хэт том болохоос сэргийлнэ)


def uploaded_image_to_record(uploaded_img) -> dict:
    """Upload хийсэн зургийг JSON-д хадгалахад тохиромжтой base64 dict болгоно."""
    raw_bytes = uploaded_img.getvalue()
    return {
        "name": uploaded_img.name,
        "mime": uploaded_img.type or "image/jpeg",
        "data": base64.b64encode(raw_bytes).decode("ascii"),
        "uploaded_at": datetime.now().isoformat(),
    }


def image_record_to_bytes(img_record: dict) -> bytes:
    return base64.b64decode(img_record["data"])


# =======================================================================================
# 2. ТУЛГАЛТЫН ЛОГИК
# =======================================================================================
def compute_actual(df: pd.DataFrame) -> pd.DataFrame:
    """Бодит = (Өглөө + Хүргэлт) - Орой"""
    df = df.copy()
    for col in ["Өглөө", "Хүргэлт", "Орой"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["Бодит"] = (df["Өглөө"] + df["Хүргэлт"]) - df["Орой"]
    return df


def find_col(columns, keywords_priority):
    """
    Багана нэрсээс түлхүүр үгтэй тохирохыг хайх (том/жижиг үсэг үл хамаарна).
    keywords_priority жагсаалтын эхний үг илүү өндөр давуу эрхтэй тул
    ойролцоо утгатай баганууд (ж: "Item #" ба "Item Name") хооронд зөв ялгана.
    """
    cols_lower = {c: str(c).lower().strip() for c in columns}
    for kw in keywords_priority:
        for c, low in cols_lower.items():
            if kw in low:
                return c
    return None


def find_header_row(raw_df: pd.DataFrame, max_scan: int = 25) -> int:
    """
    Толгой мөр нь эхний мөрөнд байхгүй тохиолдол (ж: огноо мэдээлэл дээр нь бичсэн)
    гарвал, эхний хэдэн мөрнөөс хамгийн олон түлхүүр үгтэй давхцсан мөрийг олж,
    түүнийг толгой мөр гэж тооцно.
    """
    keywords = ["item", "qty", "sold", "код", "code", "id", "нэр", "name", "plu", "price", "cost"]
    best_row, best_score = 0, -1
    for i in range(min(max_scan, len(raw_df))):
        row_vals = raw_df.iloc[i].fillna("").astype(str).str.lower().tolist()
        score = sum(1 for v in row_vals for kw in keywords if kw in v)
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def clean_code(val) -> str:
    """
    Тоон код 465.0 маягаар унших асуудлыг засаж '465' болгоно.
    Харин '007', '0012' зэрэг ЗОРИУД тэргүүн тэгтэй кодыг ХЭВЭЭР үлдээнэ
    (өмнөх хувилбарт эдгээр тэг устаж алдаа гарч байсныг засав).
    """
    s = str(val).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    # "0"-ээр эхэлсэн, зөвхөн тооноос бүрдсэн мөрийг тэр чигт нь хадгална
    if re.fullmatch(r"0\d+", s):
        return s
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return str(f)
    except (ValueError, TypeError):
        return s


def parse_system_excel(uploaded_file) -> pd.DataFrame:
    """
    Системийн Excel-ийг уншиж (Код, Нэр, Систем) баганатай нормчилно.
    - Толгой мөр өөр газар байх (ж: эхэнд огноо мөр) тохиолдлыг автоматаар илрүүлнэ.
    - "Item #"/"Item Name" зэрэг ойролцоо нэртэй баганыг зөв ялгана.
    - Дэд нийлбэр / хоосон мөрүүдийг шүүж хаяна (ЗӨВХӨН Нэр багана байгаа үед).
    - Ижил Код/Нэр давхар мөрөөр орж ирвэл (тайланд нэг бараа хэд хэдэн бүлэгт
      гарч ирдэг) тоог нь нэгтгэж нэмнэ.
    """
    uploaded_file.seek(0)
    raw = pd.read_excel(uploaded_file, engine="openpyxl", header=None)
    header_row = find_header_row(raw)

    uploaded_file.seek(0)
    df_raw = pd.read_excel(uploaded_file, engine="openpyxl", header=header_row)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    code_col = find_col(df_raw.columns, ["item #", "item#", "код", "plu", "code", "id"])
    name_col = find_col(df_raw.columns, ["item name", "нэр", "name", "бараа"])
    qty_col = find_col(df_raw.columns, ["qty sold", "qty_sold", "тоо", "sold", "qty"])

    if qty_col is None:
        raise ValueError(
            "Excel файлд 'Qty Sold' (борлуулсан тоо) багана олдсонгүй. "
            "Файлын толгой мөрийг шалгана уу."
        )
    if name_col is None and code_col is None:
        raise ValueError("Excel файлд Барааны Код эсвэл Нэр агуулсан багана олдсонгүй.")

    out = pd.DataFrame()
    out["Код"] = df_raw[code_col].apply(clean_code) if code_col else ""
    out["Нэр"] = df_raw[name_col].astype(str).str.strip() if name_col else ""
    out["Систем"] = pd.to_numeric(df_raw[qty_col], errors="coerce")

    # Дэд нийлбэр / хоосон (спэйсэр) мөрүүдийг хасах — ЗӨВХӨН Нэр багана
    # бодитоор байгаа тохиолдолд (үгүй бол бүх мөрийг санамсаргүй хаяж болзошгүй).
    if name_col:
        out["Нэр"] = out["Нэр"].replace({"nan": "", "None": ""})
        out = out[out["Нэр"].str.strip() != ""]

    out = out.dropna(subset=["Систем"])

    if out.empty:
        raise ValueError(
            "Барааны мөр олдсонгүй. Excel файл дэд-нийлбэрийн мөр л агуулсан "
            "эсвэл багана буруу таарсан байж болзошгүй."
        )

    # Нэг бараа тайланд хэд хэдэн бүлэгт (цаг/ангилал зэргээр) давхардаж
    # гарч ирдэг тул Код+Нэрээр нь нэгтгэж, тоог нь нэмнэ.
    out = out.groupby(["Код", "Нэр"], as_index=False)["Систем"].sum()
    return out


def parse_count_excel(uploaded_file) -> pd.DataFrame:
    """
    Хэрэглэгчийн бэлдсэн ТООЛЛОГЫН Excel файлыг уншиж (Код, Нэр, Өглөө, Хүргэлт,
    Орой, Тайлбар) баганатай, count_df-тэй ижил бүтэцтэй болгож нормчилно.
    - Толгой мөр өөр газар байх тохиолдлыг автоматаар илрүүлнэ (find_header_row).
    - Код, Нэр заавал биш (аль нэг нь байхад хангалттай) — гэхдээ Өглөө/Хүргэлт/Орой
      баганаас дор хаяж НЭГ нь байх шаардлагатай.
    - Ижил Код/Нэр давхар мөрөөр орж ирвэл тоог нь нэгтгэж нэмнэ.
    """
    uploaded_file.seek(0)
    raw = pd.read_excel(uploaded_file, engine="openpyxl", header=None)
    header_row = find_header_row(raw)

    uploaded_file.seek(0)
    df_raw = pd.read_excel(uploaded_file, engine="openpyxl", header=header_row)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    code_col = find_col(df_raw.columns, ["код", "plu", "code", "id", "item #", "item#"])
    name_col = find_col(df_raw.columns, ["нэр", "бараа", "name", "item name", "item"])
    morning_col = find_col(df_raw.columns, ["өглөө", "morning"])
    delivery_col = find_col(df_raw.columns, ["хүргэлт", "орлого", "delivery", "income"])
    evening_col = find_col(df_raw.columns, ["орой", "evening"])
    note_col = find_col(df_raw.columns, ["тайлбар", "note", "comment"])

    if name_col is None and code_col is None:
        raise ValueError("Тооллогын файлд Код эсвэл Нэр агуулсан багана олдсонгүй.")
    if morning_col is None and delivery_col is None and evening_col is None:
        raise ValueError(
            "Тооллогын файлд Өглөө / Хүргэлт(Орлого) / Орой баганаас дор хаяж нэг нь олдсонгүй. "
            "Файлын толгой мөрийг шалгана уу."
        )

    out = pd.DataFrame()
    out["Код"] = df_raw[code_col].apply(clean_code) if code_col else ""
    out["Нэр"] = df_raw[name_col].astype(str).str.strip() if name_col else ""
    out["Өглөө"] = pd.to_numeric(df_raw[morning_col], errors="coerce").fillna(0.0) if morning_col else 0.0
    out["Хүргэлт"] = pd.to_numeric(df_raw[delivery_col], errors="coerce").fillna(0.0) if delivery_col else 0.0
    out["Орой"] = pd.to_numeric(df_raw[evening_col], errors="coerce").fillna(0.0) if evening_col else 0.0
    out["Тайлбар"] = df_raw[note_col].astype(str).str.strip() if note_col else ""
    out["Тайлбар"] = out["Тайлбар"].replace({"nan": "", "None": ""})

    if name_col:
        out["Нэр"] = out["Нэр"].replace({"nan": "", "None": ""})
        out = out[out["Нэр"].str.strip() != ""]
    elif code_col:
        out = out[out["Код"].str.strip() != ""]

    if out.empty:
        raise ValueError(
            "Тооллогын мөр олдсонгүй. Файл дэд-нийлбэрийн мөр л агуулсан "
            "эсвэл багана буруу таарсан байж болзошгүй."
        )

    # Ижил Код+Нэр давхар мөрөөр орж ирвэл нэгтгэж нэмнэ (Тайлбарыг эхнийхээр нь авна).
    agg = {"Өглөө": "sum", "Хүргэлт": "sum", "Орой": "sum", "Тайлбар": "first"}
    out = out.groupby(["Код", "Нэр"], as_index=False).agg(agg)
    return out[["Код", "Нэр", "Өглөө", "Хүргэлт", "Орой", "Тайлбар"]]


def reconcile(df_count: pd.DataFrame, df_system: pd.DataFrame, fuzzy_threshold: int):
    """
    Код-оор эхлээд тулгана, олдохгүй бол Fuzzy search-ээр нэрээр тулгана.
    Зөрүү = Бодит - Систем

    Буцаах утга: (df, unmatched_system_df) хос (tuple).
    Санамж: unmatched_system_df-г df.attrs дотор ХАДГАЛАХГҮЙ — учир нь
    pandas-ийн зарим хувилбарт (concat/astype) attrs дотор DataFrame байхад
    "truth value of a DataFrame is ambiguous" гэсэн алдаа гарна. Тиймээс
    тусад нь, энгийн tuple-ээр буцаана.
    """
    df = compute_actual(df_count)

    code_map = {}
    if "Код" in df_system.columns:
        for _, r in df_system.iterrows():
            code = str(r["Код"]).strip()
            if code and code.lower() != "nan":
                code_map[code] = r["Систем"]

    name_map = {}
    if "Нэр" in df_system.columns:
        for _, r in df_system.iterrows():
            nm = str(r["Нэр"]).strip()
            if nm and nm.lower() != "nan":
                name_map[nm] = r["Систем"]
    system_names = list(name_map.keys())

    system_qty_list = []
    match_method_list = []
    matched_keys = set()  # системийн аль мөрүүд тулгагдсаныг мөшгих (Код эсвэл Нэр)

    for _, row in df.iterrows():
        code = str(row.get("Код", "")).strip()
        name = str(row.get("Нэр", "")).strip()
        sys_qty = None
        method = "Олдсонгүй"

        # 1) Код-оор тулгах
        if code and code in code_map:
            sys_qty = code_map[code]
            method = "Код"
            matched_keys.add(("code", code))
        # 2) Fuzzy search — нэрээр тулгах
        elif name and system_names:
            best = process.extractOne(name, system_names, scorer=fuzz.token_sort_ratio)
            if best and best[1] >= fuzzy_threshold:
                sys_qty = name_map[best[0]]
                method = f"Fuzzy ({best[1]}%) → {best[0]}"
                matched_keys.add(("name", best[0]))

        if sys_qty is None:
            sys_qty = 0.0

        system_qty_list.append(sys_qty)
        match_method_list.append(method)

    df["Систем"] = system_qty_list
    df["Тулгасан аргаас"] = match_method_list
    df["Зөрүү"] = df["Бодит"] - df["Систем"]

    # Тоолоогүй ч Системд байгаа (алгассан) мөрүүдийг тэмдэглэх
    unmatched_rows = []
    for _, r in df_system.iterrows():
        code = str(r.get("Код", "")).strip()
        name = str(r.get("Нэр", "")).strip()
        if ("code", code) in matched_keys or ("name", name) in matched_keys:
            continue
        unmatched_rows.append({"Код": code, "Нэр": name, "Систем": r["Систем"]})

    unmatched_system_df = pd.DataFrame(unmatched_rows) if unmatched_rows else pd.DataFrame(
        columns=["Код", "Нэр", "Систем"]
    )
    return df, unmatched_system_df


def color_diff(val):
    try:
        v = float(val)
    except (ValueError, TypeError):
        return ""
    if v < 0:
        return "color:#c0392b; font-weight:700; background-color:#fdecea;"
    elif v > 0:
        return "color:#1e8449; font-weight:700; background-color:#eafaf1;"
    return "color:#555;"


def styled_table(df: pd.DataFrame, diff_col="Зөрүү"):
    cols = [c for c in df.columns if c in
            ["Код", "Нэр", "Өглөө", "Хүргэлт", "Орой", "Бодит", "Систем", "Зөрүү", "Тулгасан аргаас", "Тайлбар"]]
    view = df[cols] if cols else df
    sty = view.style
    if diff_col in view.columns:
        # pandas >=2.1: Styler.map; хуучин хувилбарт applymap ашиглана.
        try:
            sty = sty.map(color_diff, subset=[diff_col])
        except AttributeError:
            sty = sty.applymap(color_diff, subset=[diff_col])
    fmt = {c: "{:.1f}" for c in ["Өглөө", "Хүргэлт", "Орой", "Бодит", "Систем", "Зөрүү"] if c in view.columns}
    sty = sty.format(fmt)
    return sty


def df_to_excel_bytes(df: pd.DataFrame, sheet_name="Тайлан") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


# =======================================================================================
# 3. SESSION STATE ЭХЛҮҮЛЭХ
# =======================================================================================
if "count_df" not in st.session_state:
    saved_current = load_json(PATH_CURRENT, None)
    if saved_current and saved_current.get("items"):
        st.session_state.count_df = pd.DataFrame(saved_current["items"])
    else:
        master_df = load_master()
        if not master_df.empty:
            rows = []
            for _, r in master_df.iterrows():
                row = empty_count_row()
                row["Код"] = r["Код"]
                row["Нэр"] = r["Нэр"]
                rows.append(row)
            st.session_state.count_df = pd.DataFrame(rows)
        else:
            st.session_state.count_df = pd.DataFrame([empty_count_row()])

if "count_images" not in st.session_state:
    saved_current = load_json(PATH_CURRENT, None)
    st.session_state.count_images = (saved_current or {}).get("images", []) or []

if "reconciled_df" not in st.session_state:
    st.session_state.reconciled_df = None

if "unmatched_system_df" not in st.session_state:
    st.session_state.unmatched_system_df = None

if "confirm_clear_table" not in st.session_state:
    st.session_state.confirm_clear_table = False

if "confirm_delete_id" not in st.session_state:
    st.session_state.confirm_delete_id = None

if "fuzzy_threshold" not in st.session_state:
    st.session_state.fuzzy_threshold = DEFAULT_FUZZY_THRESHOLD


# =======================================================================================
# 4. HEADER
# =======================================================================================
st.markdown(
    """
    <div class="cb-header">
        <h1>☕ Caffe Bene — Тооллого & Борлуулалтын систем тулгалт</h1>
        <p>Өдрийн тооллого · Системтэй тулгах · Архив & Хогийн сав · Барааны мэдээллийн сан</p>
    </div>
    """,
    unsafe_allow_html=True,
)

tab1, tab2, tab3 = st.tabs(["📝 ТООЛЛОГО", "📊 АРХИВ", "⚙️ Тохиргоо"])


# =======================================================================================
# TAB 1 — ТООЛЛОГО
# =======================================================================================
with tab1:
    st.subheader("📝 Өдрийн тооллого")

    col_date, col_btn1, col_btn2 = st.columns([2, 1, 1])
    with col_date:
        count_date = st.date_input("Тооллогын огноо", value=date.today())

    master_df = load_master()
    code_to_name = dict(zip(master_df["Код"], master_df["Нэр"])) if not master_df.empty else {}

    with st.expander("📥 Тооллогын Excel файл болон зураг оруулах (гараар бичихийн оронд)", expanded=False):
        st.caption(
            "Файл нь `Код`/`Нэр` (аль нэг нь заавал) болон `Өглөө`, `Хүргэлт`(Орлого), "
            "`Орой` баганын аль нэгийг агуулсан байх ёстой. `Тайлбар` багана заавал биш."
        )
        count_file = st.file_uploader(
            "Тооллого хийсэн Excel файл", type=["xlsx", "xls"], key="count_upload"
        )
        if count_file is not None:
            try:
                df_count_uploaded = parse_count_excel(count_file)
                st.success(f"Тооллогын файлаас {len(df_count_uploaded)} мөр амжилттай уншлаа.")
                st.dataframe(df_count_uploaded, use_container_width=True, hide_index=True)

                load_mode = st.radio(
                    "Одоогийн хүснэгттэй хэрхэн нэгтгэх вэ?",
                    ["Солих (хуучныг устгаад орлуулна)", "Нэмэх (хуучин мөрний ард нэмнэ)"],
                    horizontal=True,
                )
                if st.button("⬆️ Хүснэгтэд ачаалах", type="primary", use_container_width=True):
                    if load_mode.startswith("Солих"):
                        st.session_state.count_df = df_count_uploaded.reset_index(drop=True)
                    else:
                        st.session_state.count_df = pd.concat(
                            [st.session_state.count_df, df_count_uploaded], ignore_index=True
                        )
                    st.session_state.reconciled_df = None
                    st.session_state.unmatched_system_df = None
                    st.success("Тооллогын хүснэгт шинэчлэгдлээ.")
                    st.rerun()
            except Exception as e:
                st.error(f"Файл уншихад алдаа гарлаа: {e}")

        st.divider()
        st.write("**📷 Тооллогын нотолгоо зураг (жиших бэлэн бичсэн хуудас, дэлгэцийн зураг гэх мэт)**")
        uploaded_imgs = st.file_uploader(
            "Зураг нэмэх (олон зураг зэрэг сонгож болно)",
            type=["jpg", "jpeg", "png", "webp"],
            accept_multiple_files=True,
            key="count_image_upload",
        )
        if uploaded_imgs:
            existing_names = {img["name"] for img in st.session_state.count_images}
            added = 0
            for img in uploaded_imgs:
                size_mb = len(img.getvalue()) / (1024 * 1024)
                if size_mb > MAX_IMAGE_MB:
                    st.error(f"'{img.name}' — {size_mb:.1f}MB, {MAX_IMAGE_MB}MB-с их тул алгасав.")
                    continue
                if img.name in existing_names:
                    continue
                st.session_state.count_images.append(uploaded_image_to_record(img))
                added += 1
            if added:
                st.success(f"{added} зураг нэмэгдлээ.")
                st.rerun()

        if st.session_state.count_images:
            st.caption(f"Хавсаргасан зураг ({len(st.session_state.count_images)}):")
            img_cols = st.columns(4)
            for idx, img_rec in enumerate(list(st.session_state.count_images)):
                with img_cols[idx % 4]:
                    st.image(image_record_to_bytes(img_rec), caption=img_rec["name"], use_container_width=True)
                    if st.button("🗑️ Устгах", key=f"del_img_{idx}_{img_rec['name']}", use_container_width=True):
                        st.session_state.count_images.pop(idx)
                        st.rerun()
        else:
            st.caption("Одоогоор зураг хавсаргаагүй байна.")

    st.caption("Мөр бүрт Өглөө / Хүргэлт (Орлого) / Орой-ийн тоог оруулна уу. "
               "Код бичихэд мастер жагсаалтад байгаа бол нэр автоматаар бөглөгдөнө. "
               "Шинэ мөр нэмэхдээ хүснэгтийн доод хэсгийн **+** товч ашиглана.")

    edited_df = st.data_editor(
        st.session_state.count_df,
        num_rows="dynamic",
        use_container_width=True,
        key="count_editor",
        column_config={
            "Код": st.column_config.TextColumn("Код (PLU)", width="small"),
            "Нэр": st.column_config.TextColumn("Барааны нэр", width="medium"),
            "Өглөө": st.column_config.NumberColumn("Өглөө (Ө)", min_value=0.0, step=1.0, format="%.1f"),
            "Хүргэлт": st.column_config.NumberColumn("Хүргэлт/Орлого (Х)", min_value=0.0, step=1.0, format="%.1f"),
            "Орой": st.column_config.NumberColumn("Орой (О)", min_value=0.0, step=1.0, format="%.1f"),
            "Тайлбар": st.column_config.TextColumn("Тайлбар", width="large"),
        },
        hide_index=True,
    )

    # Код бичвэл нэрийг автоматаар бөглөх (Нэр хоосон үед л дарж бичихгүй байхаар)
    if not master_df.empty:
        for i in edited_df.index:
            c = str(edited_df.at[i, "Код"]).strip()
            existing_name = str(edited_df.at[i, "Нэр"]).strip()
            if c and c in code_to_name and not existing_name:
                edited_df.at[i, "Нэр"] = code_to_name[c]

    st.session_state.count_df = edited_df

    calc_df = compute_actual(edited_df)
    total_actual = calc_df["Бодит"].sum()

    m1, m2, m3 = st.columns(3)
    m1.metric("Мөрийн тоо", len(calc_df))
    m2.metric("Нийт Бодит (Ө+Х-О)", f"{total_actual:,.1f}")
    m3.metric("Огноо", count_date.strftime("%Y-%m-%d"))

    st.write("**Тооцоолсон Бодит зарагдсан тоо (тулгалтын өмнөх):**")
    st.dataframe(
        calc_df[["Код", "Нэр", "Өглөө", "Хүргэлт", "Орой", "Бодит", "Тайлбар"]],
        use_container_width=True, hide_index=True,
    )

    col_save1, col_save2 = st.columns(2)
    with col_save1:
        if st.button("💾 Түр хадгалах (Draft)", use_container_width=True):
            save_json(PATH_CURRENT, {
                "date": count_date.strftime("%Y-%m-%d"),
                "items": edited_df.to_dict(orient="records"),
                "images": st.session_state.count_images,
                "saved_at": datetime.now().isoformat(),
            })
            st.success("Түр хадгаллаа (хүснэгт болон зураг). Дараа нэвтрэхэд энэ өгөгдөл сэргэнэ.")
    with col_save2:
        if not st.session_state.confirm_clear_table:
            if st.button("🗑️ Хүснэгтийг цэвэрлэх", use_container_width=True):
                st.session_state.confirm_clear_table = True
                st.rerun()
        else:
            st.warning("Хүснэгт болон хавсаргасан зургууд бүгд устна. Итгэлтэй байна уу?")
            cc1, cc2 = st.columns(2)
            if cc1.button("✅ Тийм, цэвэрлэ", type="primary", use_container_width=True):
                st.session_state.count_df = pd.DataFrame([empty_count_row()])
                st.session_state.reconciled_df = None
                st.session_state.unmatched_system_df = None
                st.session_state.count_images = []
                st.session_state.confirm_clear_table = False
                st.rerun()
            if cc2.button("❌ Үгүй, болих", use_container_width=True):
                st.session_state.confirm_clear_table = False
                st.rerun()

    st.divider()
    st.subheader("🔄 Системийн Excel-тэй тулгах")
    st.caption("Excel файл нь `Код`/`ID`, `Нэр`(заавал биш) болон **`Qty Sold`** баганатай байх ёстой.")

    fuzzy_threshold = st.slider(
        "Нэрээр тулгах Fuzzy босго (%)",
        min_value=40, max_value=100,
        value=st.session_state.fuzzy_threshold,
        help="Код тулгагдаагүй үед нэрээр ойролцоо тааруулах босго оноо. "
             "Бага утга — илүү чөлөөтэй тааруулна (алдаа гарах магадлал ихэснэ).",
    )
    st.session_state.fuzzy_threshold = fuzzy_threshold

    sys_file = st.file_uploader("Системийн борлуулалтын Excel файл", type=["xlsx", "xls"], key="sys_upload")

    if sys_file is not None:
        try:
            df_system = parse_system_excel(sys_file)
            st.success(f"Системийн файлаас {len(df_system)} мөр амжилттай уншлаа.")
            if st.button("⚖️ Тулгалт хийх", type="primary", use_container_width=True):
                reconciled, unmatched_system_df = reconcile(edited_df, df_system, fuzzy_threshold)
                st.session_state.reconciled_df = reconciled
                st.session_state.unmatched_system_df = unmatched_system_df
        except Exception as e:
            st.error(f"Файл уншихад алдаа гарлаа: {e}")

    if st.session_state.reconciled_df is not None:
        rdf = st.session_state.reconciled_df
        st.write("**Тулгалтын үр дүн** (🟥 Дутсан — Улаан | 🟩 Илүүдсэн — Ногоон):")
        st.dataframe(styled_table(rdf), use_container_width=True, hide_index=True)

        d1, d2, d3 = st.columns(3)
        d1.metric("Дутсан барааны тоо", int((rdf["Зөрүү"] < 0).sum()))
        d2.metric("Илүүдсэн барааны тоо", int((rdf["Зөрүү"] > 0).sum()))
        d3.metric("Тохирсон барааны тоо", int((rdf["Зөрүү"] == 0).sum()))

        unmatched_sys = st.session_state.unmatched_system_df
        if unmatched_sys is not None and not unmatched_sys.empty:
            with st.expander(f"⚠️ Тоолоогүй ч Системд зарагдсан бараа ({len(unmatched_sys)})", expanded=False):
                st.caption("Эдгээр бараа тооллогын хүснэгтэд байхгүй эсвэл нэр/код таарахгүй байна.")
                st.dataframe(unmatched_sys, use_container_width=True, hide_index=True)

        if st.button("📦 Архивлах (Тулгалтыг баталгаажуулж хадгалах)", type="primary", use_container_width=True):
            history = load_json(PATH_HISTORY, [])
            record = {
                "id": str(uuid.uuid4()),
                "date": count_date.strftime("%Y-%m-%d"),
                "archived_at": datetime.now().isoformat(),
                "items": rdf.to_dict(orient="records"),
                "unmatched_system": unmatched_sys.to_dict(orient="records") if unmatched_sys is not None else [],
                "images": st.session_state.count_images,
            }
            history.append(record)
            save_json(PATH_HISTORY, history)

            # Түр хадгалалтыг цэвэрлэх
            save_json(PATH_CURRENT, {"date": "", "items": [], "images": [], "saved_at": ""})
            st.session_state.count_df = pd.DataFrame([empty_count_row()])
            st.session_state.reconciled_df = None
            st.session_state.unmatched_system_df = None
            st.session_state.count_images = []
            st.success(f"{count_date.strftime('%Y-%m-%d')} өдрийн тооллого архивлагдлаа!")
            st.rerun()


# =======================================================================================
# TAB 2 — АРХИВ
# =======================================================================================
with tab2:
    st.subheader("📊 Тулгалтын түүх / Архив")

    history = load_json(PATH_HISTORY, [])
    deleted = load_json(PATH_DELETED, [])

    if not history:
        st.info("Одоогоор архивласан тооллого алга байна.")
    else:
        hist_df_meta = pd.DataFrame([
            {"id": r["id"], "Огноо": r["date"], "Архивласан": r.get("archived_at", ""),
             "Мөрийн тоо": len(r.get("items", []))}
            for r in history
        ])
        hist_df_meta["Сар"] = pd.to_datetime(hist_df_meta["Огноо"], errors="coerce").dt.strftime("%Y-%m")

        months = ["Бүгд"] + sorted(hist_df_meta["Сар"].dropna().unique().tolist(), reverse=True)
        sel_month = st.selectbox("📅 Сараар шүүх", months)

        filtered_meta = hist_df_meta if sel_month == "Бүгд" else hist_df_meta[hist_df_meta["Сар"] == sel_month]

        st.dataframe(filtered_meta[["Огноо", "Архивласан", "Мөрийн тоо"]], use_container_width=True, hide_index=True)

        record_options = {f'{r["date"]} — {r["id"][:8]}': r["id"] for r in history
                           if sel_month == "Бүгд" or str(r["date"])[:7] == sel_month}

        if record_options:
            sel_label = st.selectbox("Дэлгэрэнгүй харах тайлан сонгох", list(record_options.keys()))
            sel_id = record_options[sel_label]
            sel_record = next(r for r in history if r["id"] == sel_id)
            rdf = pd.DataFrame(sel_record["items"])

            st.write(f"### 🧾 {sel_record['date']} өдрийн тайлан")
            if "Зөрүү" in rdf.columns:
                st.dataframe(styled_table(rdf), use_container_width=True, hide_index=True)
            else:
                st.dataframe(rdf, use_container_width=True, hide_index=True)

            unmatched_hist = pd.DataFrame(sel_record.get("unmatched_system", []))
            if not unmatched_hist.empty:
                with st.expander(f"⚠️ Тухайн өдөр тоолоогүй ч Системд зарагдсан бараа ({len(unmatched_hist)})"):
                    st.dataframe(unmatched_hist, use_container_width=True, hide_index=True)

            hist_images = sel_record.get("images", [])
            if hist_images:
                with st.expander(f"📷 Хавсаргасан зураг ({len(hist_images)})", expanded=False):
                    img_cols_hist = st.columns(4)
                    for idx, img_rec in enumerate(hist_images):
                        with img_cols_hist[idx % 4]:
                            st.image(
                                image_record_to_bytes(img_rec),
                                caption=img_rec.get("name", ""),
                                use_container_width=True,
                            )

            c1, c2 = st.columns(2)
            with c1:
                excel_bytes = df_to_excel_bytes(rdf, sheet_name=sel_record["date"])
                st.download_button(
                    "⬇️ Excel-ээр татах",
                    data=excel_bytes,
                    file_name=f"CaffeBene_tailan_{sel_record['date']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with c2:
                if st.session_state.confirm_delete_id != sel_id:
                    if st.button("🗑️ Устгах (Хогийн саванд шилжүүлэх)", use_container_width=True):
                        st.session_state.confirm_delete_id = sel_id
                        st.rerun()
                else:
                    st.warning("Энэ тайланг хогийн саванд шилжүүлэх үү?")
                    dc1, dc2 = st.columns(2)
                    if dc1.button("✅ Тийм", type="primary", use_container_width=True, key="confirm_del_yes"):
                        history = [r for r in history if r["id"] != sel_id]
                        sel_record["deleted_at"] = datetime.now().isoformat()
                        deleted.append(sel_record)
                        save_json(PATH_HISTORY, history)
                        save_json(PATH_DELETED, deleted)
                        st.session_state.confirm_delete_id = None
                        st.warning("Тайланг хогийн саванд шилжүүллээ.")
                        st.rerun()
                    if dc2.button("❌ Үгүй", use_container_width=True, key="confirm_del_no"):
                        st.session_state.confirm_delete_id = None
                        st.rerun()

            # Сарын нэгтгэл
            st.divider()
            st.write("### 📈 Сарын нэгтгэл")
            all_month_items = []
            for r in history:
                if sel_month == "Бүгд" or str(r["date"])[:7] == sel_month:
                    for item in r.get("items", []):
                        item = dict(item)
                        item["Огноо"] = r["date"]
                        all_month_items.append(item)
            if all_month_items:
                month_df = pd.DataFrame(all_month_items)
                if "Зөрүү" in month_df.columns:
                    summary = month_df.groupby("Нэр", dropna=False).agg(
                        Бодит=("Бодит", "sum"),
                        Систем=("Систем", "sum"),
                        Зөрүү=("Зөрүү", "sum"),
                    ).reset_index()
                    st.dataframe(styled_table(summary), use_container_width=True, hide_index=True)
                    month_excel = df_to_excel_bytes(summary, sheet_name="Сарын_нэгтгэл")
                    st.download_button(
                        "⬇️ Сарын нэгтгэлийг Excel-ээр татах",
                        data=month_excel,
                        file_name=f"CaffeBene_saryn_negtgel_{sel_month}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    st.divider()
    with st.expander(f"🗑️ Хогийн сав ({len(deleted)})", expanded=False):
        if not deleted:
            st.caption("Хогийн сав хоосон байна.")
        else:
            for r in deleted:
                cols = st.columns([3, 2, 2])
                cols[0].write(f"**{r['date']}** — {r['id'][:8]}")
                cols[1].write(f"Устгасан: {r.get('deleted_at', '—')}")
                if cols[2].button("♻️ Сэргээх", key=f"restore_{r['id']}"):
                    deleted = [d for d in deleted if d["id"] != r["id"]]
                    r.pop("deleted_at", None)
                    history.append(r)
                    save_json(PATH_HISTORY, history)
                    save_json(PATH_DELETED, deleted)
                    st.success("Тайланг сэргээлээ.")
                    st.rerun()


# =======================================================================================
# TAB 3 — БАРААНЫ МЕШЕН (Master list удирдлага / сургах)
# =======================================================================================
with tab3:
    st.subheader("⚙️ Барааны мэдээллийн сан (Master List)")
    st.caption("Шинэ барааны Код (PLU) болон Нэрийг бүртгэж, Fuzzy тулгалтын 'сургалтад' ашиглана.")

    master_df = load_master()

    with st.form("add_item_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([1, 2, 1])
        new_code = c1.text_input("Барааны код (PLU/ID)")
        new_name = c2.text_input("Барааны нэр")
        submitted = c3.form_submit_button("➕ Нэмэх", use_container_width=True)

        if submitted:
            if not new_code.strip() or not new_name.strip():
                st.error("Код болон нэрийг хоёуланг нь бөглөнө үү.")
            elif new_code.strip() in master_df["Код"].values:
                st.error(f"'{new_code}' код аль хэдийн бүртгэлтэй байна.")
            else:
                new_row = pd.DataFrame([{"Код": new_code.strip(), "Нэр": new_name.strip()}])
                master_df = pd.concat([master_df, new_row], ignore_index=True)
                save_master(master_df)
                st.success(f"'{new_name}' ({new_code}) амжилттай нэмэгдлээ.")
                st.rerun()

    st.divider()
    st.write("### 📥 Excel-ээр багцаар оруулах")
    bulk_file = st.file_uploader("Код, Нэр баганатай Excel файл", type=["xlsx", "xls"], key="bulk_master")
    if bulk_file is not None:
        try:
            df_bulk = pd.read_excel(bulk_file, engine="openpyxl")
            df_bulk.columns = [str(c).strip() for c in df_bulk.columns]
            code_col = find_col(df_bulk.columns, ["код", "code", "id", "plu"])
            name_col = find_col(df_bulk.columns, ["нэр", "name", "item", "бараа"])
            if not code_col or not name_col:
                st.error("Файлд Код болон Нэр багана олдсонгүй.")
            else:
                preview = df_bulk[[code_col, name_col]].rename(columns={code_col: "Код", name_col: "Нэр"})
                preview["Код"] = preview["Код"].apply(clean_code)
                preview["Нэр"] = preview["Нэр"].astype(str).str.strip()
                st.dataframe(preview, use_container_width=True, hide_index=True)
                if st.button("✅ Мастер жагсаалтад нэгтгэх", type="primary"):
                    merged = pd.concat([master_df, preview], ignore_index=True)
                    merged = merged.drop_duplicates(subset=["Код"], keep="last")
                    save_master(merged)
                    st.success(f"{len(preview)} мөр нэгтгэгдлээ.")
                    st.rerun()
        except Exception as e:
            st.error(f"Файл уншихад алдаа гарлаа: {e}")

    st.divider()
    st.write(f"### 📋 Одоогийн бүртгэлтэй бараанууд ({len(master_df)})")

    if master_df.empty:
        st.info("Мастер жагсаалт хоосон байна. Дээрх формоор эхний бараагаа нэмнэ үү.")
    else:
        search_q = st.text_input("🔍 Хайх (код эсвэл нэрээр)")
        show_df = master_df.copy()
        if search_q.strip():
            mask = (show_df["Код"].str.contains(search_q, case=False, na=False, regex=False) |
                    show_df["Нэр"].str.contains(search_q, case=False, na=False, regex=False))
            show_df = show_df[mask]

        edited_master = st.data_editor(
            show_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="master_editor",
            column_config={
                "Код": st.column_config.TextColumn("Код (PLU)", width="small"),
                "Нэр": st.column_config.TextColumn("Барааны нэр", width="large"),
            },
        )

        if st.button("💾 Мастер жагсаалтын өөрчлөлтийг хадгалах", type="primary", use_container_width=True):
            if search_q.strip():
                unaffected = master_df[~master_df["Код"].isin(show_df["Код"])]
                final_df = pd.concat([unaffected, edited_master], ignore_index=True)
            else:
                final_df = edited_master
            final_df = final_df.drop_duplicates(subset=["Код"], keep="last")
            save_master(final_df)
            st.success("Мастер жагсаалт шинэчлэгдлээ.")
            st.rerun()

        excel_master = df_to_excel_bytes(master_df, sheet_name="Master_List")
        st.download_button(
            "⬇️ Мастер жагсаалтыг Excel-ээр татах",
            data=excel_master,
            file_name="CaffeBene_master_items.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# =======================================================================================
# SIDEBAR — Товч заавар
# =======================================================================================
with st.sidebar:
    st.markdown("### ☕ Caffe Bene")
    st.caption("Тооллого & Тулгалтын систем")
    st.markdown("---")
    st.markdown(
        """
        **Ажиллах дараалал:**
        1. 📝 **ТООЛЛОГО** — Ө/Х/О тоог шивнэ, түр хадгална.
        2. Системийн Excel (`Qty Sold`) upload хийж **Тулгалт хийх**.
        3. Зөрүүг шалгаад **Архивлах**.
        4. 📊 **АРХИВ** — сараар харах, Excel татах, устгах/сэргээх.
        5. ⚙️ **Тохиргоо** — шинэ код/нэр бүртгэх.
        """
    )
    st.markdown("---")
    st.caption(f"Өгөгдлийн сан: `{DATA_DIR}/`")
    st.caption("© Caffe Bene — Дотоод хэрэглээний систем")
