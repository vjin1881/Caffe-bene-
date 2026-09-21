# -*- coding: utf-8 -*-
"""
=======================================================================================
 CAFFE BENE — Өдрийн тооллого, Борлуулалтын систем тулгалтын веб апп  (v2)
=======================================================================================
Ашиглах сангууд:
    pip install streamlit pandas numpy rapidfuzz openpyxl
    (rapidfuzz суугаагүй бол fuzzywuzzy-г ашиглана)

Ажиллуулах:
    streamlit run app.py

v2-т зассан зүйлс:
    * Код багана Selectbox байсныг Text болгосон (мастерт байхгүй код оруулахад унадаг байсан)
    * data_editor-ийн key-г цэвэрлэдэг болсон (ачаалсан өгөгдөл буцаж өөрчлөгдөх алдаа,
      мөр нэмэхэд шинэ мөр "алга болдог" алдаа)
    * Styler.applymap → .map (pandas 3 дээр унадаг байсан)
    * Тоон баганыг format хийхээс өмнө цэвэрлэдэг болсон
    * parse_system_excel — Код багана байхгүй үед эвдэрдэг байсныг зассан
    * st.rerun() формын дотроос гаргасан, мессеж харагддаг болсон
    * JSON-г atomic бичдэг болсон + Нөөцлөх/Сэргээх функц нэмсэн
v2-т нэмсэн зүйлс:
    * Хаягдал / Дотоод хэрэглээний багана
    * Тулгагдаагүй барааг 0 биш "тулгагдаагүй" гэж ялгадаг болсон
    * Системд зарагдсан ч тооллогод ороогүй барааг илрүүлдэг болсон
    * Fuzzy тулгалт нэг системийн барааг давхар ашиглахгүй болсон
    * Өмнөх өдрийн Орой → өнөөдрийн Өглөө автоматаар татах
    * Нэгж үнэ → зөрүүг төгрөгөөр тооцох
    * Зөвшөөрөх хязгаар (tolerance), Fuzzy босгыг sidebar-аас тохируулах
    * Excel экспорт олон хуудастай, сарын тренд шинжилгээ
=======================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import io
import uuid
import tempfile
from datetime import datetime, date, timedelta

# ---- Fuzzy сан: rapidfuzz-ийг эрхэмлэнэ, байхгүй бол fuzzywuzzy ----
FUZZ_LIB = None
try:
    from rapidfuzz import fuzz, process
    FUZZ_LIB = "rapidfuzz"
except ImportError:
    try:
        from fuzzywuzzy import fuzz, process
        FUZZ_LIB = "fuzzywuzzy"
    except ImportError:
        st.error(
            "Fuzzy сан суугаагүй байна. Терминал дээр:\n\n"
            "    pip install rapidfuzz\n\n"
            "(эсвэл: pip install fuzzywuzzy python-Levenshtein)"
        )
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
PATH_PHOTOS = os.path.join(DATA_DIR, "photos")
os.makedirs(PATH_PHOTOS, exist_ok=True)

NONE_OPT = "— Байхгүй (0 / хоосон) —"

# Тооллогын хүснэгтийн стандарт баганууд
COUNT_COLS = ["Код", "Нэр", "Өглөө", "Хүргэлт", "Орой", "Хаягдал", "Дотоод", "Тайлбар"]
NUM_COLS = ["Өглөө", "Хүргэлт", "Орой", "Хаягдал", "Дотоод"]
CALC_COLS = ["Бодит", "Тооц.борл", "Систем", "Зөрүү", "Зөрүү ₮"]

# ---- Responsive / хөнгөн загвар (CSS) ----
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans:wght@400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Noto Sans', sans-serif; }
    .main .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1280px;}
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
# 1. JSON DB ТУСЛАХ ФУНКЦУУД (UTF-8 + atomic бичилт)
# =======================================================================================
def save_json(path: str, data):
    """
    Түр файлд бичээд os.replace хийнэ. Ингэснээр бичиж байх үед програм
    унтарсан ч хуучин файл эвдрэхгүй (atomic write).
    """
    directory = os.path.dirname(path) or "."
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(dir=directory, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


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
    except (json.JSONDecodeError, UnicodeDecodeError, OSError):
        return default


def load_master() -> pd.DataFrame:
    """Мастер жагсаалт: Код / Нэр / Үнэ. Хуучин (код, нэр) форматтай ч нийцнэ."""
    data = load_json(PATH_MASTER, [])
    if not data:
        return pd.DataFrame(columns=["Код", "Нэр", "Үнэ"])
    df = pd.DataFrame(data)
    df = df.rename(columns={"code": "Код", "name": "Нэр", "price": "Үнэ"})
    for col in ["Код", "Нэр"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).replace({"nan": "", "None": ""}).str.strip()
    if "Үнэ" not in df.columns:
        df["Үнэ"] = 0.0
    df["Үнэ"] = pd.to_numeric(df["Үнэ"], errors="coerce").fillna(0.0)
    df = df[df["Код"] != ""]
    return df[["Код", "Нэр", "Үнэ"]].reset_index(drop=True)


def save_master(df: pd.DataFrame):
    records = []
    for _, r in df.iterrows():
        code = str(r.get("Код", "")).strip()
        if code == "" or code.lower() == "nan":
            continue
        try:
            price = float(pd.to_numeric(r.get("Үнэ", 0), errors="coerce"))
        except (TypeError, ValueError):
            price = 0.0
        if pd.isna(price):
            price = 0.0
        records.append({
            "code": code,
            "name": str(r.get("Нэр", "")).strip(),
            "price": price,
        })
    save_json(PATH_MASTER, records)


def empty_count_row():
    return {"Код": "", "Нэр": "", "Өглөө": 0.0, "Хүргэлт": 0.0,
            "Орой": 0.0, "Хаягдал": 0.0, "Дотоод": 0.0, "Тайлбар": ""}


def ensure_count_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Хуучин хадгалсан өгөгдөлд шинэ багана байхгүй байвал нэмж өгнө."""
    df = df.copy()
    for col in COUNT_COLS:
        if col not in df.columns:
            df[col] = 0.0 if col in NUM_COLS else ""
    for col in NUM_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    for col in ["Код", "Нэр", "Тайлбар"]:
        df[col] = df[col].astype(str).replace({"nan": "", "None": ""})
    return df[COUNT_COLS]


# =======================================================================================
# 2. ТӨЛӨВ УДИРДАХ ТУСЛАХУУД (Streamlit-ийн нийтлэг алдаанаас сэргийлэх)
# =======================================================================================
def set_count_df(df: pd.DataFrame):
    """
    Тооллогын хүснэгтийг солихдоо data_editor-ийн key-г ЗААВАЛ цэвэрлэнэ.
    Үгүй бол хуучин засварууд шинэ өгөгдлийн өөр мөрүүд дээр наалдана.
    """
    st.session_state.count_df = ensure_count_cols(df).reset_index(drop=True)
    st.session_state.pop("count_editor", None)
    st.session_state.reconciled_df = None
    st.session_state.missing_df = None


def flash(msg: str, icon: str = "✅"):
    """rerun-ы дараа харагдах мессеж дараалалд нэмнэ."""
    st.session_state.setdefault("_flash", []).append((msg, icon))


def render_flash():
    for msg, icon in st.session_state.pop("_flash", []):
        st.toast(msg, icon=icon)


def show_image(path):
    """Streamlit хувилбар хоорондын нийцэл."""
    try:
        st.image(path, use_container_width=True)
    except TypeError:
        st.image(path, use_column_width=True)


def editor_height(n_rows: int, min_rows: int = 8, max_px: int = 900) -> int:
    """
    st.data_editor анхныхаараа зөвхөн ~5-6 мөр харуулаад дотроо гүйлгэдэг тул
    (мөр олон бол хэрэглэгч "зөвхөн 5 мөр л харагдлаа" гэж андуурдаг) —
    мөрийн тооноос хамааруулж өндрийг автоматаар тооцож, дотор нь гүйлгэх
    шаардлагагүй болгоно. Хэт олон мөртэй бол max_px-д хүрээд гүйлгэдэг хэвээр.
    """
    row_px, header_px, padding_px = 35, 38, 3
    rows_to_show = max(n_rows, min_rows)
    return min(header_px + row_px * rows_to_show + padding_px, max_px)


# =======================================================================================
# 3. EXCEL УНШИХ / ТУЛГАХ ЛОГИК
# =======================================================================================
def compute_actual(df: pd.DataFrame) -> pd.DataFrame:
    """
    Бодит      = (Өглөө + Хүргэлт) − Орой        → өдөрт хорогдсон нийт тоо
    Тооц.борл  = Бодит − Хаягдал − Дотоод        → зарагдсан байх ёстой тоо
    """
    df = ensure_count_cols(df)
    df["Бодит"] = (df["Өглөө"] + df["Хүргэлт"]) - df["Орой"]
    df["Тооц.борл"] = df["Бодит"] - df["Хаягдал"] - df["Дотоод"]
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


def find_header_row(raw_df: pd.DataFrame, keywords, max_scan: int = 40) -> int:
    """Эхний хэдэн мөрнөөс хамгийн олон түлхүүр үгтэй давхцсаныг толгой мөр гэж үзнэ."""
    best_row, best_score = 0, -1
    for i in range(min(max_scan, len(raw_df))):
        row_vals = raw_df.iloc[i].fillna("").astype(str).str.lower().tolist()
        score = sum(1 for v in row_vals for kw in keywords if kw in v)
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def clean_code(val) -> str:
    """Тоон код 465.0 маягаар унших асуудлыг засаж '465' болгоно."""
    s = str(val).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
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
    ЗАСВАР: багануудыг DataFrame(dict)-ээр нэг дор үүсгэнэ. Өмнө нь хоосон
    DataFrame дээр скаляр онооход бүх Код NaN болж байсан.
    """
    uploaded_file.seek(0)
    raw = pd.read_excel(uploaded_file, engine="openpyxl", header=None)
    header_row = find_header_row(
        raw, ["item", "qty", "sold", "код", "code", "id", "нэр", "name", "plu", "price", "cost"]
    )

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

    idx = df_raw.index
    out = pd.DataFrame({
        "Код": (df_raw[code_col].apply(clean_code) if code_col is not None
                else pd.Series("", index=idx)),
        "Нэр": (df_raw[name_col].astype(str).str.strip() if name_col is not None
                else pd.Series("", index=idx)),
        "Систем": pd.to_numeric(df_raw[qty_col], errors="coerce"),
    })

    # --- Дэд нийлбэр / хоосон (спэйсэр) мөрүүдийг хасах ---
    out["Нэр"] = out["Нэр"].replace({"nan": "", "None": ""}).str.strip()
    out["Код"] = out["Код"].replace({"nan": "", "None": ""}).str.strip()
    out = out.dropna(subset=["Систем"])

    # Код багана байгаа бол дэд нийлбэрийн мөрөнд код байдаггүй тул
    # кодгүй мөрийг шууд хасна. Кодгүй файлд Нэрээр шүүнэ.
    if code_col is not None:
        out = out[out["Код"] != ""]
    else:
        out = out[out["Нэр"] != ""]

    # "Total / Subtotal / Нийт / Дүн" гэх мэт нийлбэрийн мөрийг нэрээр нь хасна
    total_pat = r"(?i)\b(?:sub\s*total|subtotal|total|grand\s*total)\b|нийт|дүн|бүгд"
    out = out[~out["Нэр"].str.contains(total_pat, regex=True, na=False)]

    if out.empty:
        raise ValueError(
            "Барааны мөр олдсонгүй. Excel файл дэд-нийлбэрийн мөр л агуулсан "
            "эсвэл багана буруу таарсан байж болзошгүй."
        )

    # Нэг бараа тайланд хэд хэдэн бүлэгт давхардаж гарч ирдэг тул нэгтгэнэ.
    out = out.groupby(["Код", "Нэр"], as_index=False)["Систем"].sum()
    return out


def list_excel_sheets(uploaded_file):
    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    return xls.sheet_names


def guess_default_sheet(sheet_names):
    """Тооллоготой холбоотой нэртэй хуудсыг эрхэмлэж, кассын тайлан зэргийг алгасна."""
    for s in sheet_names:
        low = s.lower()
        if any(k in low for k in ("тооллого", "toollogo", "inventory", "count")):
            return s
    for s in sheet_names:
        low = s.lower()
        if any(k in low for k in ("касс", "cash", "хаалт")):
            continue
        return s
    return sheet_names[0]


def read_excel_with_header_guess(uploaded_file, sheet_name):
    """Толгой мөр эхний мөрөнд биш байх тохиолдлыг автоматаар илрүүлнэ."""
    uploaded_file.seek(0)
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl", header=None)
    keywords = ["код", "№", "id", "plu", "нэр", "name", "өглөө", "morning", "хүргэлт", "орлого",
                "delivery", "орой", "evening", "гаралт", "хаягдал", "дотоод", "систем",
                "зөрүү", "тайлбар", "comment", "note"]
    best_row = find_header_row(raw, keywords, max_scan=40)
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl", header=best_row)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=0, how="all")
    return df, best_row


def guess_count_column_defaults(df: pd.DataFrame) -> dict:
    """Тооллогын баганууд эхлэн таамаглана."""
    cols = list(df.columns)
    g = {
        "code": find_col(cols, ["код", "№", "no", "plu", "id"]),
        "name": find_col(cols, ["нэр", "name", "бараа"]),
        "morning": find_col(cols, ["өглөө", "morning"]),
        "delivery": find_col(cols, ["хүргэлт", "орлого", "delivery"]),
        "evening": find_col(cols, ["орой", "evening"]),
        "waste": find_col(cols, ["хаягдал", "гэмтэл", "waste"]),
        "internal": find_col(cols, ["дотоод", "ажилтан", "internal"]),
        "note": find_col(cols, ["тайлбар", "comment", "note"]),
    }
    # "Өглөө" баганад нэр байхгүй тохиолдол — "Хүргэлт"-ийн зүүн талын
    # тоон утгатай баганыг таамаглал болгоно.
    if g["morning"] is None and g["delivery"] is not None:
        idx = cols.index(g["delivery"])
        if idx > 0:
            candidate = cols[idx - 1]
            if pd.to_numeric(df[candidate], errors="coerce").notna().sum() > 0:
                g["morning"] = candidate
    return g


def reconcile(df_count: pd.DataFrame, df_system: pd.DataFrame,
              price_map: dict, fuzzy_threshold: int):
    """
    Код-оор эхлээд тулгана, олдохгүй бол Fuzzy search-ээр нэрээр тулгана.

    Чухал зарчим:
      * Тулгагдаагүй барааг 0 БИШ, NaN гэж үлдээнэ (0 зарагдсантай хольж болохгүй)
      * Нэг системийн барааг хоёр тооллогын мөрөнд давхар оноохгүй
      * Системд зарагдсан ч тооллогод огт ороогүй барааг тусад нь буцаана

    Буцаах: (тулгасан хүснэгт, тооллогод ороогүй системийн бараанууд)
    """
    df = compute_actual(df_count)

    code_map, name_map = {}, {}
    for _, r in df_system.iterrows():
        code = str(r.get("Код", "")).strip()
        name = str(r.get("Нэр", "")).strip()
        if code and code.lower() != "nan":
            code_map[code] = float(r["Систем"])
        if name and name.lower() != "nan":
            name_map[name] = float(r["Систем"])
    system_names = list(name_map.keys())

    used_codes, used_names = set(), set()
    system_qty_list, method_list = [], []

    for _, row in df.iterrows():
        code = str(row.get("Код", "")).strip()
        name = str(row.get("Нэр", "")).strip()
        sys_qty, method = np.nan, "Олдсонгүй"

        # 1) Код-оор тулгах (хамгийн найдвартай)
        if code and code in code_map:
            sys_qty = code_map[code]
            method = "Код"
            used_codes.add(code)
            for nm, q in name_map.items():
                if nm == name:
                    used_names.add(nm)
        # 2) Fuzzy search — нэрээр тулгах
        elif name and system_names:
            best = process.extractOne(name, system_names, scorer=fuzz.token_sort_ratio)
            if best and best[1] >= fuzzy_threshold:
                matched_name = best[0]
                if matched_name in used_names:
                    method = f"⚠️ Давхар таарсан → {matched_name}"
                else:
                    sys_qty = name_map[matched_name]
                    method = f"Нэр {int(best[1])}% → {matched_name}"
                    used_names.add(matched_name)
                    for c, q in code_map.items():
                        if q == sys_qty and c not in used_codes:
                            pass  # код тодорхойгүй тул хөндөхгүй

        system_qty_list.append(sys_qty)
        method_list.append(method)

    df["Систем"] = system_qty_list
    df["Тулгалт"] = method_list
    df["Зөрүү"] = df["Тооц.борл"] - df["Систем"]

    # Мөнгөн дүн
    df["Зөрүү ₮"] = df.apply(
        lambda r: r["Зөрүү"] * price_map.get(str(r["Код"]).strip(), 0.0)
        if pd.notna(r["Зөрүү"]) else np.nan,
        axis=1,
    )

    # --- Системд байгаа ч тооллогод ороогүй бараа ---
    missing_rows = []
    for _, r in df_system.iterrows():
        code = str(r.get("Код", "")).strip()
        name = str(r.get("Нэр", "")).strip()
        if code and code in used_codes:
            continue
        if name and name in used_names:
            continue
        if float(r["Систем"]) == 0:
            continue
        missing_rows.append({
            "Код": code,
            "Нэр": name,
            "Систем": float(r["Систем"]),
            "Дүн ₮": float(r["Систем"]) * price_map.get(code, 0.0),
        })
    df_missing = pd.DataFrame(missing_rows, columns=["Код", "Нэр", "Систем", "Дүн ₮"])

    return df, df_missing


def find_duplicate_codes(df: pd.DataFrame) -> pd.DataFrame:
    """Тооллогын хүснэгтэд нэг код хоёр мөрөөр орсныг илрүүлнэ."""
    tmp = df.copy()
    tmp["Код"] = tmp["Код"].astype(str).str.strip()
    tmp = tmp[tmp["Код"] != ""]
    counts = tmp.groupby("Код").size()
    dups = counts[counts > 1]
    if dups.empty:
        return pd.DataFrame(columns=["Код", "Нэр", "Давтагдсан"])
    rows = []
    for code, n in dups.items():
        names = ", ".join(sorted(set(tmp[tmp["Код"] == code]["Нэр"].astype(str))))
        rows.append({"Код": code, "Нэр": names, "Давтагдсан": int(n)})
    return pd.DataFrame(rows)


# =======================================================================================
# 4. ХҮСНЭГТ ХАРУУЛАХ (pandas хувилбар хамаарахгүй, аюулгүй)
# =======================================================================================
def _styler_map(sty, func, subset):
    """pandas 2.1-ээс өмнө .applymap, дараа нь .map."""
    if hasattr(sty, "map"):
        return sty.map(func, subset=subset)
    return sty.applymap(func, subset=subset)


def make_diff_colorizer(tolerance: float = 0.0):
    def color_diff(val):
        try:
            v = float(val)
        except (ValueError, TypeError):
            return ""
        if pd.isna(v):
            return "color:#888; background-color:#f4f4f4;"
        if abs(v) <= tolerance:
            return "color:#555;"
        if v < 0:
            return "color:#c0392b; font-weight:700; background-color:#fdecea;"
        return "color:#1e8449; font-weight:700; background-color:#eafaf1;"
    return color_diff


def styled_table(df: pd.DataFrame, tolerance: float = 0.0):
    """
    ЗАСВАР: format хийхээс өмнө тоон баганыг ЗААВАЛ цэвэрлэнэ.
    Үгүй бол архиваас ирсэн холимог төрөлтэй багана ValueError өгч хүснэгт нээгдэхгүй.
    """
    order = ["Код", "Нэр", "Өглөө", "Хүргэлт", "Орой", "Хаягдал", "Дотоод",
             "Бодит", "Тооц.борл", "Систем", "Зөрүү", "Зөрүү ₮", "Тулгалт", "Тайлбар", "Огноо"]
    cols = [c for c in order if c in df.columns]
    view = df[cols].copy() if cols else df.copy()

    numeric_here = [c for c in NUM_COLS + CALC_COLS if c in view.columns]
    for c in numeric_here:
        view[c] = pd.to_numeric(view[c], errors="coerce")

    sty = view.style
    for diff_col in ["Зөрүү", "Зөрүү ₮"]:
        if diff_col in view.columns:
            sty = _styler_map(sty, make_diff_colorizer(tolerance), [diff_col])

    fmt = {c: "{:,.1f}" for c in numeric_here if c != "Зөрүү ₮"}
    if "Зөрүү ₮" in view.columns:
        fmt["Зөрүү ₮"] = "{:,.0f}"
    sty = sty.format(fmt, na_rep="—")
    return sty


def safe_table(df: pd.DataFrame, tolerance: float = 0.0):
    """Загварлаж чадахгүй бол ч гэсэн хүснэгтийг заавал харуулна."""
    try:
        st.dataframe(styled_table(df, tolerance), use_container_width=True, hide_index=True)
    except Exception:
        st.dataframe(df, use_container_width=True, hide_index=True)


def df_to_excel_bytes(sheets: dict) -> bytes:
    """sheets = {'хуудасны нэр': DataFrame, ...}"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        wrote = False
        for name, d in sheets.items():
            if d is None or len(d) == 0:
                continue
            safe_name = str(name)[:31].replace("/", "-").replace("\\", "-").replace(":", "-")
            d.to_excel(writer, index=False, sheet_name=safe_name)
            wrote = True
        if not wrote:
            pd.DataFrame({"Мэдээлэл": ["Өгөгдөл байхгүй"]}).to_excel(
                writer, index=False, sheet_name="Хоосон")
    return buf.getvalue()


# =======================================================================================
# 5. SESSION STATE ЭХЛҮҮЛЭХ
# =======================================================================================
if "count_df" not in st.session_state:
    saved_current = load_json(PATH_CURRENT, None)
    if saved_current and saved_current.get("items"):
        st.session_state.count_df = ensure_count_cols(pd.DataFrame(saved_current["items"]))
    else:
        m = load_master()
        if not m.empty:
            rows = []
            for _, r in m.iterrows():
                row = empty_count_row()
                row["Код"] = r["Код"]
                row["Нэр"] = r["Нэр"]
                rows.append(row)
            st.session_state.count_df = ensure_count_cols(pd.DataFrame(rows))
        else:
            st.session_state.count_df = ensure_count_cols(pd.DataFrame([empty_count_row()]))

st.session_state.setdefault("reconciled_df", None)
st.session_state.setdefault("missing_df", None)
st.session_state.setdefault("pending_master_rerun", False)

render_flash()


# =======================================================================================
# 6. SIDEBAR — Тохиргоо, Нөөцлөлт, Заавар
# =======================================================================================
with st.sidebar:
    st.markdown("### ☕ Caffe Bene")
    st.caption("Тооллого & Тулгалтын систем v2")
    st.markdown("---")

    st.markdown("#### ⚙️ Тохиргоо")
    fuzzy_threshold = st.slider(
        "Нэрээр тулгах босго (%)", min_value=50, max_value=100, value=80, step=5,
        help="Өндөр байх тусам зөвхөн маш ойрхон нэрсийг тулгана. Буруу тулгалт "
             "олон гарвал энэ тоог нэмэгдүүлнэ.",
    )
    tolerance = st.number_input(
        "Зөвшөөрөх зөрүү (±)", min_value=0.0, value=0.0, step=0.5,
        help="Энэ хязгаарт багтсан жижиг зөрүүг улаанаар тэмдэглэхгүй.",
    )
    st.caption(f"Fuzzy сан: `{FUZZ_LIB}`")

    st.markdown("---")
    st.markdown("#### 💾 Нөөцлөх / Сэргээх")
    st.caption("⚠️ Онлайн сервер дээр ажиллуулж байвал файлууд арилах эрсдэлтэй. "
               "Өдөр бүр нөөцөө татаж авахыг зөвлөе.")

    backup_payload = {
        "exported_at": datetime.now().isoformat(),
        "master": load_json(PATH_MASTER, []),
        "history": load_json(PATH_HISTORY, []),
        "deleted": load_json(PATH_DELETED, []),
        "current": load_json(PATH_CURRENT, {}),
    }
    st.download_button(
        "⬇️ Бүх өгөгдлийг нөөцлөх (.json)",
        data=json.dumps(backup_payload, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name=f"CaffeBene_backup_{date.today():%Y%m%d}.json",
        mime="application/json",
        use_container_width=True,
    )

    restore_file = st.file_uploader("Нөөцөөс сэргээх", type=["json"], key="restore_upload")
    if restore_file is not None:
        if st.button("♻️ Сэргээх (одоогийн өгөгдлийг дарна)", use_container_width=True):
            try:
                payload = json.loads(restore_file.getvalue().decode("utf-8"))
                save_json(PATH_MASTER, payload.get("master", []))
                save_json(PATH_HISTORY, payload.get("history", []))
                save_json(PATH_DELETED, payload.get("deleted", []))
                save_json(PATH_CURRENT, payload.get("current", {}))
                st.session_state.pop("count_df", None)
                flash("Нөөцөөс амжилттай сэргээлээ.", "♻️")
                st.rerun()
            except Exception as e:
                st.error(f"Сэргээхэд алдаа гарлаа: {e}")

    st.markdown("---")
    st.markdown(
        """
        **Ажиллах дараалал:**
        1. 📝 **ТООЛЛОГО** — Ө/Х/О, хаягдал, дотоод хэрэглээг шивнэ.
        2. Системийн Excel (`Qty Sold`) upload хийж **Тулгалт хийх**.
        3. Зөрүүг шалгаад **Архивлах**.
        4. 📊 **АРХИВ** — сараар харах, тренд шинжлэх, Excel татах.
        5. ⚙️ **БАРААНЫ САН** — код / нэр / үнэ бүртгэх.
        """
    )
    st.caption(f"Өгөгдлийн сан: `{DATA_DIR}/`")
    st.caption("© Caffe Bene — Дотоод хэрэглээний систем")


# =======================================================================================
# 7. HEADER
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

tab1, tab2, tab3 = st.tabs(["📝 ТООЛЛОГО", "📊 АРХИВ", "⚙️ БАРААНЫ САН"])


# =======================================================================================
# TAB 1 — ТООЛЛОГО
# =======================================================================================
with tab1:
    st.subheader("📝 Өдрийн тооллого")

    master_df = load_master()
    code_to_name = dict(zip(master_df["Код"], master_df["Нэр"])) if not master_df.empty else {}
    price_map = (dict(zip(master_df["Код"], master_df["Үнэ"]))
                 if not master_df.empty else {})

    c_date, c_carry = st.columns([2, 2])
    with c_date:
        count_date = st.date_input("Тооллогын огноо", value=date.today())

    # ---------------------------------------------------------------------------
    # Өмнөх өдрийн Орой → өнөөдрийн Өглөө
    # ---------------------------------------------------------------------------
    with c_carry:
        st.write("")
        if st.button("⏪ Өмнөх өдрийн Орой → Өглөө болгон татах", use_container_width=True):
            history_all = load_json(PATH_HISTORY, [])
            prior = [r for r in history_all if str(r.get("date", "")) < count_date.strftime("%Y-%m-%d")]
            if not prior:
                st.warning("Өмнөх өдрийн архивласан тооллого олдсонгүй.")
            else:
                prior.sort(key=lambda r: str(r.get("date", "")))
                last = prior[-1]
                prev_items = pd.DataFrame(last.get("items", []))
                if "Орой" not in prev_items.columns:
                    st.warning("Өмнөх тайланд 'Орой' багана байхгүй байна.")
                else:
                    prev_map = {}
                    for _, r in prev_items.iterrows():
                        key = str(r.get("Код", "")).strip() or str(r.get("Нэр", "")).strip()
                        if key:
                            prev_map[key] = pd.to_numeric(r.get("Орой", 0), errors="coerce")
                    cur = ensure_count_cols(st.session_state.count_df)
                    filled = 0
                    for i in cur.index:
                        key = str(cur.at[i, "Код"]).strip() or str(cur.at[i, "Нэр"]).strip()
                        if key in prev_map and pd.notna(prev_map[key]):
                            cur.at[i, "Өглөө"] = float(prev_map[key])
                            filled += 1
                    set_count_df(cur)
                    flash(f"{last['date']} өдрийн Орой-оос {filled} мөрийн Өглөө бөглөгдлөө.", "⏪")
                    st.rerun()

    # ---------------------------------------------------------------------------
    # Excel файлаас ачаалах
    # ---------------------------------------------------------------------------
    with st.expander("📥 Тооллогыг Excel файлаас ачаалах (гараар шивэхийн оронд)", expanded=False):
        st.caption(
            "Ажлын Excel файлаа upload хийгээд баганыг доор тохируулна. Толгой мөр, "
            "баганын байршил ямар ч байсан автоматаар таамаглаж, шаардлагатай бол "
            "гараар засах боломжтой."
        )
        count_file = st.file_uploader("Тооллогын Excel файл", type=["xlsx", "xls"], key="count_upload")

        if count_file is not None:
            try:
                sheet_names = list_excel_sheets(count_file)
                if len(sheet_names) > 1:
                    default_sheet = guess_default_sheet(sheet_names)
                    sel_sheet = st.selectbox(
                        "Хуудас (Sheet) сонгох", sheet_names,
                        index=sheet_names.index(default_sheet), key="count_sheet_sel",
                    )
                else:
                    sel_sheet = sheet_names[0]

                df_import, header_row_idx = read_excel_with_header_guess(count_file, sel_sheet)
                st.caption(f"'{sel_sheet}' хуудасны {header_row_idx + 1}-р мөрийг толгой мөр гэж "
                           f"тооцлоо. Эхний 5 мөр:")
                st.dataframe(df_import.head(5), use_container_width=True, hide_index=True)

                guesses = guess_count_column_defaults(df_import)
                cols_all = list(df_import.columns)
                opts_req = cols_all
                opts_opt = [NONE_OPT] + cols_all

                def _idx(opts, val):
                    try:
                        return opts.index(val)
                    except (ValueError, TypeError):
                        return 0

                st.write("**Баганын харгалзаа:**")
                c1, c2, c3 = st.columns(3)
                code_sel = c1.selectbox("Код багана", opts_req,
                                        index=_idx(opts_req, guesses["code"]), key="imp_code")
                name_sel = c2.selectbox("Нэр багана", opts_req,
                                        index=_idx(opts_req, guesses["name"]), key="imp_name")
                morning_sel = c3.selectbox("Өглөө багана", opts_opt,
                                           index=_idx(opts_opt, guesses["morning"]), key="imp_morning")

                c4, c5, c6 = st.columns(3)
                delivery_sel = c4.selectbox("Хүргэлт багана", opts_opt,
                                            index=_idx(opts_opt, guesses["delivery"]), key="imp_deliv")
                evening_sel = c5.selectbox("Орой багана", opts_opt,
                                           index=_idx(opts_opt, guesses["evening"]), key="imp_even")
                note_sel = c6.selectbox("Тайлбар багана", opts_opt,
                                        index=_idx(opts_opt, guesses["note"]), key="imp_note")

                c7, c8 = st.columns(2)
                waste_sel = c7.selectbox("Хаягдал багана", opts_opt,
                                         index=_idx(opts_opt, guesses["waste"]), key="imp_waste")
                internal_sel = c8.selectbox("Дотоод хэрэглээ багана", opts_opt,
                                            index=_idx(opts_opt, guesses["internal"]), key="imp_intern")

                load_mode = st.radio(
                    "Ачаалах горим",
                    ["Одоогийн хүснэгтийг орлуулах", "Одоогийн хүснэгтэд нэмж холбох"],
                    horizontal=True, key="imp_mode",
                )

                if st.button("📥 Тооллогын хүснэгтэд ачаалах", type="primary", use_container_width=True):
                    def _num_col(sel):
                        if sel == NONE_OPT:
                            return pd.Series(0.0, index=df_import.index)
                        return pd.to_numeric(df_import[sel], errors="coerce").fillna(0.0)

                    def _text_col(sel):
                        if sel == NONE_OPT:
                            return pd.Series("", index=df_import.index)
                        return (df_import[sel].astype(str)
                                .replace({"nan": "", "None": ""}).str.strip())

                    new_df = pd.DataFrame({
                        "Код": df_import[code_sel].apply(clean_code),
                        "Нэр": (df_import[name_sel].astype(str).str.strip()
                                .replace({"nan": "", "None": ""})),
                        "Өглөө": _num_col(morning_sel),
                        "Хүргэлт": _num_col(delivery_sel),
                        "Орой": _num_col(evening_sel),
                        "Хаягдал": _num_col(waste_sel),
                        "Дотоод": _num_col(internal_sel),
                        "Тайлбар": _text_col(note_sel),
                    })

                    # Код ч, нэр ч байхгүй мөрүүдийг (дэд нийлбэр, хоосон зай) хасах
                    new_df = new_df[(new_df["Нэр"].str.strip() != "") |
                                    (new_df["Код"].str.strip() != "")]
                    new_df = new_df.reset_index(drop=True)

                    if new_df.empty:
                        st.error("Ачаалах мөр олдсонгүй. Багануудаа шалгана уу.")
                    else:
                        if load_mode == "Одоогийн хүснэгтийг орлуулах":
                            set_count_df(new_df)
                        else:
                            merged = pd.concat(
                                [ensure_count_cols(st.session_state.count_df), new_df],
                                ignore_index=True)
                            set_count_df(merged)
                        flash(f"{len(new_df)} мөр амжилттай ачааллаа.", "📥")
                        st.rerun()
            except Exception as e:
                st.error(f"Файл уншихад алдаа гарлаа: {e}")

    # ---------------------------------------------------------------------------
    # Тооллогын хүснэгт
    # ---------------------------------------------------------------------------
    st.caption("Мөр бүрт Өглөө / Хүргэлт / Орой, мөн Хаягдал болон Дотоод хэрэглээг оруулна уу. "
               "Шинэ мөр нэмэхдээ хүснэгтийн доод хэсгийн **+** товч ашиглана.")

    # Нэрийг мастераас бөглөх — data_editor-т ОРОХЫН ӨМНӨ хийнэ
    # (ингэснээр шивсэн даруйд харагдана)
    base_df = ensure_count_cols(st.session_state.count_df)
    if code_to_name:
        for i in base_df.index:
            c = str(base_df.at[i, "Код"]).strip()
            if c and c in code_to_name and str(base_df.at[i, "Нэр"]).strip() == "":
                base_df.at[i, "Нэр"] = code_to_name[c]

    # ЗАСВАР: Код багана нь SelectboxColumn БАЙЖ БОЛОХГҮЙ.
    # Мастерт байхгүй код орж ирвэл Streamlit алдаа өгч, нүднүүдийг хоослодог.
    edited_df = st.data_editor(
        base_df,
        num_rows="dynamic",
        use_container_width=True,
        height=editor_height(len(base_df)),
        key="count_editor",
        column_config={
            "Код": st.column_config.TextColumn("Код (PLU)", width="small"),
            "Нэр": st.column_config.TextColumn("Барааны нэр", width="medium"),
            "Өглөө": st.column_config.NumberColumn("Өглөө (Ө)", step=1.0, format="%.1f"),
            "Хүргэлт": st.column_config.NumberColumn("Хүргэлт (Х)", step=1.0, format="%.1f"),
            "Орой": st.column_config.NumberColumn("Орой (О)", step=1.0, format="%.1f"),
            "Хаягдал": st.column_config.NumberColumn("Хаягдал", step=1.0, format="%.1f"),
            "Дотоод": st.column_config.NumberColumn("Дотоод хэрэглээ", step=1.0, format="%.1f"),
            "Тайлбар": st.column_config.TextColumn("Тайлбар", width="large"),
        },
        hide_index=True,
    )
    if len(base_df) > 20:
        st.caption(f"📋 Нийт **{len(base_df)}** мөр ачаалагдсан байна "
                   f"(доошоо гүйлгээд бүгдийг харах боломжтой).")

    edited_df = ensure_count_cols(edited_df)
    st.session_state.count_df = edited_df

    calc_df = compute_actual(edited_df)

    # --- Шалгалтууд ---
    dups = find_duplicate_codes(edited_df)
    if not dups.empty:
        st.warning(f"⚠️ {len(dups)} код давхардсан байна — давхар тоологдох эрсдэлтэй.")
        st.dataframe(dups, use_container_width=True, hide_index=True)

    neg = calc_df[calc_df["Тооц.борл"] < 0]
    if not neg.empty:
        st.warning(f"⚠️ {len(neg)} мөрийн тооцоот борлуулалт сөрөг байна "
                   f"(Орой > Өглөө+Хүргэлт). Шивэлт шалгана уу.")
        st.dataframe(neg[["Код", "Нэр", "Өглөө", "Хүргэлт", "Орой", "Тооц.борл"]],
                     use_container_width=True, hide_index=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Мөрийн тоо", len(calc_df))
    m2.metric("Нийт Бодит (Ө+Х−О)", f"{calc_df['Бодит'].sum():,.1f}")
    m3.metric("Хаягдал + Дотоод", f"{(calc_df['Хаягдал'] + calc_df['Дотоод']).sum():,.1f}")
    m4.metric("Тооцоот борлуулалт", f"{calc_df['Тооц.борл'].sum():,.1f}")

    with st.expander("🔎 Тооцоолсон дүн (тулгалтын өмнөх)", expanded=False):
        st.dataframe(
            calc_df[["Код", "Нэр", "Өглөө", "Хүргэлт", "Орой", "Хаягдал",
                     "Дотоод", "Бодит", "Тооц.борл", "Тайлбар"]],
            use_container_width=True, hide_index=True,
        )

    col_save1, col_save2 = st.columns(2)
    with col_save1:
        if st.button("💾 Түр хадгалах (Draft)", use_container_width=True):
            save_json(PATH_CURRENT, {
                "date": count_date.strftime("%Y-%m-%d"),
                "items": edited_df.to_dict(orient="records"),
                "saved_at": datetime.now().isoformat(),
            })
            flash("Түр хадгаллаа. Дараа нэвтрэхэд энэ өгөгдөл сэргэнэ.", "💾")
            st.rerun()
    with col_save2:
        if st.button("🗑️ Хүснэгтийг цэвэрлэх", use_container_width=True):
            set_count_df(pd.DataFrame([empty_count_row()]))
            flash("Хүснэгтийг цэвэрлэлээ.", "🗑️")
            st.rerun()

    # ---------------------------------------------------------------------------
    # Тулгалт
    # ---------------------------------------------------------------------------
    st.divider()
    st.subheader("🔄 Системийн Excel-тэй тулгах")
    st.caption("Excel файл нь `Код`/`ID`, `Нэр` болон **`Qty Sold`** баганатай байх ёстой.")

    sys_file = st.file_uploader("Системийн борлуулалтын Excel файл",
                                type=["xlsx", "xls"], key="sys_upload")

    if sys_file is not None:
        try:
            df_system = parse_system_excel(sys_file)
            st.success(f"Системийн файлаас {len(df_system)} мөр уншлаа "
                       f"(нийт {df_system['Систем'].sum():,.0f} ширхэг).")
            if st.button("⚖️ Тулгалт хийх", type="primary", use_container_width=True):
                rec, miss = reconcile(edited_df, df_system, price_map, fuzzy_threshold)
                st.session_state.reconciled_df = rec
                st.session_state.missing_df = miss
                st.rerun()
        except Exception as e:
            st.error(f"Файл уншихад алдаа гарлаа: {e}")

    if st.session_state.reconciled_df is not None:
        rdf = st.session_state.reconciled_df
        mdf = st.session_state.missing_df

        st.write("**Тулгалтын үр дүн** — 🟥 Дутсан | 🟩 Илүүдсэн | ⬜ Тулгагдаагүй")
        safe_table(rdf, tolerance)

        short_n = int((rdf["Зөрүү"] < -tolerance).sum())
        over_n = int((rdf["Зөрүү"] > tolerance).sum())
        ok_n = int((rdf["Зөрүү"].abs() <= tolerance).sum())
        nomatch_n = int(rdf["Систем"].isna().sum())

        d1, d2, d3, d4 = st.columns(4)
        d1.metric("🟥 Дутсан", short_n)
        d2.metric("🟩 Илүүдсэн", over_n)
        d3.metric("✅ Тохирсон", ok_n)
        d4.metric("⬜ Тулгагдаагүй", nomatch_n)

        loss = rdf.loc[rdf["Зөрүү"] < 0, "Зөрүү ₮"].sum(skipna=True)
        if pd.notna(loss) and loss != 0:
            st.error(f"💸 Дутагдлын нийт дүн: **{abs(loss):,.0f} ₮** "
                     f"(мастер жагсаалтад үнэ оруулсан бараануудаар)")

        if nomatch_n > 0:
            with st.expander(f"⬜ Тулгагдаагүй {nomatch_n} бараа", expanded=False):
                st.caption("Эдгээрийг систем дээр олж чадсангүй. Код буруу, эсвэл нэр "
                           "хэт өөр байж болно. Мастер санд кодыг нь зөв бүртгэвэл дараагийн "
                           "удаа автоматаар тулгагдана.")
                st.dataframe(rdf[rdf["Систем"].isna()][["Код", "Нэр", "Тооц.борл", "Тулгалт"]],
                             use_container_width=True, hide_index=True)

        if mdf is not None and not mdf.empty:
            st.warning(f"⚠️ Системд зарагдсан ч тооллогод ОРООГҮЙ {len(mdf)} бараа байна.")
            st.caption("Энэ жагсаалт хамгийн анхаарал татах ёстой хэсэг — тоологдоогүй "
                       "бараа бүртгэлээс бүрэн гадуур үлдэж байна гэсэн үг.")
            st.dataframe(mdf, use_container_width=True, hide_index=True)

        # --- Архивлах ---
        st.write("**📷 Нотлох баримт хавсаргах (заавал биш)**")
        evidence_photos = st.file_uploader(
            "Кассын хуудас, гар бичмэл тооллогын зураг",
            type=["png", "jpg", "jpeg"], accept_multiple_files=True,
            key="evidence_photo_upload",
        )

        if st.button("📦 Архивлах (Тулгалтыг баталгаажуулж хадгалах)",
                     type="primary", use_container_width=True):
            record_id = str(uuid.uuid4())

            photo_paths = []
            if evidence_photos:
                record_photo_dir = os.path.join(PATH_PHOTOS, record_id)
                os.makedirs(record_photo_dir, exist_ok=True)
                for i, photo in enumerate(evidence_photos):
                    ext = os.path.splitext(photo.name)[1] or ".jpg"
                    fpath = os.path.join(record_photo_dir, f"{i+1:02d}{ext}")
                    with open(fpath, "wb") as f:
                        f.write(photo.getbuffer())
                    photo_paths.append(fpath)

            history = load_json(PATH_HISTORY, [])
            record = {
                "id": record_id,
                "date": count_date.strftime("%Y-%m-%d"),
                "archived_at": datetime.now().isoformat(),
                "items": json.loads(rdf.to_json(orient="records")),  # NaN → null
                "missing": json.loads(mdf.to_json(orient="records")) if mdf is not None else [],
                "photos": photo_paths,
                "settings": {"fuzzy_threshold": fuzzy_threshold, "tolerance": tolerance},
            }
            history.append(record)
            save_json(PATH_HISTORY, history)

            save_json(PATH_CURRENT, {"date": "", "items": [], "saved_at": ""})
            set_count_df(pd.DataFrame([empty_count_row()]))
            flash(f"{count_date:%Y-%m-%d} өдрийн тооллого архивлагдлаа!"
                  + (f" ({len(photo_paths)} зурагтай)" if photo_paths else ""), "📦")
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
        hist_meta = pd.DataFrame([
            {"id": r["id"], "Огноо": r.get("date", ""),
             "Архивласан": str(r.get("archived_at", ""))[:19].replace("T", " "),
             "Мөрийн тоо": len(r.get("items", []))}
            for r in history
        ])
        hist_meta["Сар"] = pd.to_datetime(hist_meta["Огноо"], errors="coerce").dt.strftime("%Y-%m")

        months = ["Бүгд"] + sorted(hist_meta["Сар"].dropna().unique().tolist(), reverse=True)
        sel_month = st.selectbox("📅 Сараар шүүх", months)

        filtered = hist_meta if sel_month == "Бүгд" else hist_meta[hist_meta["Сар"] == sel_month]
        st.dataframe(filtered[["Огноо", "Архивласан", "Мөрийн тоо"]],
                     use_container_width=True, hide_index=True)

        record_options = {
            f'{r.get("date","?")} — {r["id"][:8]}': r["id"]
            for r in sorted(history, key=lambda x: str(x.get("date", "")), reverse=True)
            if sel_month == "Бүгд" or str(r.get("date", ""))[:7] == sel_month
        }

        if record_options:
            sel_label = st.selectbox("Дэлгэрэнгүй харах тайлан", list(record_options.keys()))
            sel_id = record_options[sel_label]
            sel_record = next(r for r in history if r["id"] == sel_id)
            rdf = pd.DataFrame(sel_record.get("items", []))
            mdf = pd.DataFrame(sel_record.get("missing", []))

            st.write(f"### 🧾 {sel_record.get('date','?')} өдрийн тайлан")
            if not rdf.empty:
                safe_table(rdf, tolerance)

                if "Зөрүү" in rdf.columns:
                    z = pd.to_numeric(rdf["Зөрүү"], errors="coerce")
                    zt = pd.to_numeric(rdf.get("Зөрүү ₮", pd.Series(dtype=float)), errors="coerce")
                    a1, a2, a3 = st.columns(3)
                    a1.metric("🟥 Дутсан", int((z < 0).sum()))
                    a2.metric("🟩 Илүүдсэн", int((z > 0).sum()))
                    a3.metric("💸 Дутагдлын дүн",
                              f"{abs(zt[z < 0].sum(skipna=True)):,.0f} ₮" if not zt.empty else "—")
            else:
                st.info("Энэ тайланд мөр байхгүй байна.")

            if not mdf.empty:
                st.warning(f"⚠️ Тооллогод ороогүй {len(mdf)} бараа")
                st.dataframe(mdf, use_container_width=True, hide_index=True)

            photo_paths = [p for p in sel_record.get("photos", []) if os.path.exists(p)]
            if photo_paths:
                st.write(f"**📷 Хавсаргасан зураг ({len(photo_paths)}):**")
                photo_cols = st.columns(min(3, len(photo_paths)))
                for i, p in enumerate(photo_paths):
                    with photo_cols[i % len(photo_cols)]:
                        show_image(p)

            c1, c2 = st.columns(2)
            with c1:
                sheets = {"Бүгд": rdf}
                if "Зөрүү" in rdf.columns:
                    z = pd.to_numeric(rdf["Зөрүү"], errors="coerce")
                    sheets["Дутсан"] = rdf[z < 0]
                    sheets["Илүүдсэн"] = rdf[z > 0]
                    sheets["Тулгагдаагүй"] = rdf[z.isna()]
                sheets["Тоологдоогүй"] = mdf
                st.download_button(
                    "⬇️ Excel-ээр татах (олон хуудастай)",
                    data=df_to_excel_bytes(sheets),
                    file_name=f"CaffeBene_tailan_{sel_record.get('date','')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with c2:
                if st.button("🗑️ Устгах (Хогийн саванд шилжүүлэх)", use_container_width=True):
                    history = [r for r in history if r["id"] != sel_id]
                    sel_record["deleted_at"] = datetime.now().isoformat()
                    deleted.append(sel_record)
                    save_json(PATH_HISTORY, history)
                    save_json(PATH_DELETED, deleted)
                    flash("Тайланг хогийн саванд шилжүүллээ.", "🗑️")
                    st.rerun()

        # -----------------------------------------------------------------------
        # Сарын нэгтгэл + тренд
        # -----------------------------------------------------------------------
        st.divider()
        st.write("### 📈 Нэгтгэл ба тренд")

        all_items = []
        for r in history:
            if sel_month != "Бүгд" and str(r.get("date", ""))[:7] != sel_month:
                continue
            for item in r.get("items", []):
                item = dict(item)
                item["Огноо"] = r.get("date", "")
                all_items.append(item)

        if not all_items:
            st.caption("Нэгтгэх өгөгдөл алга.")
        else:
            month_df = pd.DataFrame(all_items)
            for c in CALC_COLS + NUM_COLS:
                if c in month_df.columns:
                    month_df[c] = pd.to_numeric(month_df[c], errors="coerce")

            if "Зөрүү" in month_df.columns:
                month_df["Код"] = month_df.get("Код", "").astype(str)
                month_df["Нэр"] = month_df.get("Нэр", "").astype(str)

                agg_spec = {}
                for src, dst in [("Тооц.борл", "Тооц.борл"), ("Бодит", "Бодит"),
                                 ("Систем", "Систем"), ("Зөрүү", "Зөрүү"), ("Зөрүү ₮", "Зөрүү ₮")]:
                    if src in month_df.columns:
                        agg_spec[dst] = (src, "sum")

                # Код+Нэрээр бүлэглэнэ (зөвхөн нэрээр бүлэглэвэл ижил нэртэй
                # өөр код нийлж алдаа өгнө)
                summary = month_df.groupby(["Код", "Нэр"], dropna=False).agg(**agg_spec).reset_index()
                summary = summary.sort_values("Зөрүү") if "Зөрүү" in summary.columns else summary

                st.write("**Нэгтгэсэн дүн:**")
                safe_table(summary, tolerance)

                # Тренд: хэдэн өдөр дутсан бэ
                st.write("**Хамгийн тогтмол дутдаг бараа (өдрийн тоогоор):**")
                trend = (month_df[month_df["Зөрүү"] < -tolerance]
                         .groupby(["Код", "Нэр"], dropna=False)
                         .agg(Дутсан_өдөр=("Огноо", "nunique"),
                              Нийт_зөрүү=("Зөрүү", "sum"))
                         .reset_index()
                         .sort_values("Дутсан_өдөр", ascending=False)
                         .head(20))
                if trend.empty:
                    st.caption("Дутагдал бүртгэгдээгүй байна. 👍")
                else:
                    st.dataframe(trend, use_container_width=True, hide_index=True)
                    st.caption("Олон өдөр давтагдсан дутагдал нь санамсаргүй алдаа биш, "
                               "тогтмол асуудал (жор, хэмжээ, эсвэл бүртгэлийн) байх магадлалтай.")

                st.download_button(
                    "⬇️ Нэгтгэлийг Excel-ээр татах",
                    data=df_to_excel_bytes({"Нэгтгэл": summary, "Тренд": trend}),
                    file_name=f"CaffeBene_negtgel_{sel_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.caption("Энэ хугацааны тайлангуудад тулгалтын мэдээлэл алга.")

    # ---- Хогийн сав ----
    st.divider()
    with st.expander(f"🗑️ Хогийн сав ({len(deleted)})", expanded=False):
        if not deleted:
            st.caption("Хогийн сав хоосон байна.")
        else:
            for r in list(deleted):
                cols = st.columns([3, 3, 2])
                cols[0].write(f"**{r.get('date','?')}** — {r['id'][:8]}")
                cols[1].caption(f"Устгасан: {str(r.get('deleted_at','—'))[:19].replace('T',' ')}")
                if cols[2].button("♻️ Сэргээх", key=f"restore_{r['id']}"):
                    deleted = [d for d in deleted if d["id"] != r["id"]]
                    r.pop("deleted_at", None)
                    history.append(r)
                    save_json(PATH_HISTORY, history)
                    save_json(PATH_DELETED, deleted)
                    flash("Тайланг сэргээлээ.", "♻️")
                    st.rerun()

            if st.button("❌ Хогийн савыг бүрмөсөн хоослох"):
                for r in deleted:
                    for p in r.get("photos", []):
                        try:
                            if os.path.exists(p):
                                os.remove(p)
                        except OSError:
                            pass
                save_json(PATH_DELETED, [])
                flash("Хогийн савыг бүрмөсөн хоослолоо.", "❌")
                st.rerun()


# =======================================================================================
# TAB 3 — БАРААНЫ САН (Master list)
# =======================================================================================
with tab3:
    st.subheader("⚙️ Барааны мэдээллийн сан (Master List)")
    st.caption("Код (PLU), Нэр, Нэгж үнэ бүртгэнэ. Үнэ оруулсан бараануудын зөрүү "
               "автоматаар төгрөгөөр тооцогдоно.")

    master_df = load_master()

    with st.form("add_item_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([1, 2, 1, 1])
        new_code = c1.text_input("Код (PLU/ID)")
        new_name = c2.text_input("Барааны нэр")
        new_price = c3.number_input("Нэгж үнэ (₮)", min_value=0.0, step=100.0, value=0.0)
        submitted = c4.form_submit_button("➕ Нэмэх", use_container_width=True)

    # ЗАСВАР: st.rerun()-г формын ГАДНА дуудна (дотор нь дуудвал алдаа өгнө)
    if submitted:
        if not new_code.strip() or not new_name.strip():
            st.error("Код болон нэрийг хоёуланг нь бөглөнө үү.")
        elif new_code.strip() in master_df["Код"].values:
            st.error(f"'{new_code}' код аль хэдийн бүртгэлтэй байна.")
        else:
            new_row = pd.DataFrame([{"Код": new_code.strip(),
                                     "Нэр": new_name.strip(),
                                     "Үнэ": float(new_price)}])
            save_master(pd.concat([master_df, new_row], ignore_index=True))
            flash(f"'{new_name}' ({new_code}) нэмэгдлээ.", "➕")
            st.rerun()

    st.divider()
    st.write("### 📥 Excel-ээр багцаар оруулах")
    bulk_file = st.file_uploader("Код / Нэр / Үнэ баганатай Excel файл",
                                 type=["xlsx", "xls"], key="bulk_master")
    if bulk_file is not None:
        try:
            df_bulk = pd.read_excel(bulk_file, engine="openpyxl")
            df_bulk.columns = [str(c).strip() for c in df_bulk.columns]
            code_col = find_col(df_bulk.columns, ["код", "code", "id", "plu"])
            name_col = find_col(df_bulk.columns, ["нэр", "name", "item", "бараа"])
            price_col = find_col(df_bulk.columns, ["үнэ", "price", "өртөг", "cost"])

            if not code_col or not name_col:
                st.error("Файлд Код болон Нэр багана олдсонгүй.")
            else:
                preview = pd.DataFrame({
                    "Код": df_bulk[code_col].apply(clean_code),
                    "Нэр": df_bulk[name_col].astype(str).str.strip(),
                    "Үнэ": (pd.to_numeric(df_bulk[price_col], errors="coerce").fillna(0.0)
                            if price_col else 0.0),
                })
                preview = preview[preview["Код"].str.strip() != ""]
                st.dataframe(preview, use_container_width=True, hide_index=True)
                if st.button("✅ Мастер жагсаалтад нэгтгэх", type="primary"):
                    merged = pd.concat([master_df, preview], ignore_index=True)
                    merged = merged.drop_duplicates(subset=["Код"], keep="last")
                    save_master(merged)
                    flash(f"{len(preview)} мөр нэгтгэгдлээ.", "✅")
                    st.rerun()
        except Exception as e:
            st.error(f"Файл уншихад алдаа гарлаа: {e}")

    st.divider()
    st.write(f"### 📋 Бүртгэлтэй бараанууд ({len(master_df)})")

    if master_df.empty:
        st.info("Мастер жагсаалт хоосон байна. Дээрх формоор эхний бараагаа нэмнэ үү.")
    else:
        search_q = st.text_input("🔍 Хайх (код эсвэл нэрээр)")
        show_df = master_df.copy()
        if search_q.strip():
            mask = (show_df["Код"].str.contains(search_q, case=False, na=False) |
                    show_df["Нэр"].str.contains(search_q, case=False, na=False))
            show_df = show_df[mask]

        visible_codes = set(show_df["Код"])

        edited_master = st.data_editor(
            show_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=editor_height(len(show_df)),
            key="master_editor",
            column_config={
                "Код": st.column_config.TextColumn("Код (PLU)", width="small"),
                "Нэр": st.column_config.TextColumn("Барааны нэр", width="large"),
                "Үнэ": st.column_config.NumberColumn("Нэгж үнэ (₮)", min_value=0.0,
                                                     step=100.0, format="%.0f"),
            },
        )

        if st.button("💾 Өөрчлөлтийг хадгалах", type="primary", use_container_width=True):
            if search_q.strip():
                unaffected = master_df[~master_df["Код"].isin(visible_codes)]
                final_df = pd.concat([unaffected, edited_master], ignore_index=True)
            else:
                final_df = edited_master
            final_df = final_df.drop_duplicates(subset=["Код"], keep="last")
            save_master(final_df)
            st.session_state.pop("master_editor", None)
            flash("Мастер жагсаалт шинэчлэгдлээ.", "💾")
            st.rerun()

        st.download_button(
            "⬇️ Мастер жагсаалтыг Excel-ээр татах",
            data=df_to_excel_bytes({"Master_List": master_df}),
            file_name="CaffeBene_master_items.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
