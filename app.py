"""Run with: python -m streamlit run app.py"""

import hashlib
import os
import shutil
import threading
import io
import re
import subprocess
import sys
import tempfile
from datetime import date, datetime
from numbers import Real
from pathlib import Path
from copy import copy

from openpyxl import load_workbook

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from openpyxl.utils.datetime import from_excel


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "template.docx"
MASTER_PATH = next(
    (
        APP_DIR / name
        for name in ("Master_Warranty_Template-v1.xlsx", "Master_Warranty_Template-v2.xlsx")
        if (APP_DIR / name).exists()
    ),
    APP_DIR / "Master_Warranty_Template-v2.xlsx",
)
MASTER_SHEET = "Master_Product_Warranty"
MASTER_LOCK = threading.RLock()
REQUIRED_COLUMNS = ("Billing Date", "Customer name", "Billing", "Article")
SERIAL_COLUMNS = ("ManufSerialNumber", "Serial Number")
MASTER_ARTICLE_COLUMNS = (
    "Article Code", "Article", "Article No.", "Article No", "Material", "Material No.",
    "Material No", "Product Code",
)
MASTER_PRODUCT_TERM_COLUMNS = (
    "Product Warranty Term",
    "Warranty Condition", "Warranty Conditions", "Warranty", "Warranty Period",
    "Warranty Terms", "Condition", "Conditions",
)
MASTER_BASE_WARRANTY_COLUMNS = ("Standard Base Warranty",)
MASTER_CATEGORY_COLUMNS = ("Brand / Category", "Category")
MASTER_PERFORMANCE_WARRANTY_COLUMNS = ("Performance Warranty",)
MASTER_PERFORMANCE_TERM_COLUMNS = ("Performance Warranty Term",)
ENGLISH_MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)


def clean_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, Real) and not isinstance(value, bool):
        if float(value).is_integer():
            return str(int(value))
    return str(value).strip()


def find_column(columns, candidates):
    def normalize(name):
        return " ".join(str(name).split()).casefold()

    names = {normalize(column): column for column in columns}
    return next((names[normalize(name)] for name in candidates if normalize(name) in names), None)


def find_prefixed_column(columns, candidates):
    def normalize(name):
        return " ".join(str(name).split()).casefold()

    names = [(normalize(column), column) for column in columns if clean_value(column)]
    for candidate in candidates:
        prefix = normalize(candidate)
        for normalized, column in names:
            if normalized == prefix or normalized.startswith(prefix + " "):
                return column
    return None


def usable_warranty_text(value):
    text = clean_value(value)
    return "" if not text or text == "-" or text.startswith("=") else text


def build_product_warranty(category, years):
    years = clean_value(years)
    if not years or years == "-":
        return ""
    if "solar panel" in clean_value(category).casefold():
        return f"รับประกันถ้าแผงไม่มีไฟจ่ายออก {years} ปี นับแต่วันที่ส่งมอบสินค้า"
    return f"รับประกัน {years} ปี นับแต่วันที่ส่งมอบสินค้า"


def build_performance_warranty(years):
    years = clean_value(years)
    if not years or years == "-":
        return ""
    return f"รับประกันประสิทธิภาพการทำงานของแผง {years} ปี นับแต่วันส่งมอบสินค้า"


def format_date(value):
    if isinstance(value, date) and not pd.isna(value):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, Real) and not isinstance(value, bool) and 1 <= value <= 2958465:
        parsed = from_excel(value)
        if isinstance(parsed, datetime):
            return parsed.strftime("%d.%m.%Y")
    text = clean_value(value)
    for date_format in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, date_format).strftime("%d.%m.%Y")
        except ValueError:
            continue
    raise ValueError(f"Billing Date ไม่ถูกต้อง ({text or 'ว่าง'})")


def format_document_date(value):
    return "\u00a0".join(
        (f"{value.day:02d}", ENGLISH_MONTHS[value.month - 1], str(value.year))
    )


def load_master_warranty(source=MASTER_PATH):
    raw = pd.read_excel(
        source,
        sheet_name=MASTER_SHEET,
        header=None,
        dtype=object,
        keep_default_na=False,
    )
    header_row = next(
        (
            index
            for index, row in raw.iloc[:25].iterrows()
            if find_prefixed_column(row.tolist(), MASTER_ARTICLE_COLUMNS) is not None
        ),
        None,
    )
    if header_row is None:
        raise ValueError("ไม่พบแถวหัวตาราง Article Code ในชีต Master_Product_Warranty")

    master = raw.iloc[header_row + 1:].copy()
    master.columns = [clean_value(value) for value in raw.iloc[header_row]]
    article_col = find_prefixed_column(master.columns, MASTER_ARTICLE_COLUMNS)
    product_term_col = find_prefixed_column(master.columns, MASTER_PRODUCT_TERM_COLUMNS)
    base_warranty_col = find_prefixed_column(master.columns, MASTER_BASE_WARRANTY_COLUMNS)
    category_col = find_prefixed_column(master.columns, MASTER_CATEGORY_COLUMNS)
    performance_warranty_col = find_prefixed_column(
        master.columns, MASTER_PERFORMANCE_WARRANTY_COLUMNS
    )
    performance_term_col = find_prefixed_column(master.columns, MASTER_PERFORMANCE_TERM_COLUMNS)
    if article_col is None or (product_term_col is None and base_warranty_col is None):
        raise ValueError(
            "ต้องมี Article Code และ Product Warranty Term "
            "หรือ Standard Base Warranty ในชีต Master_Product_Warranty"
        )

    terms = {}
    for _, row in master.iterrows():
        article = clean_value(row[article_col])
        if not article:
            continue
        product_term = usable_warranty_text(row[product_term_col]) if product_term_col else ""
        if not product_term and base_warranty_col:
            category = row[category_col] if category_col else ""
            product_term = build_product_warranty(category, row[base_warranty_col])
        performance_term = (
            usable_warranty_text(row[performance_term_col]) if performance_term_col else ""
        )
        if not performance_term and performance_warranty_col:
            performance_term = build_performance_warranty(row[performance_warranty_col])
        term = "\n".join(value for value in (product_term, performance_term) if value)
        if not term:
            continue
        if article in terms and terms[article] != term:
            raise ValueError(f"Article {article} มีเงื่อนไขรับประกันซ้ำที่ระบุข้อความต่างกัน")
        terms[article] = term
    return terms



def master_header(ws):
    """Discover the same header row used by load_master_warranty()."""
    for cells in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 25)):
        labels = [clean_value(cell.value) for cell in cells]
        if find_prefixed_column(labels, MASTER_ARTICLE_COLUMNS) is not None:
            header_row = cells[0].row
            return header_row, labels
    raise ValueError("ไม่พบแถวหัวตาราง Article Code ในชีต Master_Product_Warranty")


def master_sheet(wb):
    if MASTER_SHEET not in wb.sheetnames:
        raise ValueError(f"ไม่พบชีต {MASTER_SHEET}")
    return wb[MASTER_SHEET]


def normalized_header(name):
    return " ".join(clean_value(name).split()).casefold()


def preview_master_merge(original_bytes, uploaded_bytes):
    """Stage a non-destructive Excel merge in memory; nothing is saved yet."""
    incoming_terms = load_master_warranty(io.BytesIO(uploaded_bytes))
    if not incoming_terms:
        raise ValueError("ไฟล์ที่อัปโหลดไม่มี Article พร้อมเงื่อนไขรับประกันที่ใช้ได้")
    incoming_wb = load_workbook(io.BytesIO(uploaded_bytes))
    incoming_ws = master_sheet(incoming_wb)
    source_header_row, source_headers = master_header(incoming_ws)
    source_article_header = find_prefixed_column(source_headers, MASTER_ARTICLE_COLUMNS)
    source_article_col = source_headers.index(source_article_header) + 1
    source_lookup = {normalized_header(name): i for i, name in enumerate(source_headers, 1) if name}

    if original_bytes is None:
        # The first upload becomes the durable master, retaining its workbook formatting.
        return uploaded_bytes, [], sorted(incoming_terms), [], incoming_terms

    original_terms = load_master_warranty(io.BytesIO(original_bytes))
    wb = load_workbook(io.BytesIO(original_bytes))
    ws = master_sheet(wb)
    target_header_row, target_headers = master_header(ws)
    target_article_header = find_prefixed_column(target_headers, MASTER_ARTICLE_COLUMNS)
    article_col = target_headers.index(target_article_header) + 1
    target_lookup = {normalized_header(name): i for i, name in enumerate(target_headers, 1) if name}

    # Require the same warranty fields so that the merge cannot silently drop or
    # misinterpret warranty conditions from differently structured workbooks.
    warranty_prefixes = (
        MASTER_PRODUCT_TERM_COLUMNS + MASTER_BASE_WARRANTY_COLUMNS +
        MASTER_CATEGORY_COLUMNS + MASTER_PERFORMANCE_TERM_COLUMNS +
        MASTER_PERFORMANCE_WARRANTY_COLUMNS
    )
    for header in source_headers:
        if header and find_prefixed_column([header], warranty_prefixes):
            if normalized_header(header) not in target_lookup:
                raise ValueError(
                    f"Master หลักไม่มีคอลัมน์ '{header}' กรุณาใช้ไฟล์ที่มีโครงสร้างคอลัมน์ตรงกับ Master หลัก"
                )
    matched = [(source_col, target_lookup[key]) for key, source_col in source_lookup.items() if key in target_lookup]
    if not matched:
        raise ValueError("ไม่พบคอลัมน์ที่ตรงกับ Master หลัก")

    existing = {}
    for row in range(target_header_row + 1, ws.max_row + 1):
        article = clean_value(ws.cell(row, article_col).value)
        if article:
            if article in existing:
                raise ValueError(f"Master หลักมี Article ซ้ำ: {article}")
            existing[article] = row

    seen = set()
    for row in range(source_header_row + 1, incoming_ws.max_row + 1):
        article = clean_value(incoming_ws.cell(row, source_article_col).value)
        if article not in incoming_terms:
            continue
        if article in seen:
            raise ValueError(f"ไฟล์ที่อัปโหลดมี Article ซ้ำ: {article}")
        seen.add(article)
        if article in existing:
            target_row = existing[article]
        else:
            target_row = ws.max_row + 1
            # Retain formatting without copying old article data or formulas.
            if target_row > target_header_row + 1:
                template_row = target_row - 1
                for col in range(1, ws.max_column + 1):
                    src_style = ws.cell(template_row, col)
                    dst_style = ws.cell(target_row, col)
                    if src_style.has_style:
                        dst_style._style = copy(src_style._style)
                    dst_style.number_format = src_style.number_format
            existing[article] = target_row
        for source_col, target_col in matched:
            source_value = incoming_ws.cell(row, source_col).value
            source_name = source_headers[source_col - 1]
            is_warranty_field = bool(find_prefixed_column([source_name], warranty_prefixes))
            # For warranties, empty means clear the old condition. For other fields,
            # preserve prior metadata unless a new value was supplied.
            if is_warranty_field or clean_value(source_value):
                ws.cell(target_row, target_col).value = source_value
        ws.cell(target_row, article_col).value = article

    output = io.BytesIO()
    wb.save(output)
    merged_bytes = output.getvalue()
    merged_terms = load_master_warranty(io.BytesIO(merged_bytes))
    # Abort before saving if an incomplete column mapping would produce the wrong term.
    for article, term in incoming_terms.items():
        if merged_terms.get(article) != term:
            raise ValueError(
                f"เงื่อนไข Article {article} หลังรวมข้อมูลไม่ตรงกับไฟล์ที่อัปโหลด กรุณาตรวจสอบโครงสร้าง Master"
            )
    added = sorted(set(incoming_terms) - set(original_terms))
    changed = sorted(article for article in incoming_terms if article in original_terms
                     and incoming_terms[article] != original_terms[article])
    unchanged = sorted(article for article in incoming_terms if article in original_terms
                       and incoming_terms[article] == original_terms[article])
    return merged_bytes, changed, added, unchanged, merged_terms


def save_master_persistently(merged_bytes, expected_hash):
    """Create a timestamped backup and atomically replace the local master file."""
    MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    with MASTER_LOCK:
        current = MASTER_PATH.read_bytes() if MASTER_PATH.exists() else None
        actual_hash = hashlib.sha256(current).hexdigest() if current is not None else None
        if actual_hash != expected_hash:
            raise RuntimeError("Master ถูกแก้ไขหลังจาก Preview กรุณารีเฟรชและตรวจสอบอีกครั้ง")
        # Backups are never overwritten, even with multiple saves in one second.
        if current is not None:
            backup_dir = MASTER_PATH.parent / "master_backups"
            backup_dir.mkdir(exist_ok=True)
            suffix = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            shutil.copy2(MASTER_PATH, backup_dir / f"{MASTER_PATH.stem}_{suffix}.xlsx")
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(dir=MASTER_PATH.parent, suffix=".xlsx", delete=False) as fh:
                temp_path = Path(fh.name)
                fh.write(merged_bytes)
                fh.flush()
                os.fsync(fh.fileno())
            os.replace(temp_path, MASTER_PATH)
        finally:
            if temp_path is not None and temp_path.exists():
                temp_path.unlink()


def load_billing_file(content):
    sales = pd.read_excel(io.BytesIO(content), dtype=object, keep_default_na=False)
    sales = sales.dropna(how="all")
    sales = sales.loc[~sales.apply(lambda column: column.map(clean_value).eq("")).all(axis=1)].copy()
    headers = {name: find_column(sales.columns, (name,)) for name in REQUIRED_COLUMNS}
    missing = [name for name, header in headers.items() if header is None]
    serial_cols = [
        column
        for name in SERIAL_COLUMNS
        if (column := find_column(sales.columns, (name,))) is not None
    ]
    if not serial_cols:
        missing.append("ManufSerialNumber หรือ Serial Number")
    if missing:
        raise ValueError("ขาดคอลัมน์ " + ", ".join(missing))

    billing = sales[[headers[name] for name in REQUIRED_COLUMNS]].copy()
    billing.columns = REQUIRED_COLUMNS
    # Prefer ManufSerialNumber whenever that column exists; otherwise use Serial Number.
    billing["Serial Number"] = sales[serial_cols[0]].map(clean_value)
    for name in ("Billing", "Customer name", "Article"):
        billing[name] = billing[name].map(clean_value)
    if billing.empty:
        raise ValueError("ไฟล์ไม่มีรายการสินค้า")
    if billing["Billing"].eq("").any():
        raise ValueError("มีรายการที่เว้น Billing ว่าง กรุณาตรวจไฟล์ประจำบิล")
    return billing


def prepare_certificate(billing_no, billing, warranty_map, extensions, issue_date):
    customers = billing["Customer name"].loc[billing["Customer name"].ne("")].unique()
    if len(customers) != 1:
        raise ValueError("บิลที่เลือกต้องระบุชื่อลูกค้า Customer name เพียงชื่อเดียว")
    if billing["Article"].eq("").any():
        raise ValueError("บิลที่เลือกมีรายการที่เว้น Article ว่าง")

    items = []
    conditions = {}
    for article, serial, billing_date in billing[
        ["Article", "Serial Number", "Billing Date"]
    ].itertuples(index=False, name=None):
        item = {"article": article, "sn": serial, "start_date": format_date(billing_date)}
        items.append(item)
        if article not in conditions:
            extension = ""
            if article in extensions:
                years, reference = extensions[article]
                extension = f"{years} Years Warranty Extension Refer to Billing No. {reference}"
            conditions[article] = {
                "article": article,
                "term_text": warranty_map.get(article, ""),
                "extension": extension,
            }
    return {
        "doc_date": format_document_date(issue_date),
        "billing_no": billing_no,
        "customer_name": customers[0],
        "items": items,
        "product_conditions": list(conditions.values()),
    }


def generate_docx(context):
    template = DocxTemplate(TEMPLATE_PATH)
    template.render(context, autoescape=True)
    output = io.BytesIO()
    template.save(output)
    return output.getvalue()


def generate_pdf(word_bytes):
    with tempfile.TemporaryDirectory(prefix="warranty-") as work_dir:
        source = Path(work_dir) / "certificate.docx"
        output = source.with_suffix(".pdf")
        source.write_bytes(word_bytes)
        if sys.platform == "win32":
            import pythoncom
            from win32com.client import DispatchEx

            pythoncom.CoInitialize()
            word = document = None
            try:
                word = DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                document = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
                document.ExportAsFixedFormat(str(output), 17)
            finally:
                try:
                    if document is not None:
                        document.Close(False)
                finally:
                    try:
                        if word is not None:
                            word.Quit(False)
                    finally:
                        pythoncom.CoUninitialize()
        elif sys.platform == "darwin":
            subprocess.run(
                [
                    sys.executable, "-c",
                    "from docx2pdf import convert; import sys; convert(*sys.argv[1:], keep_active=True)",
                    str(source), str(output),
                ],
                check=True,
                capture_output=True,
                timeout=90,
            )
        else:
            raise RuntimeError("การส่งออก PDF ต้องใช้ Microsoft Word บน Windows หรือ macOS")
        pdf = output.read_bytes()
        if not pdf.startswith(b"%PDF-"):
            raise RuntimeError("Microsoft Word ไม่ได้ส่งออกไฟล์ PDF ที่อ่านได้")
        return pdf


def main():
    st.set_page_config(page_title="Warranty Certificate Generator", layout="wide")
    st.title("Warranty Certificate Generator")
    st.caption("อัปโหลดไฟล์ประจำบิล ตรวจสอบเงื่อนไข และดาวน์โหลดหนังสือรับประกัน")

    st.subheader("1. อัปโหลดไฟล์ Excel ประจำบิล")
    uploaded = st.file_uploader(
        "ไฟล์ Excel (.xlsx)",
        type=["xlsx"],
        help="ใช้คอลัมน์ Billing Date, Customer name, Billing, Article และ ManufSerialNumber หรือ Serial Number (ให้ความสำคัญกับ ManufSerialNumber ก่อนเสมอ)",
    )
    st.subheader("2. อัปโหลดเงื่อนไขรับประกันสินค้า (Master Warranty)")
    uploaded_master = st.file_uploader(
        "อัปโหลดไฟล์เพื่อเพิ่มหรืออัปเดต Master Warranty (.xlsx)",
        type=["xlsx"],
        key="master-upload",
        help="ใช้ชีต Master_Product_Warranty และคอลัมน์ Article Code พร้อม Product Warranty Term หรือ Standard Base Warranty",
    )
    # Read the persisted file on every Streamlit rerun. Uploaded data is not used
    # for certificates until the user confirms and saves it to disk.
    master_bytes = MASTER_PATH.read_bytes() if MASTER_PATH.is_file() else None
    master_hash = hashlib.sha256(master_bytes).hexdigest() if master_bytes is not None else None
    warranty_map = {}
    if master_bytes is not None:
        try:
            warranty_map = load_master_warranty(io.BytesIO(master_bytes))
        except Exception as exc:
            st.error(f"อ่าน Master Warranty หลักไม่สำเร็จ ({exc})")
            return
    else:
        st.info("ยังไม่มี Master หลัก กรุณาอัปโหลดไฟล์ Master Warranty เพื่อตั้งต้น")

    if uploaded_master is not None:
        incoming_content = uploaded_master.getvalue()
        try:
            merged_bytes, changed, added, unchanged, incoming_terms = preview_master_merge(
                master_bytes, incoming_content
            )
        except Exception as exc:
            st.error(f"ตรวจสอบ Master ที่อัปโหลดไม่สำเร็จ ({exc})")
        else:
            st.write(f"ตัวอย่างผลการนำเข้า: เพิ่มใหม่ {len(added)} | เปลี่ยนเงื่อนไข {len(changed)} | เงื่อนไขเดิม {len(unchanged)}")
            comparison = []
            for article in added + changed + unchanged:
                comparison.append({
                    "Article": article,
                    "สถานะ": "เพิ่มใหม่" if article in added else ("เงื่อนไขเปลี่ยน" if article in changed else "เงื่อนไขเดิม"),
                    "เงื่อนไขเดิม": warranty_map.get(article, ""),
                    "เงื่อนไขใหม่": incoming_terms[article],
                })
            st.dataframe(comparison, hide_index=True, use_container_width=True)
            st.warning("การบันทึกจะอัปเดตไฟล์ Master Warranty หลักในโฟลเดอร์โปรแกรม และมีการสำรองไฟล์เดิมก่อนบันทึก")
            confirm = st.checkbox("ยืนยันการเพิ่ม/อัปเดต Master Warranty หลัก", key="confirm-master-save")
            if st.button("Save to Master (บันทึกถาวร)", disabled=not confirm, type="primary"):
                try:
                    save_master_persistently(merged_bytes, master_hash)
                    st.session_state["master-save-message"] = (
                        f"บันทึก Master สำเร็จ: เพิ่ม {len(added)} Article, "
                        f"เปลี่ยนเงื่อนไข {len(changed)} Article"
                    )
                    st.rerun()
                except Exception as exc:
                    st.error(f"ไม่สามารถบันทึก Master Warranty ได้ ({exc})")
    save_message = st.session_state.pop("master-save-message", None)
    if save_message:
        st.success(save_message)
    if master_bytes is not None:
        st.caption(f"Master ปัจจุบัน: {MASTER_PATH.name} ({len(warranty_map)} Article) — เก็บไว้บนเครื่องที่รัน Streamlit")
        st.download_button(
            "Download Master Warranty (สำรองข้อมูล)", master_bytes,
            file_name=MASTER_PATH.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if not warranty_map:
        st.warning("ยังไม่มี Master Warranty ที่บันทึกถาวร กรุณาอัปโหลดและกด Save ก่อนออกหนังสือรับประกัน")
        return
    master_id = master_hash

    if uploaded is None:
        return

    content = uploaded.getvalue()
    upload_id = hashlib.sha256(content).hexdigest()
    try:
        sales = load_billing_file(content)
    except Exception as exc:
        st.error(f"อ่านไฟล์ประจำบิลไม่สำเร็จ ({exc})")
        return

    billing_no = st.selectbox("Refer Billing No.", sales["Billing"].unique(), key=f"bill-{upload_id}")
    billing = sales.loc[sales["Billing"].eq(billing_no)]
    articles = billing["Article"].loc[billing["Article"].ne("")].unique().tolist()
    scope = f"{upload_id}-{billing_no}"

    st.subheader("3. Warranty Extension")
    has_extension = st.radio(
        "บิลนี้มีสินค้าที่ซื้อ Warranty Extension เพิ่มเติมหรือไม่?",
        ("No", "Yes"),
        horizontal=True,
        key=f"extension-{scope}",
    )
    extensions = {}
    extension_ready = True
    if has_extension == "Yes":
        selected = st.multiselect("เลือก Article ที่ซื้อประกันเพิ่ม", articles, key=f"articles-{scope}")
        for article in selected:
            st.write(f"Article {article}")
            years_col, reference_col = st.columns([1, 2])
            years = years_col.number_input(
                "จำนวนปีที่ต่อเพิ่ม", min_value=1, value=5, step=1, key=f"years-{scope}-{article}",
            )
            reference = reference_col.text_input(
                "Billing No. อ้างอิงการซื้อประกันเพิ่ม", key=f"reference-{scope}-{article}",
            ).strip()
            extensions[article] = (years, reference)
        extension_ready = bool(selected) and all(reference for _, reference in extensions.values())
        if not extension_ready:
            st.warning("เลือก Article และกรอก Billing No. อ้างอิงให้ครบทุก Article ที่ต่อประกัน")

    st.subheader("4. ตรวจสอบและสร้างเอกสาร")
    try:
        context = prepare_certificate(billing_no, billing, warranty_map, extensions, date.today())
    except ValueError as exc:
        st.error(str(exc))
        return

    st.write("Dear", context["customer_name"])
    by_article = {condition["article"]: condition for condition in context["product_conditions"]}
    rows = []
    for number, item in enumerate(context["items"], 1):
        condition = by_article[item["article"]]
        term = condition["term_text"] or "กรุณาเพิ่มเงื่อนไขใน Master"
        if condition["extension"]:
            term += f" ({condition['extension']})"
        rows.append({
            "No.": number,
            "Article": item["article"],
            "Serial Number": item["sn"],
            "Warranty Start Date": item["start_date"],
            "Warranty Condition": term,
        })
    st.dataframe(rows, hide_index=True)
    missing_terms = [article for article in articles if article not in warranty_map]
    if missing_terms:
        st.warning("กรุณาเพิ่มเงื่อนไขใน Master สำหรับ Article " + ", ".join(missing_terms))
    if missing_terms or not extension_ready:
        return

    filename = "Warranty_" + re.sub(r"[^A-Za-z0-9_-]+", "_", billing_no)[:100]
    result_key = (upload_id, master_id, context)
    if st.button("Generate Warranty Certificate", type="primary"):
        st.session_state.pop("certificate", None)
        try:
            word_bytes = generate_docx(context)
        except Exception as exc:
            st.error(f"สร้างไฟล์ Word ไม่สำเร็จ ({exc})")
            return
        st.session_state.certificate = {"key": result_key, "word": word_bytes}

    result = st.session_state.get("certificate")
    if not result or result["key"] != result_key:
        return
    st.download_button(
        "Download Word (.docx)", result["word"], file_name=f"{filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    if st.button("Generate PDF"):
        try:
            with st.spinner("กำลังแปลงเอกสารด้วย Microsoft Word"):
                result["pdf"] = generate_pdf(result["word"])
        except Exception as exc:
            st.warning(f"สร้าง PDF ไม่สำเร็จ กรุณาตรวจสอบ Microsoft Word บน Windows หรือ macOS ({exc})")
    if "pdf" in result:
        st.download_button("Download PDF (.pdf)", result["pdf"], file_name=f"{filename}.pdf", mime="application/pdf")


if __name__ == "__main__":
    main()
