"""Run with: python -m streamlit run app.py"""

import hashlib
import io
import re
import subprocess
import sys
import tempfile
from datetime import date, datetime
from numbers import Real
from pathlib import Path

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


def load_master_warranty():
    raw = pd.read_excel(
        MASTER_PATH,
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
    try:
        warranty_map = load_master_warranty()
    except Exception as exc:
        st.error(f"อ่าน Master Warranty ไม่สำเร็จ ({exc})")
        return
    if not warranty_map:
        st.info("กรอก Article และ Warranty Condition ในชีต Master_Product_Warranty ก่อนออกเอกสาร")
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

    st.subheader("2. Warranty Extension")
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

    st.subheader("3. ตรวจสอบและสร้างเอกสาร")
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
    result_key = (upload_id, context)
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
