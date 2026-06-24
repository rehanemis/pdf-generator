import streamlit as st
import openpyxl
import zipfile
import io
import os
import uuid
from datetime import datetime, date
from pypdf import PdfReader, PdfWriter

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PDF Generator",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Customer PDF Generator")
st.markdown("Upload your Excel file and PDF template to generate one PDF per customer.")

# ── Helper functions ────────────────────────────────────────────────────────────

def format_date(value, fmt="%d/%m/%Y"):
    """Convert Excel date or string to formatted string."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime(fmt)
    try:
        # Try parsing common string formats
        for f in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(str(value), f).strftime(fmt)
            except ValueError:
                continue
    except Exception:
        pass
    return str(value)

def generate_ref(prefix, index):
    """Generate a unique reference number."""
    return f"{prefix}-{str(index).zfill(4)}"

def safe_filename(name):
    """Make a string safe for use as a filename."""
    keepchars = (" ", ".", "_", "-")
    return "".join(c if (c.isalnum() or c in keepchars) else "_" for c in str(name)).strip()

def get_excel_rows(excel_bytes):
    """Read Excel file and return list of dicts (one per row)."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    headers = [str(cell.value).strip() if cell.value else f"col_{i}" 
               for i, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1)))]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(cell is not None for cell in row):
            rows.append(dict(zip(headers, row)))
    return headers, rows

def fill_pdf_form(pdf_bytes, field_map):
    """Fill fillable PDF form fields using field_map dict {field_name: value}."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    writer.append(reader)
    writer.update_page_form_field_values(writer.pages[0], field_map)
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()

def list_pdf_fields(pdf_bytes):
    """Return list of fillable field names in PDF."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    fields = reader.get_fields()
    if fields:
        return list(fields.keys())
    return []

# ── Sidebar settings ────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    ref_prefix = st.text_input("Reference Number Prefix", value="REF-2024")
    date_format = st.selectbox("Date Format in PDF", [
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%B %d, %Y",
        "%d %B %Y"
    ])
    naming_col = st.text_input("Column to use for PDF filename", value="Full Name",
                                help="Which Excel column to use for naming each PDF")
    st.markdown("---")
    st.markdown("**Built for:** PDF Automation Workflow")
    st.markdown("**Version:** 1.0")

# ── Step 1: Upload files ────────────────────────────────────────────────────────
st.header("Step 1: Upload Files")
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("📊 Upload Excel File (.xlsx)", type=["xlsx"])
with col2:
    pdf_template = st.file_uploader("📄 Upload PDF Template (.pdf)", type=["pdf"])

# ── Step 2: Preview & Map columns ──────────────────────────────────────────────
if excel_file and pdf_template:
    excel_bytes = excel_file.read()
    pdf_bytes = pdf_template.read()

    headers, rows = get_excel_rows(excel_bytes)
    pdf_fields = list_pdf_fields(pdf_bytes)

    st.success(f"✅ Excel loaded: **{len(rows)} customers**, **{len(headers)} columns**")

    if pdf_fields:
        st.success(f"✅ PDF has **{len(pdf_fields)} fillable fields** detected")
    else:
        st.warning("⚠️ No fillable fields found in PDF. The PDF may be a flat/image PDF. "
                   "Please use a PDF with fillable form fields (created in Adobe Acrobat, "
                   "LibreOffice, etc.)")

    # Show Excel preview
    st.header("Step 2: Preview Data")
    with st.expander("👀 Preview Excel Data (first 5 rows)", expanded=True):
        preview_rows = rows[:5]
        if preview_rows:
            st.dataframe(preview_rows)

    # Map Excel columns → PDF fields
    if pdf_fields:
        st.header("Step 3: Map Excel Columns → PDF Fields")
        st.markdown("Match each PDF field to the correct Excel column.")

        col_options = ["-- skip --"] + headers
        field_mapping = {}

        with st.expander("🔗 Field Mapping", expanded=True):
            # Auto-suggest common mappings
            auto_map = {
                "full_name": ["Full Name", "Name", "Customer Name", "fullname"],
                "first_name": ["First Name", "firstname", "First"],
                "last_name": ["Last Name", "lastname", "Surname"],
                "dob": ["Date of Birth", "DOB", "Birth Date", "Birthday"],
                "address": ["Address", "Street", "Street Address"],
                "city": ["City", "Town"],
                "state": ["State", "Province"],
                "zip": ["ZIP", "ZIP Code", "Postal Code", "Postcode"],
                "email": ["Email", "Email Address", "E-mail"],
                "phone": ["Phone", "Phone Number", "Mobile", "Tel"],
            }

            def suggest(pdf_field):
                """Suggest a column based on PDF field name."""
                pf_lower = pdf_field.lower().replace(" ", "_").replace("-", "_")
                for key, candidates in auto_map.items():
                    if key in pf_lower:
                        for c in candidates:
                            if c in headers:
                                return c
                # Try direct match
                for h in headers:
                    if h.lower().strip() == pdf_field.lower().strip():
                        return h
                return "-- skip --"

            cols = st.columns(2)
            for i, field in enumerate(pdf_fields):
                with cols[i % 2]:
                    suggested = suggest(field)
                    default_idx = col_options.index(suggested) if suggested in col_options else 0
                    chosen = st.selectbox(
                        f"PDF: `{field}`",
                        options=col_options,
                        index=default_idx,
                        key=f"map_{field}"
                    )
                    if chosen != "-- skip --":
                        field_mapping[field] = chosen

        # Date fields
        st.markdown("**Which columns are dates?** (will be formatted automatically)")
        date_cols = st.multiselect(
            "Select date columns",
            options=headers,
            default=[h for h in headers if any(d in h.lower() for d in ["date", "dob", "birth", "expiry"])]
        )

        # ── Step 4: Generate ──────────────────────────────────────────────────
        st.header("Step 4: Generate PDFs")
        st.markdown(f"Ready to generate **{len(rows)} PDFs** with mapping for **{len(field_mapping)} fields**.")

        if st.button("🚀 Generate All PDFs", type="primary", use_container_width=True):
            if not field_mapping:
                st.error("Please map at least one field before generating.")
            else:
                progress = st.progress(0)
                status = st.empty()
                zip_buffer = io.BytesIO()

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for i, row in enumerate(rows):
                        # Build field values
                        values = {}
                        for pdf_field, excel_col in field_mapping.items():
                            val = row.get(excel_col, "")
                            if excel_col in date_cols:
                                val = format_date(val, date_format)
                            else:
                                val = str(val) if val is not None else ""
                            values[pdf_field] = val

                        # Add reference number
                        ref = generate_ref(ref_prefix, i + 1)
                        # Try to fill ref field if it exists
                        for rf in ["reference", "ref", "ref_number", "reference_number"]:
                            if rf in pdf_fields:
                                values[rf] = ref

                        # Fill PDF
                        try:
                            filled = fill_pdf_form(pdf_bytes, values)
                        except Exception as e:
                            st.warning(f"Row {i+1}: Error filling PDF — {e}")
                            continue

                        # Name the file
                        name_val = row.get(naming_col, f"customer_{i+1}")
                        filename = safe_filename(str(name_val)) + f"_{ref}.pdf"
                        zf.writestr(filename, filled)

                        progress.progress((i + 1) / len(rows))
                        status.text(f"Processing {i+1}/{len(rows)}: {filename}")

                zip_buffer.seek(0)
                progress.progress(1.0)
                status.success(f"✅ Done! {len(rows)} PDFs generated.")

                st.download_button(
                    label="📥 Download All PDFs (ZIP)",
                    data=zip_buffer,
                    file_name=f"generated_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
    else:
        st.info("💡 **Tip:** Your PDF needs fillable form fields. Create them in Adobe Acrobat, "
                "LibreOffice Draw, or any PDF form editor. Then re-upload.")

else:
    st.info("👆 Please upload both an Excel file and a PDF template to get started.")

# ── Footer ──────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("PDF Generator v1.0 — Upload Excel → Map Fields → Download PDFs")
