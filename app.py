import streamlit as st
from io import BytesIO
import tempfile
import zipfile
import hashlib
from docx import Document

from src.DocumentGeneratorStrategicActivites import StrategicDocGenerator

st.set_page_config(layout="wide")
st.title("Excel → Word Generator")

TEMPLATE_STRATEGIC_ACTIVITIES_PATH = "template_strategic_activities.docx"
TEMPLATE_BUDGET_PATH = "template_budget.docx"


# -------------------------------------------------
# Utility: Create hash of uploaded file
# -------------------------------------------------
def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()


# -------------------------------------------------
# Cached generator (returns serializable bytes)
# -------------------------------------------------
@st.cache_data(show_spinner=True)
def cached_generate_documents(file_bytes, template_path):
    # Write uploaded bytes to a temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(file_bytes)
        temp_excel_path = tmp.name

    # Use the StrategicDocGenerator class
    generator = StrategicDocGenerator(temp_excel_path, template_path)
    docs_dict = generator.generate_documents()

    # Serialize Word docs to bytes
    serialized_docs = {}
    for division, doc_obj in docs_dict.items():
        buffer = BytesIO()
        doc_obj.save(buffer)
        serialized_docs[division] = buffer.getvalue()

    return serialized_docs


# -------------------------------------------------
# Upload Excel
# -------------------------------------------------
uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx"])

# Session state initialization
if "docs_bytes" not in st.session_state:
    st.session_state.docs_bytes = None

if "file_hash" not in st.session_state:
    st.session_state.file_hash = None


if uploaded_excel:

    file_bytes = uploaded_excel.read()
    current_hash = get_file_hash(file_bytes)

    if st.button("Generate Documents"):

        # Regenerate only if file changed
        if st.session_state.file_hash != current_hash:

            docs_bytes = cached_generate_documents(
                file_bytes,
                TEMPLATE_STRATEGIC_ACTIVITIES_PATH
            )

            st.session_state.docs_bytes = docs_bytes
            st.session_state.file_hash = current_hash

        st.success("Documents Generated Successfully!")

    # -------------------------------------------------
    # Display Generated Documents
    # -------------------------------------------------
    if st.session_state.docs_bytes:

        docs_bytes = st.session_state.docs_bytes

        # ---------------- ZIP DOWNLOAD ----------------
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for division, doc_bytes in docs_bytes.items():
                zf.writestr(f"{division}.docx", doc_bytes)

        zip_buffer.seek(0)

        st.download_button(
            label="⬇ Download All as ZIP",
            data=zip_buffer,
            file_name="generated_documents.zip",
            mime="application/zip"
        )

        st.markdown("---")
        st.subheader("Generated Documents")

        # ---------------- INDIVIDUAL DOCUMENTS ----------------
        for division, doc_bytes in docs_bytes.items():

            with st.expander(f"📄 {division}", expanded=False):

                # Rebuild document for preview only
                doc_obj = Document(BytesIO(doc_bytes))

                # -------- HEADER (Preview + Download side-by-side) --------
                col1, col2 = st.columns([6, 2])

                with col1:
                    st.markdown("### Document Preview")

                with col2:
                    st.download_button(
                        label="⬇ Download",
                        data=doc_bytes,
                        file_name=f"{division}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"download_{division}"
                    )

                st.markdown("---")

                # -------- PARAGRAPHS --------
                for para in doc_obj.paragraphs:
                    if para.text.strip():
                        st.write(para.text)

                # -------- TABLES --------
                for table_index, table in enumerate(doc_obj.tables):

                    st.markdown(f"#### Table {table_index + 1}")

                    table_data = []
                    for row in table.rows:
                        row_data = [cell.text.strip() for cell in row.cells]
                        table_data.append(row_data)

                    if table_data:
                        st.table(table_data)