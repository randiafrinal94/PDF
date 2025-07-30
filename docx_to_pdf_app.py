import streamlit as st
from docx2pdf import convert as convert_docx
import tempfile
import os
import shutil
import comtypes.client

st.set_page_config(page_title="📄 Word & Excel ke PDF", page_icon="📝")
st.title("📄 Konversi DOCX & XLSX ke PDF")
st.markdown("Upload file Word (`.docx`) dan/atau Excel (`.xlsx`) lalu konversi otomatis ke PDF.")

uploaded_files = st.file_uploader(
    "Upload file DOCX dan XLSX",
    type=["docx", "xlsx", "xls"],
    accept_multiple_files=True
)

def convert_excel_to_pdf(input_path, output_path):
    excel = comtypes.client.CreateObject("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(input_path)
    wb.ExportAsFixedFormat(0, output_path)  # 0 = PDF
    wb.Close(False)
    excel.Quit()

# Inisialisasi state untuk menyimpan hasil konversi
if "converted_pdfs" not in st.session_state:
    st.session_state.converted_pdfs = []

if uploaded_files and st.button("Konversi Semua ke PDF"):
    with st.spinner("🔄 Mengonversi..."):
        temp_input_dir = tempfile.mkdtemp()
        temp_output_dir = tempfile.mkdtemp()
        result_list = []

        try:
            for file in uploaded_files:
                input_path = os.path.join(temp_input_dir, file.name)
                with open(input_path, "wb") as f:
                    f.write(file.read())

                filename_wo_ext = os.path.splitext(file.name)[0]
                output_path = os.path.join(temp_output_dir, f"{filename_wo_ext}.pdf")

                if file.name.lower().endswith(".docx"):
                    convert_docx(input_path, output_path)
                elif file.name.lower().endswith((".xlsx", ".xls")):
                    convert_excel_to_pdf(input_path, output_path)

                # Simpan hasil ke memory (bukan file langsung)
                with open(output_path, "rb") as f:
                    result_list.append({
                        "name": f"{filename_wo_ext}.pdf",
                        "content": f.read()
                    })

            # Simpan ke session_state agar tidak hilang saat rerun
            st.session_state.converted_pdfs = result_list
            st.success("✅ Konversi selesai!")

        except Exception as e:
            st.error(f"❌ Gagal mengonversi: {e}")
        finally:
            shutil.rmtree(temp_input_dir)
            shutil.rmtree(temp_output_dir)

# Tampilkan tombol download untuk semua hasil konversi
if st.session_state.get("converted_pdfs"):
    st.markdown("### 📥 Unduh File PDF:")
    for pdf in st.session_state.converted_pdfs:
        st.download_button(
            label=f"📄 Download {pdf['name']}",
            data=pdf["content"],
            file_name=pdf["name"],
            mime="application/pdf"
        )