import streamlit as st
from pypdf import PdfWriter, PdfReader
from streamlit_sortables import sort_items
import tempfile
import os

st.set_page_config(page_title="📎 PDF Merger with Reorder", page_icon="📎")
st.title("📎 Combine PDF Files with Drag-and-Drop Reorder")

uploaded_files = st.file_uploader("Upload file PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    filenames = [f.name for f in uploaded_files]
    st.write(f"📄 {len(filenames)} file terunggah:")
    for f in filenames:
        st.write(f"• {f}")

    st.markdown("### ✋ Drag & Drop untuk mengatur urutan:")

    # ✅ Hapus argumen label
    reordered_filenames = sort_items(filenames, direction="vertical")

    # Petakan nama file ke file asli
    filename_to_file = {f.name: f for f in uploaded_files}
    reordered_files = [filename_to_file[name] for name in reordered_filenames]

    st.markdown("### 📑 Urutan Final:")
    for i, name in enumerate(reordered_filenames, 1):
        st.write(f"{i}. {name}")

    if st.button("Gabungkan PDF"):
        with st.spinner("🔄 Menggabungkan file..."):
            writer = PdfWriter()
            temp_files = []

            try:
                for uploaded in reordered_files:
                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    temp_file.write(uploaded.read())
                    temp_file.close()
                    temp_files.append(temp_file.name)

                    reader = PdfReader(temp_file.name)
                    for page in reader.pages:
                        writer.add_page(page)

                output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                with open(output_path, "wb") as f_out:
                    writer.write(f_out)

                with open(output_path, "rb") as f:
                    st.success("✅ Penggabungan selesai!")
                    st.download_button(
                        label="📥 Download PDF Gabungan",
                        data=f.read(),
                        file_name="Hasil_Gabungan.pdf",
                        mime="application/pdf"
                    )

            finally:
                for file in temp_files:
                    os.unlink(file)
                if os.path.exists(output_path):
                    os.unlink(output_path)