import streamlit as st
import pandas as pd
from pptx import Presentation
from dispatcher import run_strips_template
import io
import os
from path_helpers import get_base_path

st.set_page_config(layout="wide", page_title="Financial Buyers Presentation Tool")

# === Header with Lincoln logo
st.image(os.path.join(get_base_path(), "logos", "lincolninternational.png"), width=200)
st.markdown("<h2 style='font-family:Arial; color:#003366;'>Buyers Presentation Tool</h2>", unsafe_allow_html=True)

st.markdown("<hr style='border:1px solid #eee'>", unsafe_allow_html=True)

# === Instructions
st.markdown("<h4 style='font-family:Arial; color:#003366;'>Instructions</h4>", unsafe_allow_html=True)
st.markdown("""
<ul style='color:#003366; font-family:Arial;'>
    <li> Upload your Excel file containing the buyers data.</li>
    <li> Adjust the sheet and template settings if needed.</li>
    <li> Click <b>Run Presentation</b> to generate your PowerPoint deck.</li>
</ul>
""", unsafe_allow_html=True)

st.markdown("<hr style='border:1px solid #eee'>", unsafe_allow_html=True)

# === Settings
st.markdown("<h4 style='font-family:Arial; color:#003366;'>📂 Excel & PowerPoint Settings</h4>", unsafe_allow_html=True)

# Upload file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
st.caption("Please ensure your Excel file maintains the 'Python Strip Mask' structure to guarantee accurate slide creation.")
sheet_name = st.text_input("Sheet name", value="Python Financials Mask")
template_file = st.selectbox(
    "Select PPT template file",
    options=["financials_templates.pptx", "financials_templates_wide.pptx"],
    index=0  # defaults to first
)
output_file = st.text_input("Output PPT file name", value="buyers_presentation.pptx")

# Step 1 - Layout options
template_options = {
    "Financial Buyer Strips (Dry Powder only)": 1,
    "Strips with key Financials (Dry Powder + AUM)" : 2,
    "Classic Buyer Strips (Português)" : 3,
    "Strips with key Financials (Português)" : 4
}
selected_template = st.selectbox("Choose Layout Style", list(template_options.keys()))
template_number = template_options[selected_template]


st.markdown("<hr style='border:1px solid #eee'>", unsafe_allow_html=True)

if uploaded_file is not None:
    # Just display filename for user confidence
    st.write(f"✓ Uploaded: {uploaded_file.name}")
    
    
    # When button is pressed
    if st.button("Generate Presentation"):
        try:
            df = pd.read_excel(
                uploaded_file,
                sheet_name=sheet_name,
                header=1,
                usecols="B:B, E:N, P:V, X:AD, AF:AK"
            ).dropna(subset=['fund_name']).drop(columns=['fund_name'])
            st.success(f"✓ Loaded {len(df)} buyers from uploaded file.")

            prs = Presentation(os.path.join(get_base_path(), template_file))
            run_strips_template(template_number, prs=prs, df=df)
            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)
            st.success("✓ Presentation generated successfully.")
            st.download_button(
                label="Download Presentation",
                data=pptx_io,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"✘ Something went wrong: {e}")
else:
    st.info("⬆Please upload an Excel file to begin.")
