import streamlit as st
from detector import TemplateDetector
import os

st.set_page_config(
    page_title="Outdated Template Detector",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Outdated Template Detector")

st.write(
    "Upload a Word (.docx) or PDF document to check "
    "whether it is using the latest approved template."
)

uploaded_file = st.file_uploader(
    "Choose a file",
    type=["docx", "pdf"]
)

if uploaded_file is not None:

    # Save uploaded file temporarily
    with open(uploaded_file.name, "wb") as f:
        f.write(uploaded_file.getbuffer())

    detector = TemplateDetector("master.xlsx")

    result = detector.detect(uploaded_file.name)

    st.subheader("Detection Result")

    if "error" in result:
        st.error(result["error"])

    else:

        st.write(f"**Template ID:** {result['template_id']}")
        st.write(f"**Document Version:** {result['document_version']}")
        st.write(f"**Latest Version:** {result['latest_version']}")

        if result["status"] == "Latest":
            st.success(result["message"])

        elif result["status"] == "Outdated":
            st.warning(result["message"])

        else:
            st.error(result["message"])

    # Delete uploaded file
    os.remove(uploaded_file.name)