import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import os
from datetime import datetime
import io
import tempfile

# --- Page config ---
st.set_page_config(page_title="📝 PPT Data Entry", layout="centered")

# --- Custom CSS for cream background and black text ---
st.markdown(
    """
    <style>
    .stApp {
        background-color: #fffdd0;  /* Cream background */
        color: black;
    }
    .css-1aumxhk, .stTextInput>div>div>input, .stTextArea>div>textarea {
        color: black !important;
        background-color: #ffffff !important;
    }
    .stButton>button {
        background-color: #0073e6;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("<h1 style='text-align: center; color: darkblue;'>📊 PPT Data Entry Form</h1>", unsafe_allow_html=True)
st.write("---")

date_format = "%d-%m-%Y"

# --- Step tracker ---
step = st.session_state.get("step", 1)

# --- FORM ---
with st.form("data_form"):
    if step == 1:
        st.subheader("1️⃣ Basic Information")
        plant_name = st.text_input("Plant Name", placeholder="Enter Plant Name")
        equipment_name = st.text_input("Equipment", placeholder="Enter Equipment Name")
        case_enabler = st.text_input("Case Enabler", placeholder="Enter Case Enabler")
        downtime_hours = st.text_input("Downtime Hours", placeholder="Enter downtime hours")
        equipment_image = st.file_uploader("📷 Upload Equipment Image", type=['png','jpg','jpeg'])
        if st.form_submit_button("Next Step"):
            st.session_state.step = 2
            st.experimental_rerun()

    elif step == 2:
        st.subheader("2️⃣ Observation & Recommendation")
        observation_date = st.date_input("Observation Date")
        observation = st.text_input("Observation", placeholder="Enter observation details")
        date_recommendation = st.date_input("Date of Recommendation")
        recommendation = st.text_input("Recommendation", placeholder="Enter recommendation")
        if st.form_submit_button("Next Step"):
            st.session_state.step = 3
            st.experimental_rerun()

    elif step == 3:
        st.subheader("3️⃣ Corrective Actions")
        date_corrective_action = st.date_input("Date of Corrective Action Taken")
        corrective_action_details = st.text_input("Corrective Action Details")
        date_closed_report = st.date_input("Date of Closed Report")
        closed_report_status = st.text_input("Closed Report Status")
        if st.form_submit_button("Next Step"):
            st.session_state.step = 4
            st.experimental_rerun()

    elif step == 4:
        st.subheader("4️⃣ Machine Details & Trends")
        machine_details = st.text_area("Machine Details", placeholder="Enter machine details")
        trend_for_image1 = st.text_input("Trend for Image-1", placeholder="Enter trend description")
        trend_image1 = st.file_uploader("📷 Upload Trend Image 1", type=['png','jpg','jpeg'])
        trend_for_image2 = st.text_input("Trend for Image-2", placeholder="Enter trend description")
        trend_image2 = st.file_uploader("📷 Upload Trend Image 2", type=['png','jpg','jpeg'])
        if st.form_submit_button("✅ Generate PPT"):
            # --- Generate PPT ---
            template_path = "template.pptx"
            prs = Presentation(template_path)

            replacements = {
                "{{Plant Name}}": plant_name,
                "{{Equipment}}": equipment_name,
                "{{Case Enabler}}": case_enabler,
                "{{Downtime Hours}}": downtime_hours,
                "{{Observation Date}}": observation_date.strftime(date_format),
                "{{Observation}}": observation,
                "{{Date of Recommendation}}": date_recommendation.strftime(date_format),
                "{{Recommendation}}": recommendation,
                "{{Date of Corrective action Taken}}": date_corrective_action.strftime(date_format),
                "{{Corrective Action Details}}": corrective_action_details,
                "{{Date of closed Report}}": date_closed_report.strftime(date_format),
                "{{Closed Report Status}}": closed_report_status,
                "{{Machine Details}}": machine_details,
                "{{Trend for Image-1}}": trend_for_image1,
                "{{Trend for Image-2}}": trend_for_image2
            }

            def replace_placeholder(para, replacements):
                full_text = "".join(run.text for run in para.runs)
                for key, value in replacements.items():
                    full_text = full_text.replace(key, value)
                for run in para.runs:
                    run.text = ""
                if para.runs:
                    para.runs[0].text = full_text
                else:
                    para.add_run().text = full_text

            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        if "{{Equipment Image}}" in shape.text and equipment_image:
                            shape.text = ""
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                                tmp.write(equipment_image.getvalue())
                                tmp_path = tmp.name
                            slide.shapes.add_picture(tmp_path, left, top, width, height)
                            os.unlink(tmp_path)
                        elif "{{Trend Image 1}}" in shape.text and trend_image1:
                            shape.text = ""
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                                tmp.write(trend_image1.getvalue())
                                tmp_path = tmp.name
                            slide.shapes.add_picture(tmp_path, left, top, width, height)
                            os.unlink(tmp_path)
                        elif "{{Trend Image 2}}" in shape.text and trend_image2:
                            shape.text = ""
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                                tmp.write(trend_image2.getvalue())
                                tmp_path = tmp.name
                            slide.shapes.add_picture(tmp_path, left, top, width, height)
                            os.unlink(tmp_path)
                        else:
                            for para in shape.text_frame.paragraphs:
                                replace_placeholder(para, replacements)

            ppt_bytes = io.BytesIO()
            prs.save(ppt_bytes)
            ppt_bytes.seek(0)

            st.success("🎉 PPT generated successfully!")
            st.download_button(
                "⬇️ Download PPT",
                ppt_bytes,
                file_name=f"Equipment_{plant_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
