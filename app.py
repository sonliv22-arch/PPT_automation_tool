import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import os
from datetime import datetime
import io
import tempfile

st.set_page_config(page_title="PPT Automation Tool", layout="wide")
st.title("📊 PPT Data Entry Form")

st.markdown("---")
st.markdown("<h4 style='color:blue'>Plant & Equipment Details</h4>", unsafe_allow_html=True)

# --- FORM ---
with st.form("data_form"):

    # --- Plant & Equipment ---
    col1, col2 = st.columns(2)
    with col1:
        plant_name = st.text_input("Plant Name")
        equipment_name = st.text_input("Equipment")
        case_enabler = st.text_input("Case Enabler")
    with col2:
        downtime_hours = st.text_input("Downtime Hours")
        observation_date = st.date_input("Observation Date")
        observation = st.text_input("Observation")

    st.markdown("---")
    st.markdown("<h4 style='color:blue'>Recommendations & Actions</h4>", unsafe_allow_html=True)
    
    # --- Recommendations ---
    col3, col4 = st.columns(2)
    with col3:
        date_recommendation = st.date_input("Date of Recommendation")
        recommendation = st.text_input("Recommendation")
        date_corrective_action = st.date_input("Date of Corrective Action Taken")
    with col4:
        corrective_action_details = st.text_input("Corrective Action Details")
        date_closed_report = st.date_input("Date of Closed Report")
        closed_report_status = st.text_input("Closed Report Status")

    st.markdown("---")
    st.markdown("<h4 style='color:blue'>Machine & Trend Details</h4>", unsafe_allow_html=True)
    
    with st.expander("Machine Details & Trend Text"):
        machine_details = st.text_area("Machine Details")
        trend_for_image1 = st.text_input("Trend for Image-1")
        trend_for_image2 = st.text_input("Trend for Image-2")
    
    st.markdown("---")
    st.markdown("<h4 style='color:blue'>Upload Images</h4>", unsafe_allow_html=True)

    # --- Images ---
    col_img1, col_img2, col_img3 = st.columns(3)
    with col_img1:
        equipment_image = st.file_uploader("Equipment Image", type=['png','jpg','jpeg'])
    with col_img2:
        trend_image1 = st.file_uploader("Trend Image 1", type=['png','jpg','jpeg'])
    with col_img3:
        trend_image2 = st.file_uploader("Trend Image 2", type=['png','jpg','jpeg'])
    
    submit = st.form_submit_button("Generate PPT", use_container_width=True)

# --- ON SUBMIT ---
if submit:
    template_path = "template.pptx"
    prs = Presentation(template_path)
    
    date_format = "%d-%m-%Y"
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
    
    # Replace text and insert images
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
    
    # --- SAVE PPT TO BYTES ---
    ppt_bytes = io.BytesIO()
    prs.save(ppt_bytes)
    ppt_bytes.seek(0)
    
    st.success("✅ PPT generated successfully!")
    st.download_button(
        "Download PPT", 
        ppt_bytes, 
        file_name="output.pptx", 
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
