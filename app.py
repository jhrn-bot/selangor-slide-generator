import datetime
import io
import pandas as pd
import qrcode
import requests
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

st.set_page_config(
    page_title="Selangor Epi Review Slide Generator",
    page_icon="📊",
    layout="centered"
)

# --- Calculate Epid Week (CDC Standard: Sunday start) ---
def get_previous_epi_week():
    today = datetime.date.today()
    # Subtract 7 days to get last week's date
    last_week_date = today - datetime.timedelta(days=7)
    
    # Calculate Epi Week (CDC style: Sunday starts the week)
    jan_1 = datetime.date(last_week_date.year, 1, 1)
    jan_1_day = jan_1.isoweekday() % 7  # Sunday = 0
    
    first_sunday = jan_1 if jan_1_day == 0 else jan_1 + datetime.timedelta(days=(7 - jan_1_day))
    
    if last_week_date < first_sunday:
        epi_week = 1
        year = last_week_date.year
    else:
        days_diff = (last_week_date - first_sunday).days
        epi_week = (days_diff // 7) + 1
        year = last_week_date.year
        
    return epi_week, year

default_week, default_year = get_previous_epi_week()

st.title("📊 Selangor Epi Review Slide Generator")
st.caption("Generate title slides automatically with the correct Epidemiological Week.")

st.divider()

# Input options
col1, col2 = st.columns(2)
with col1:
    epi_week = st.number_input("Epid Week (ME)", min_value=1, max_value=53, value=default_week)
with col2:
    year = st.number_input("Year", min_value=2020, max_value=2030, value=default_year)

title_text = st.text_input("Title", value="Selangor Epidemiology Review")
qr_link = st.text_input("QR Code Link / URL", value="https://moh.gov.my")

st.info(f"Generated text will display: **ME {epi_week:02d} / {year}**")

# Function to generate presentation
def generate_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Background card
    shape = slide.shapes.add_shape(
        1, Inches(0.5), Inches(0.5), Inches(12.333), Inches(6.5)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(180, 180, 180)
    shape.line.width = Pt(1.5)

    # Title Text
    txBox = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.333), Inches(1.2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(27, 54, 93)
    p.alignment = PP_ALIGN.CENTER

    # Subtitle ME Text
    txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(4.2), Inches(10.333), Inches(0.8))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = f"ME {epi_week} / {year}"
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(27, 54, 93)
    p2.alignment = PP_ALIGN.CENTER

    # Generate QR Code
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(qr_link)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")
    
    qr_buffer = io.BytesIO()
    qr_img.save(qr_buffer, format="PNG")
    qr_buffer.seek(0)
    
    slide.shapes.add_picture(qr_buffer, Inches(9.8), Inches(3.8), width=Inches(2.2))

    # Save to buffer
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("🚀 Generate Slide", type="primary", use_container_width=True):
    pptx_buffer = generate_pptx()
    st.success("Slide generated!")
    st.download_button(
        label="📥 Download Presentation (.pptx)",
        data=pptx_buffer,
        file_name=f"Selangor_Epi_Review_ME{epi_week}_{year}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
