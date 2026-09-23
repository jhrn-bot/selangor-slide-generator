import datetime
import io
import os
import streamlit as st
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

# Automatically retrieve the correct ME and Year
epi_week, year = get_previous_epi_week()

st.title("📊 Selangor Epi Review Slide Generator")
st.write(f"This application automatically generates the title slide for **ME {epi_week:02d} / {year}**.")
st.divider()

# Function to generate the presentation
def generate_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Top Navy Border
    top_border = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.4))
    top_border.fill.solid()
    top_border.fill.fore_color.rgb = RGBColor(27, 54, 93)
    top_border.line.fill.background()
    
    # Bottom Navy Border
    bottom_border = slide.shapes.add_shape(1, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4))
    bottom_border.fill.solid()
    bottom_border.fill.fore_color.rgb = RGBColor(27, 54, 93)
    bottom_border.line.fill.background()
    
    # Left Navy Corner Accents
    left_border = slide.shapes.add_shape(1, Inches(0), Inches(0.4), Inches(0.4), Inches(1.5))
    left_border.fill.solid()
    left_border.fill.fore_color.rgb = RGBColor(27, 54, 93)
    left_border.line.fill.background()
    
    left_border_bottom = slide.shapes.add_shape(1, Inches(0), Inches(5.6), Inches(0.4), Inches(1.5))
    left_border_bottom.fill.solid()
    left_border_bottom.fill.fore_color.rgb = RGBColor(27, 54, 93)
    left_border_bottom.line.fill.background()

    # Right Navy Corner Accents
    right_border = slide.shapes.add_shape(1, Inches(12.933), Inches(0.4), Inches(0.4), Inches(1.5))
    right_border.fill.solid()
    right_border.fill.fore_color.rgb = RGBColor(27, 54, 93)
    right_border.line.fill.background()
    
    right_border_bottom = slide.shapes.add_shape(1, Inches(12.933), Inches(5.6), Inches(0.4), Inches(1.5))
    right_border_bottom.fill.solid()
    right_border_bottom.fill.fore_color.rgb = RGBColor(27, 54, 93)
    right_border_bottom.line.fill.background()

    # Background Inner Card (Grey thin border)
    shape = slide.shapes.add_shape(1, Inches(0.5), Inches(0.5), Inches(12.333), Inches(6.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(220, 220, 220)
    shape.line.width = Pt(1.5)

    # Insert Logo (If logo.png exists in the folder)
    if os.path.exists("logo.png"):
        slide.shapes.add_picture("logo.png", Inches(5.66), Inches(1.0), width=Inches(2.0))

    # Main Title Text
    txBox = slide.shapes.add_textbox(Inches(1.5), Inches(3.2), Inches(10.333), Inches(1.2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Selangor Epidemiology Review"
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.name = 'Calibri'
    p.font.color.rgb = RGBColor(27, 54, 93)
    p.alignment = PP_ALIGN.CENTER

    # Small Grey Divider Line
    line = slide.shapes.add_shape(9, Inches(6.0), Inches(4.3), Inches(1.3), Inches(0))
    line.line.color.rgb = RGBColor(190, 190, 190)
    line.line.width = Pt(2.5)

    # Subtitle Text (Automated ME Date)
    txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(4.6), Inches(10.333), Inches(0.8))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = f"ME {epi_week:02d} / {year}"
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.name = 'Calibri'
    p2.font.color.rgb = RGBColor(27, 54, 93)
    p2.alignment = PP_ALIGN.CENTER

    # Insert Fixed QR Code (If qr.png exists in the folder)
    if os.path.exists("qr.png"):
        slide.shapes.add_picture("qr.png", Inches(10.0), Inches(4.2), width=Inches(2.2))

    # Save to memory buffer
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# Simplified UI: Just one button
if st.button("🚀 Generate Slide", type="primary", use_container_width=True):
    pptx_buffer = generate_pptx()
    st.success(f"Slide successfully generated for ME {epi_week:02d} / {year}!")
    
    st.download_button(
        label="📥 Download Presentation (.pptx)",
        data=pptx_buffer,
        file_name=f"Selangor_Epi_Review_ME{epi_week:02d}_{year}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
