import datetime
import io
import os
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

st.set_page_config(
    page_title="Selangor Epi Review Slide Generator",
    page_icon="📊",
    layout="centered"
)

# --- Calculate Epid Week (Malaysia Time UTC+8) ---
def get_previous_epi_week():
    myt_zone = datetime.timezone(datetime.timedelta(hours=8))
    today = datetime.datetime.now(myt_zone).date()
    
    last_week_date = today - datetime.timedelta(days=7)
    
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

epi_week, year = get_previous_epi_week()

st.title("📊 Selangor Epi Review Slide Generator")
st.write(f"Target Output: **ME {epi_week:02d} / {year}** (Malaysia Time UTC+8)")
st.divider()

def generate_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    NAVY = RGBColor(16, 44, 87)
    
    # --- 1. Navy Corner Brackets (Drawn first so central card overlays inner edges) ---
    # Top-Left Bracket
    tl_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(2.2), Inches(0.25))
    tl_h.fill.solid(); tl_h.fill.fore_color.rgb = NAVY; tl_h.line.fill.background()
    tl_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(0.25), Inches(1.8))
    tl_v.fill.solid(); tl_v.fill.fore_color.rgb = NAVY; tl_v.line.fill.background()

    # Top-Right Bracket
    tr_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.933), Inches(0.2), Inches(2.2), Inches(0.25))
    tr_h.fill.solid(); tr_h.fill.fore_color.rgb = NAVY; tr_h.line.fill.background()
    tr_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.883), Inches(0.2), Inches(0.25), Inches(1.8))
    tr_v.fill.solid(); tr_v.fill.fore_color.rgb = NAVY; tr_v.line.fill.background()

    # Bottom-Left Bracket
    bl_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(7.05), Inches(2.2), Inches(0.25))
    bl_h.fill.solid(); bl_h.fill.fore_color.rgb = NAVY; bl_h.line.fill.background()
    bl_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(5.5), Inches(0.25), Inches(1.8))
    bl_v.fill.solid(); bl_v.fill.fore_color.rgb = NAVY; bl_v.line.fill.background()

    # Bottom-Right Bracket
    br_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.933), Inches(7.05), Inches(2.2), Inches(0.25))
    br_h.fill.solid(); br_h.fill.fore_color.rgb = NAVY; br_h.line.fill.background()
    br_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.883), Inches(5.5), Inches(0.25), Inches(1.8))
    br_v.fill.solid(); br_v.fill.fore_color.rgb = NAVY; br_v.line.fill.background()

    # --- 2. Central White Card ---
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.42), Inches(0.42), Inches(12.493), Inches(6.66))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(255, 255, 255)
    card.line.color.rgb = RGBColor(210, 210, 210)
    card.line.width = Pt(1.5)

    # --- 3. Content Inside Card ---
    # Logo
    if os.path.exists("logo.png"):
        slide.shapes.add_picture("logo.png", Inches(5.66), Inches(0.85), width=Inches(2.0))

    # Main Title
    txBox = slide.shapes.add_textbox(Inches(1.5), Inches(2.95), Inches(10.333), Inches(1.2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Selangor Epidemiology Review"
    p.font.size = Pt(50)
    p.font.bold = True
    p.font.name = 'Calibri'
    p.font.color.rgb = NAVY
    p.alignment = PP_ALIGN.CENTER

    # Short Line Divider
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.916), Inches(4.15), Inches(1.5), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(200, 200, 200)
    line.line.fill.background()

    # Subtitle Text (ME Week)
    txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(4.35), Inches(10.333), Inches(0.8))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = f"ME {epi_week:02d} / {year}"
    p2.font.size = Pt(28)
    p2.font.bold = True
    p2.font.name = 'Calibri'
    p2.font.color.rgb = NAVY
    p2.alignment = PP_ALIGN.CENTER

    # Fixed QR Code
    if os.path.exists("qr.png"):
        slide.shapes.add_picture("qr.png", Inches(9.6), Inches(3.75), width=Inches(2.4))

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("🚀 Generate Slide", type="primary", use_container_width=True):
    pptx_buffer = generate_pptx()
    st.success(f"Slide generated for ME {epi_week:02d} / {year}!")
    
    st.download_button(
        label="📥 Download Presentation (.pptx)",
        data=pptx_buffer,
        file_name=f"Selangor_Epi_Review_ME{epi_week:02d}_{year}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
