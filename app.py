import datetime
import io
import os
import urllib.request
import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

st.set_page_config(
    page_title="Selangor Epi Review Slide Generator",
    page_icon="📊",
    layout="centered"
)

GOOGLE_SLIDES_ID = "1QFVgrEPqgiDditxLQhHRLapQOqaCjZnt"
VALID_DISTRICTS = [
    'GOMBAK', 'HULU LANGAT', 'HULU SELANGOR', 'KLANG', 
    'KUALA LANGAT', 'KUALA SELANGOR', 'PETALING', 
    'SABAK BERNAM', 'SEPANG'
]

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

# --- UI Setup ---
st.title("📊 Selangor Epi Review Slide Generator")
st.write(f"Target Output: **ME {epi_week:02d} / {year}** (Malaysia Time UTC+8)")

st.divider()
st.subheader("1. Upload Data")
uploaded_file = st.file_uploader("Upload raw Excel data for Analisa e-Notifikasi", type=["xlsx", "xls"])

def fetch_google_slides_pptx(file_id):
    url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        return io.BytesIO(response.read())

# --- Data Processing Logic ---
def process_data(df, target_me):
    col_me = df.columns[7]
    col_dt = df.columns[123] 
    col_br = df.columns[69]

    df_clean = df[df[col_dt].isin(VALID_DISTRICTS)].copy()
    df_clean[col_me] = pd.to_numeric(df_clean[col_me], errors='coerce')
    
    df_semasa = df_clean[df_clean[col_me] == target_me]
    df_kumulatif = df_clean[df_clean[col_me] <= target_me]
    
    def get_stats(df_subset):
        total = len(df_subset)
        counts = df_subset[col_br].value_counts().to_dict()
        
        daftar_notifikasi = counts.get('Daftar Notifikasi', 0)
        daftar_kes = counts.get('Daftar Kes', 0)
        abai = counts.get('Abai Notifikasi', 0)
        belum = counts.get('Belum Ambil Tindakan', 0)
        batal = counts.get('Batal Daftar', 0)
        
        return {
            'total': total,
            'daftar_notifikasi': daftar_notifikasi,
            'daftar_kes': daftar_kes,
            'abai': abai,
            'belum': belum,
            'batal': batal,
            'pct_daftar_notif': f"{(daftar_notifikasi/total*100):.2f}%" if total else "0.00%",
            'pct_daftar_kes': f"{(daftar_kes/total*100):.2f}%" if total else "0.00%",
            'pct_abai': f"{(abai/total*100):.2f}%" if total else "0.00%",
            'pct_belum': f"{(belum/total*100):.2f}%" if total else "0.00%",
            'pct_batal': f"{(batal/total*100):.2f}%" if total else "0.00%",
        }
        
    return get_stats(df_semasa), get_stats(df_kumulatif)

# --- Presentation Generation ---
def generate_pptx(stats_semasa, stats_kumulatif):
    try:
        remote_buffer = fetch_google_slides_pptx(GOOGLE_SLIDES_ID)
        prs = Presentation(remote_buffer)
    except Exception as e:
        st.warning(f"Could not fetch Google Slides directly: {e}. Generating standalone slides.")
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    # 1. Format existing tables
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        cell.vertical_anchor = MSO_ANCHOR.MIDDLE 
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.alignment = PP_ALIGN.CENTER
                            for run in paragraph.runs:
                                run.font.size = Pt(11)

    # 2. Build Slide 1 (Title Slide)
    if len(prs.slides) > 0:
        slide1 = prs.slides[0]
        for shape in list(slide1.shapes):
            sp = shape._element
            sp.getparent().remove(sp)
    else:
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])

    NAVY = RGBColor(16, 44, 87)

    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(2.2), Inches(0.25)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()
    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(0.25), Inches(1.8)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()
    
    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.933), Inches(0.2), Inches(2.2), Inches(0.25)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()
    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.883), Inches(0.2), Inches(0.25), Inches(1.8)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()
    
    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(7.05), Inches(2.2), Inches(0.25)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()
    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(5.5), Inches(0.25), Inches(1.8)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()

    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.933), Inches(7.05), Inches(2.2), Inches(0.25)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()
    slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.883), Inches(5.5), Inches(0.25), Inches(1.8)).fill.solid()
    slide1.shapes[-1].fill.fore_color.rgb = NAVY; slide1.shapes[-1].line.fill.background()

    card = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.42), Inches(0.42), Inches(12.493), Inches(6.66))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255)
    card.line.color.rgb = RGBColor(210, 210, 210); card.line.width = Pt(1.5)

    if os.path.exists("logo.png"):
        slide1.shapes.add_picture("logo.png", Inches(5.66), Inches(0.85), width=Inches(2.0))

    txBox = slide1.shapes.add_textbox(Inches(1.5), Inches(2.85), Inches(10.333), Inches(1.2))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Selangor Epidemiology Review"; p.font.size = Pt(50); p.font.bold = True
    p.font.name = 'Calibri'; p.font.color.rgb = NAVY; p.alignment = PP_ALIGN.CENTER

    line = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.916), Inches(4.05), Inches(1.5), Inches(0.03))
    line.fill.solid(); line.fill.fore_color.rgb = RGBColor(200, 200, 200); line.line.fill.background()

    txBox2 = slide1.shapes.add_textbox(Inches(1.5), Inches(4.25), Inches(10.333), Inches(0.8))
    p2 = txBox2.text_frame.paragraphs[0]
    p2.text = f"ME {epi_week:02d} / {year}"; p2.font.size = Pt(28); p2.font.bold = True
    p2.font.name = 'Calibri'; p2.font.color.rgb = NAVY; p2.alignment = PP_ALIGN.CENTER

    if os.path.exists("qr.png"):
        slide1.shapes.add_picture("qr.png", Inches(10.35), Inches(4.55), width=Inches(2.2))

    # --- 3. Build Slide 2 (Analisa e-Notifikasi) ---
    slide2_layout = prs.slide_layouts[6]
    slide2 = prs.slides.add_slide(slide2_layout)
    
    # REVISED SLIDE REORDERING FIX
    slide_id = prs.slides._sldIdLst[-1]
    prs.slides._sldIdLst.remove(slide_id)
    prs.slides._sldIdLst.insert(1, slide_id)

    if os.path.exists("logo.png"):
        slide2.shapes.add_picture("logo.png", Inches(0.5), Inches(0.4), width=Inches(1.5))

    txBox = slide2.shapes.add_textbox(Inches(2.2), Inches(0.5), Inches(8), Inches(0.8))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Analisa e-Notifikasi"
    p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"ME {epi_week:02d} /{year}"
    p2.font.size = Pt(18); p2.font.color.rgb = NAVY

    rows, cols = 6, 5
    table_shape = slide2.shapes.add_table(rows, cols, Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.5))
    table = table_shape.table

    table.columns[0].width = Inches(2.333)
    table.columns[1].width = Inches(2.25)
    table.columns[2].width = Inches(2.25)
    table.columns[3].width = Inches(2.25)
    table.columns[4].width = Inches(2.25)

    headers = [
        "Minggu Epid", f"ME {epi_week:02d} (Semasa)", "", f"Kumulatif Sehingga ME {epi_week:02d}", ""
    ]
    for i, txt in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = txt
        cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(230, 240, 250)

    table.cell(0, 1).merge(table.cell(0, 2))
    table.cell(0, 3).merge(table.cell(0, 4))

    row_labels = ["Jumlah Notifikasi", "Daftar Notifikasi", "Daftar Kes", "Abai Notifikasi", "Belum Ambil Tindakan", "Batal Daftar"]
    
    for i in range(1, 6):
        table.cell(i, 0).text = row_labels[i]
        table.cell(i, 0).fill.solid(); table.cell(i, 0).fill.fore_color.rgb = RGBColor(230, 240, 250)

    table.cell(1, 1).text = f"{stats_semasa['total']:,}"
    table.cell(1, 2).fill.solid(); table.cell(1, 2).fill.fore_color.rgb = RGBColor(0, 0, 0)
    table.cell(1, 3).text = f"{stats_kumulatif['total']:,}"
    table.cell(1, 4).fill.solid(); table.cell(1, 4).fill.fore_color.rgb = RGBColor(0, 0, 0)

    table.cell(2, 1).text = f"{stats_semasa['daftar_notifikasi']:,}"
    table.cell(2, 2).text = stats_semasa['pct_daftar_notif']
    table.cell(2, 3).text = f"{stats_kumulatif['daftar_notifikasi']:,}"
    table.cell(2, 4).text = stats_kumulatif['pct_daftar_notif']

    table.cell(3, 1).text = f"{stats_semasa['daftar_kes']:,}"
    table.cell(3, 2).text = stats_semasa['pct_daftar_kes']
    table.cell(3, 3).text = f"{stats_kumulatif['daftar_kes']:,}"
    table.cell(3, 4).text = stats_kumulatif['pct_daftar_kes']

    table.cell(4, 1).text = f"{stats_semasa['abai']:,}"
    table.cell(4, 2).text = stats_semasa['pct_abai']
    table.cell(4, 3).text = f"{stats_kumulatif['abai']:,}"
    table.cell(4, 4).text = stats_kumulatif['pct_abai']

    table.cell(5, 0).text = "Belum Ambil Tindakan\nBatal Daftar"
    table.cell(5, 1).text = f"{stats_semasa['belum']:,}\n{stats_semasa['batal']:,}"
    table.cell(5, 2).text = f"{stats_semasa['pct_belum']}\n{stats_semasa['pct_batal']}"
    table.cell(5, 3).text = f"{stats_kumulatif['belum']:,}\n{stats_kumulatif['batal']:,}"
    table.cell(5, 4).text = f"{stats_kumulatif['pct_belum']}\n{stats_kumulatif['pct_batal']}"

    for row in table.rows:
        for cell in row.cells:
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE 
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(16)
                    run.font.name = 'Calibri'
                    
    myt_zone = datetime.timezone(datetime.timedelta(hours=8))
    now = datetime.datetime.now(myt_zone)
    timestamp_str = now.strftime("%d/%m/%Y @ %I.%M%p").upper()
    
    txBox = slide2.shapes.add_textbox(Inches(0.2), Inches(6.9), Inches(10), Inches(0.4))
    p = txBox.text_frame.paragraphs[0]
    p.text = f"(Sumber : Sistem e-notifikasi, KKM muat turun pada ({timestamp_str}))"
    p.font.size = Pt(10); p.font.italic = True
    
    bottom_banner = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

st.divider()

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    st.success(f"Data loaded successfully: {len(df_raw)} records found.")
    
    if st.button("🚀 Generate Slide Deck", type="primary", use_container_width=True):
        with st.spinner("Processing data and generating slides..."):
            stats_semasa, stats_kumulatif = process_data(df_raw, epi_week)
            pptx_buffer = generate_pptx(stats_semasa, stats_kumulatif)
            st.success(f"Slide deck ready! Includes custom Title Slide and Analisa e-Notifikasi for ME {epi_week:02d}.")
            st.download_button(
                label="📥 Download Presentation (.pptx)",
                data=pptx_buffer,
                file_name=f"Selangor_Epi_Review_ME{epi_week:02d}_{year}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
else:
    st.info("⚠️ Please upload the Excel file to enable the Generate button.")
