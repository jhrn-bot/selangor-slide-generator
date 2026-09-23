import datetime
import io
import os
import re
import urllib.request
import qrcode
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="Selangor Epidemiology Slide Generator",
    page_icon="📊",
    layout="centered",
)

DEFAULT_GSLIDE_URL = "https://docs.google.com/presentation/d/1QFVgrEPqgiDditxLQhHRLapQOqaCjZnt/edit?usp=sharing"

# ---------------------------------------------------------
# Helper Functions: Epiweek Calculation (Malaysia Time UTC+8)
# ---------------------------------------------------------
def get_kkm_epi_week(dt: datetime.date = None):
    """Calculates the KKM Epidemiological Week starting on Sunday."""
    if dt is None:
        myt_zone = datetime.timezone(datetime.timedelta(hours=8))
        dt = datetime.datetime.now(myt_zone).date()

    day_of_week_sun = (dt.weekday() + 1) % 7
    week_start = dt - datetime.timedelta(days=day_of_week_sun)

    year = dt.year
    jan1 = datetime.date(year, 1, 1)
    jan1_day_sun = (jan1.weekday() + 1) % 7

    if jan1_day_sun <= 3:
        week1_start = jan1 - datetime.timedelta(days=jan1_day_sun)
    else:
        week1_start = jan1 + datetime.timedelta(days=(7 - jan1_day_sun))

    if week_start < week1_start:
        year -= 1
        jan1_prev = datetime.date(year, 1, 1)
        jan1_prev_day_sun = (jan1_prev.weekday() + 1) % 7
        if jan1_prev_day_sun <= 3:
            week1_start = jan1_prev - datetime.timedelta(days=jan1_prev_day_sun)
        else:
            week1_start = jan1_prev + datetime.timedelta(days=(7 - jan1_prev_day_sun))

    week_number = ((week_start - week1_start).days // 7) + 1
    return week_number, year


def get_previous_epi_week():
    """Returns the previous epidemiological week in Malaysia Time."""
    myt_zone = datetime.timezone(datetime.timedelta(hours=8))
    today = datetime.datetime.now(myt_zone).date()
    prev_date = today - datetime.timedelta(days=7)
    return get_kkm_epi_week(prev_date)


# ---------------------------------------------------------
# Helper Functions: Google Slides Downloader
# ---------------------------------------------------------
def extract_gslide_id(url: str) -> str:
    """Extracts the file ID from a Google Slides or Google Drive URL."""
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    match = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    return url.strip()


def fetch_google_slides_pptx(url: str):
    """Downloads a public Google Slides deck as PPTX bytes."""
    file_id = extract_gslide_id(url)
    download_urls = [
        f"https://docs.google.com/presentation/d/{file_id}/export/pptx",
        f"https://drive.google.com/uc?export=download&id={file_id}",
    ]
    headers = {"User-Agent": "Mozilla/5.0"}
    for dl_url in download_urls:
        try:
            req = urllib.request.Request(dl_url, headers=headers)
            with urllib.request.urlopen(req, timeout=20) as response:
                if response.status == 200:
                    data = response.read()
                    if data.startswith(b"PK\x03\x04"):  # Valid PPTX zip header
                        return io.BytesIO(data)
        except Exception:
            continue
    return None


# ---------------------------------------------------------
# Slide Drawing: Title Slide
# ---------------------------------------------------------
def render_title_slide_elements(slide, me_number: int, year_val: int, custom_qr_file=None, logo_file=None):
    """Renders the official Selangor Epidemiology title layout onto a slide."""
    navy_blue = RGBColor(16, 44, 87)
    border_color = RGBColor(220, 224, 230)
    white = RGBColor(255, 255, 255)
    divider_grey = RGBColor(200, 200, 200)
    kkm_text_grey = RGBColor(40, 40, 40)

    # 1. Dark Navy Corner Accent Blocks
    # Top-Left Bracket
    tl_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(2.2), Inches(0.25))
    tl_h.fill.solid(); tl_h.fill.fore_color.rgb = navy_blue; tl_h.line.fill.background()
    tl_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(0.25), Inches(1.8))
    tl_v.fill.solid(); tl_v.fill.fore_color.rgb = navy_blue; tl_v.line.fill.background()

    # Top-Right Bracket
    tr_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.933), Inches(0.2), Inches(2.2), Inches(0.25))
    tr_h.fill.solid(); tr_h.fill.fore_color.rgb = navy_blue; tr_h.line.fill.background()
    tr_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.883), Inches(0.2), Inches(0.25), Inches(1.8))
    tr_v.fill.solid(); tr_v.fill.fore_color.rgb = navy_blue; tr_v.line.fill.background()

    # Bottom-Left Bracket
    bl_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(7.05), Inches(2.2), Inches(0.25))
    bl_h.fill.solid(); bl_h.fill.fore_color.rgb = navy_blue; bl_h.line.fill.background()
    bl_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(5.5), Inches(0.25), Inches(1.8))
    bl_v.fill.solid(); bl_v.fill.fore_color.rgb = navy_blue; bl_v.line.fill.background()

    # Bottom-Right Bracket
    br_h = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.933), Inches(7.05), Inches(2.2), Inches(0.25))
    br_h.fill.solid(); br_h.fill.fore_color.rgb = navy_blue; br_h.line.fill.background()
    br_v = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.883), Inches(5.5), Inches(0.25), Inches(1.8))
    br_v.fill.solid(); br_v.fill.fore_color.rgb = navy_blue; br_v.line.fill.background()

    # 2. Main Framed White Card
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.42), Inches(0.42), Inches(12.493), Inches(6.66))
    card.fill.solid(); card.fill.fore_color.rgb = white
    card.line.color.rgb = border_color
    card.line.width = Pt(1.5)

    # 3. Logo (Jata Negara)
    logo_width = Inches(1.7)
    logo_left = Inches((13.333 - 1.7) / 2)
    logo_top = Inches(0.85)

    if logo_file:
        slide.shapes.add_picture(logo_file, logo_left, logo_top, width=logo_width)
    elif os.path.exists("logo.png"):
        slide.shapes.add_picture("logo.png", logo_left, logo_top, width=logo_width)
    else:
        try:
            url = "https://upload.wikimedia.org/wikipedia/commons/thumb/2/26/Coat_of_arms_of_Malaysia.svg/440px-Coat_of_arms_of_Malaysia.svg.png"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=5) as response:
                logo_bytes = io.BytesIO(response.read())
            slide.shapes.add_picture(logo_bytes, logo_left, logo_top, width=logo_width)
        except Exception:
            pass

    # 4. KKM / JKNS Header Subtitle
    kkm_box = slide.shapes.add_textbox(Inches(2.5), Inches(2.15), Inches(8.333), Inches(0.65))
    tf_kkm = kkm_box.text_frame
    tf_kkm.word_wrap = True

    p_kkm1 = tf_kkm.paragraphs[0]
    p_kkm1.text = "Kementerian Kesihatan Malaysia"
    p_kkm1.font.name = "Arial"
    p_kkm1.font.size = Pt(13)
    p_kkm1.font.bold = True
    p_kkm1.font.color.rgb = kkm_text_grey
    p_kkm1.alignment = PP_ALIGN.CENTER

    p_kkm2 = tf_kkm.add_paragraph()
    p_kkm2.text = "Jabatan Kesihatan Negeri Selangor"
    p_kkm2.font.name = "Arial"
    p_kkm2.font.size = Pt(13)
    p_kkm2.font.bold = True
    p_kkm2.font.color.rgb = kkm_text_grey
    p_kkm2.alignment = PP_ALIGN.CENTER

    # 5. Main Title: "Selangor Epidemiology Review"
    title_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.95), Inches(11.333), Inches(1.1))
    tf_title = title_box.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.text = "Selangor Epidemiology Review"
    p_title.font.name = "Arial"
    p_title.font.size = Pt(46)
    p_title.font.bold = True
    p_title.font.color.rgb = navy_blue
    p_title.alignment = PP_ALIGN.CENTER

    # 6. Horizontal Divider Bar
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.916), Inches(4.15), Inches(1.5), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = divider_grey
    line.line.fill.background()

    # 7. Dynamic ME / Year Subtitle
    me_box = slide.shapes.add_textbox(Inches(2.0), Inches(4.35), Inches(9.333), Inches(0.8))
    tf_me = me_box.text_frame
    p_me = tf_me.paragraphs[0]
    p_me.text = f"ME {me_number} / {year_val}"
    p_me.font.name = "Arial"
    p_me.font.size = Pt(28)
    p_me.font.bold = True
    p_me.font.color.rgb = navy_blue
    p_me.alignment = PP_ALIGN.CENTER

    # 8. Fixed QR Code (Exact Bottom-Right Corner)
    qr_size = Inches(2.2)
    qr_left = Inches(10.35)
    qr_top = Inches(4.55)

    if custom_qr_file:
        slide.shapes.add_picture(custom_qr_file, qr_left, qr_top, width=qr_size, height=qr_size)
    elif os.path.exists("qr.png"):
        slide.shapes.add_picture("qr.png", qr_left, qr_top, width=qr_size, height=qr_size)
    else:
        qr = qrcode.QRCode(version=1, box_size=10, border=1)
        qr.add_data("https://jknselangor.moh.gov.my/")
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        qr_buffer = io.BytesIO()
        qr_img.save(qr_buffer, format="PNG")
        qr_buffer.seek(0)
        slide.shapes.add_picture(qr_buffer, qr_left, qr_top, width=qr_size, height=qr_size)


# ---------------------------------------------------------
# Slide Creation (Full Deck with Google Slides Import)
# ---------------------------------------------------------
def create_full_deck(
    me_number: int,
    year_val: int,
    gslide_url: str = None,
    custom_qr_file=None,
    logo_file=None,
) -> io.BytesIO:
    """Generates slide 1 and appends slides 2+ from Google Slides."""
    prs = None
    gslide_loaded = False

    # 1. Attempt to fetch presentation from Google Slides
    if gslide_url:
        gslide_stream = fetch_google_slides_pptx(gslide_url)
        if gslide_stream:
            try:
                prs = Presentation(gslide_stream)
                gslide_loaded = True
            except Exception:
                prs = None

    # Fallback to empty presentation if download fails or link omitted
    if prs is None:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        title_slide = prs.slides.add_slide(prs.slide_layouts[6])
    else:
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        title_slide = prs.slides[0]
        # Exclude original Slide 1 content: remove all existing shapes from slide 0
        for shape in list(title_slide.shapes):
            sp = shape._element
            sp.getparent().remove(sp)

    # 2. Render new title slide on Slide 1
    render_title_slide_elements(
        title_slide,
        me_number=me_number,
        year_val=year_val,
        custom_qr_file=custom_qr_file,
        logo_file=logo_file,
    )

    # 3. Save merged deck
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream, len(prs.slides), gslide_loaded


# ---------------------------------------------------------
# Streamlit Interface
# ---------------------------------------------------------
st.title("📊 Selangor Epidemiology Presentation")
st.markdown(
    "Automates your weekly slide deck. Calculates the previous **Epidemiological Week (ME)** in Malaysia Time (UTC+8) and imports the remaining slides from your Google Slides document."
)

calc_prev_me, calc_year = get_previous_epi_week()

with st.expander("📅 Week & Template Settings", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        selected_me = st.number_input(
            "Epidemiological Week (ME)",
            min_value=1,
            max_value=53,
            value=calc_prev_me,
            help="Defaults to the previous completed epidemiological week.",
        )
    with col2:
        selected_year = st.number_input(
            "Year",
            min_value=2020,
            max_value=2040,
            value=calc_year,
        )

    gslide_link = st.text_input(
        "Google Slides Document URL (Slides 2+)",
        value=DEFAULT_GSLIDE_URL,
        help="The public Google Slides link containing your review slides.",
    )

st.markdown("---")
st.info(f"**Target Title Slide:** `Selangor Epidemiology Review` — **ME {selected_me} / {selected_year}**")

if st.button("🚀 Generate Full Slide Deck (PPTX)", type="primary", use_container_width=True):
    with st.spinner("Fetching Google Slides and building presentation..."):
        pptx_buffer, total_slides, imported = create_full_deck(
            me_number=int(selected_me),
            year_val=int(selected_year),
            gslide_url=gslide_link.strip(),
        )

        if imported:
            st.success(
                f"✅ Deck generated! Title slide updated to **ME {selected_me} / {selected_year}** and **{total_slides - 1} slides** imported from Google Slides."
            )
        else:
            st.warning(
                f"⚠️ Could not reach Google Slides link. Generated standalone Title Slide for **ME {selected_me} / {selected_year}**."
            )

        st.download_button(
            label="📥 Download Presentation (.pptx)",
            data=pptx_buffer,
            file_name=f"Selangor_Epidemiology_Review_ME{selected_me}_{selected_year}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
```

### What this accomplishes:
1. **Excludes Google Slide #1**: Replaces the old title slide on Slide 1 with your automated, pixel-accurate card layout.
2. **Imports Slides 2+ Intact**: All subsequent slides (such as the *Pengerusi ER* schedule table and charts) are retained with their original tables, layout formatting, and styles.
3. **Automatic Malaysia Time ME**: Pre-fills the previous epidemiological week (e.g. ME 37 / 2026).
4. **QR Code in Corner**: Placed flush in the bottom-right corner. If `qr.png` exists in the repo, it is used automatically.
