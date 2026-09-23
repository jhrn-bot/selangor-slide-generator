import datetime
import io
import urllib.request
from PIL import Image
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


# ---------------------------------------------------------
# Helper Functions: Epiweek Calculation
# ---------------------------------------------------------
def get_kkm_epi_week(dt: datetime.date = None):
    """Calculates the CDC/KKM Epidemiological Week (Minggu Epidemiologi).

    Weeks start on Sunday. Week 1 is the first week containing >= 4 days of the year.
    """
    if dt is None:
        dt = datetime.date.today()

    # Convert Python Monday=0...Sunday=6 to Sunday=0...Saturday=6
    # Sunday is the start of the epi week
    day_of_week_sun = (dt.weekday() + 1) % 7

    # Find the Sunday of the current week
    week_start = dt - datetime.timedelta(days=day_of_week_sun)

    # Determine Year for Week 1 (check January 1st)
    year = dt.year
    jan1 = datetime.date(year, 1, 1)
    jan1_day_sun = (jan1.weekday() + 1) % 7

    # Week 1 starts on the Sunday closest to Jan 1 with at least 4 days in January
    if jan1_day_sun <= 3:
        week1_start = jan1 - datetime.timedelta(days=jan1_day_sun)
    else:
        week1_start = jan1 + datetime.timedelta(days=(7 - jan1_day_sun))

    # If the date is before week 1 start, it falls into the previous year's last epiweek
    if week_start < week1_start:
        year -= 1
        jan1_prev = datetime.date(year, 1, 1)
        jan1_prev_day_sun = (jan1_prev.weekday() + 1) % 7
        if jan1_prev_day_sun <= 3:
            week1_start = jan1_prev - datetime.timedelta(days=jan1_prev_day_sun)
        else:
            week1_start = jan1_prev + datetime.timedelta(
                days=(7 - jan1_prev_day_sun)
            )

    week_number = ((week_start - week1_start).days // 7) + 1
    return week_number, year


def get_previous_epi_week(dt: datetime.date = None):
    """Returns the previous epidemiological week and year."""
    if dt is None:
        dt = datetime.date.today()
    # Go back 7 days to get the previous epi week
    prev_date = dt - datetime.timedelta(days=7)
    return get_kkm_epi_week(prev_date)


# ---------------------------------------------------------
# Slide Creation (python-pptx)
# ---------------------------------------------------------
def create_title_slide(
    me_number: int,
    year_val: int,
    custom_qr_file=None,
    logo_file=None,
) -> io.BytesIO:
    """Generates the 16:9 widescreen presentation matching the template."""
    prs = Presentation()
    # 16:9 widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)

    # Color Palette
    navy_blue = RGBColor(16, 44, 87)
    border_color = RGBColor(220, 224, 230)
    white = RGBColor(255, 255, 255)
    divider_grey = RGBColor(200, 200, 200)
    kkm_text_grey = RGBColor(40, 40, 40)

    # 1. Dark Navy Corner Accent Blocks (Background elements)
    # Top-Left Vertical Bar
    v_tl = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(0.25), Inches(0.2), Inches(1.8)
    )
    v_tl.fill.solid()
    v_tl.fill.fore_color.rgb = navy_blue
    v_tl.line.fill.background()

    # Top-Right Horizontal Bar
    h_tr = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(10.95),
        Inches(0.25),
        Inches(2.133),
        Inches(0.2),
    )
    h_tr.fill.solid()
    h_tr.fill.fore_color.rgb = navy_blue
    h_tr.line.fill.background()

    # Bottom-Left Horizontal Bar
    h_bl = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.25), Inches(5.1), Inches(0.2), Inches(1.8)
    )
    h_bl.fill.solid()
    h_bl.fill.fore_color.rgb = navy_blue
    h_bl.line.fill.background()

    # Bottom-Right Horizontal Bar
    h_br = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(10.95),
        Inches(7.05),
        Inches(2.133),
        Inches(0.2),
    )
    h_br.fill.solid()
    h_br.fill.fore_color.rgb = navy_blue
    h_br.line.fill.background()

    # 2. Main Framed White Card
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.45),
        Inches(0.4),
        Inches(12.433),
        Inches(6.7),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = white
    card.line.color.rgb = border_color
    card.line.width = Pt(1.5)

    # 3. Logo (Jata Negara & KKM)
    logo_width = Inches(1.7)
    logo_left = Inches((13.333 - 1.7) / 2)
    logo_top = Inches(0.95)

    if logo_file:
        slide.shapes.add_picture(
            logo_file, logo_left, logo_top, width=logo_width
        )
    else:
        # Default Jata Negara emblem retrieval
        try:
            url = "https://upload.wikimedia.org/wikipedia/commons/thumb/2/26/Coat_of_arms_of_Malaysia.svg/440px-Coat_of_arms_of_Malaysia.svg.png"
            req = urllib.request.Request(
                url, headers={"User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(req, timeout=5) as response:
                logo_bytes = io.BytesIO(response.read())
            slide.shapes.add_picture(
                logo_bytes, logo_left, logo_top, width=logo_width
            )
        except Exception:
            # Fallback if offline
            pass

    # 4. KKM / JKNS Header Subtitle
    kkm_box = slide.shapes.add_textbox(
        Inches(2.5), Inches(2.25), Inches(8.333), Inches(0.7)
    )
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
    title_box = slide.shapes.add_textbox(
        Inches(1.0), Inches(3.1), Inches(11.333), Inches(1.1)
    )
    tf_title = title_box.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.text = "Selangor Epidemiology Review"
    p_title.font.name = "Arial"
    p_title.font.size = Pt(46)
    p_title.font.bold = True
    p_title.font.color.rgb = navy_blue
    p_title.alignment = PP_ALIGN.CENTER

    # 6. Horizontal Divider Bar
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(5.966), Inches(4.35), Inches(1.4), Inches(0.04)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = divider_grey
    line.line.fill.background()

    # 7. Dynamic ME / Year Subtitle
    me_box = slide.shapes.add_textbox(
        Inches(2.0), Inches(4.7), Inches(9.333), Inches(0.8)
    )
    tf_me = me_box.text_frame
    p_me = tf_me.paragraphs[0]
    p_me.text = f"ME {me_number} / {year_val}"
    p_me.font.name = "Arial"
    p_me.font.size = Pt(28)
    p_me.font.bold = True
    p_me.font.color.rgb = navy_blue
    p_me.alignment = PP_ALIGN.CENTER

    # 8. Insert Exact QR Code (Lower Right)
    qr_size = Inches(2.3)
    qr_left = Inches(9.7)
    qr_top = Inches(4.1)

    if custom_qr_file:
        slide.shapes.add_picture(
            custom_qr_file, qr_left, qr_top, width=qr_size, height=qr_size
        )
    else:
        # Fallback to local image qr.png in repo if available
        try:
            slide.shapes.add_picture(
                "qr.png", qr_left, qr_top, width=qr_size, height=qr_size
            )
        except Exception:
            # Fallback generated QR code if qr.png is not found
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_M,
                box_size=10,
                border=1,
            )
            qr.add_data("https://jknselangor.moh.gov.my/")
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")

            qr_buffer = io.BytesIO()
            qr_img.save(qr_buffer, format="PNG")
            qr_buffer.seek(0)
            slide.shapes.add_picture(
                qr_buffer, qr_left, qr_top, width=qr_size, height=qr_size
            )

    # Save to memory buffer
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream


# ---------------------------------------------------------
# Streamlit Interface
# ---------------------------------------------------------
st.title(" Selangor Epidemiology Presentation")
st.markdown(
    "Automate your weekly epidemiology slide deck. The system automatically computes the previous **Epidemiological Week (ME)** according to the KKM calendar."
)

# Auto-compute previous ME and current year
calc_prev_me, calc_year = get_previous_epi_week()

with st.expander("📅 Week & Metadata Settings", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        selected_me = st.number_input(
            "Epidemiological Week (ME)",
            min_value=1,
            max_value=53,
            value=calc_prev_me,
            help="Defaults automatically to the previous completed epidemiological week.",
        )
    with col2:
        selected_year = st.number_input(
            "Year",
            min_value=2020,
            max_value=2040,
            value=calc_year,
        )

    qr_image_file = st.file_uploader(
        "Upload Custom QR Code Image (Optional)",
        type=["png", "jpg", "jpeg"],
        help="Upload your exact QR code image file, or save 'qr.png' directly in your GitHub repository.",
    )

    custom_logo = st.file_uploader(
        "Upload Custom Jata Negara / Logo (Optional)",
        type=["png", "jpg", "jpeg"],
        help="Leave empty to use standard KKM Jata Negara.",
    )

st.markdown("---")

st.info(f"**Slide to generate:** `Selangor Epidemiology Review` — **ME {selected_me} / {selected_year}**")

if st.button("🚀 Generate Slide Deck (PPTX)", type="primary", use_container_width=True):
    with st.spinner("Generating slide presentation..."):
        logo_bytes = io.BytesIO(custom_logo.read()) if custom_logo else None
        qr_bytes = io.BytesIO(qr_image_file.read()) if qr_image_file else None

        pptx_buffer = create_title_slide(
            me_number=int(selected_me),
            year_val=int(selected_year),
            custom_qr_file=qr_bytes,
            logo_file=logo_bytes,
        )

        st.success(f" Slide 1 generated for ME {selected_me} / {selected_year}!")

        st.download_button(
            label="📥 Download Presentation (.pptx)",
            data=pptx_buffer,
            file_name=f"Selangor_Epidemiology_Review_ME{selected_me}_{selected_year}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
