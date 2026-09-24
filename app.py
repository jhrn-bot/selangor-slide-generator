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
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
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

DISTRICT_ABBR = [
    ('GOMBAK', 'GBK'),
    ('HULU LANGAT', 'HL'),
    ('HULU SELANGOR', 'HS'),
    ('KLANG', 'KLG'),
    ('KUALA LANGAT', 'KL'),
    ('KUALA SELANGOR', 'KS'),
    ('PETALING', 'PTG'),
    ('SABAK BERNAM', 'SB'),
    ('SEPANG', 'SPG')
]

INCLUSION_DIAGNOSES = (
    'HFMD', 'FOOD POISONING', 'COVID-19', 
    'MERS-COV', 'DIPHTERIA', 'DIPHTHERIA', 'CHIKUNGUNYA', 
    'MALARIA', 'LEPROSY'
)

EXCLUSION_DIAGNOSES_14 = (
    'HFMD', 'FOOD POISONING', 'COVID-19', 
    'MERS-COV', 'DIPHTERIA', 'DIPHTHERIA', 'CHIKUNGUNYA', 
    'MALARIA', 'LEPROSY', 'DENGUE/DHF', 'DENGUE', 
    'MEASLES', 'HIV/AIDS', 'HIV'
)

# Diagnosis Notification Period Rules
DIAG_TEMPOH_24H = [
    'DENGUE/DHF', 'DENGUE', 'COVID-19', 'HFMD', 'MEASLES', 
    'AVIAN INFLUENZA', 'FOOD POISONING', 'MERS-COV', 'MALARIA', 
    'DIPHTERIA', 'DIPHTHERIA', 'CHIKUNGUNYA', 'RABIES', 
    'MONKEYPOX', 'MPOX', 'POLIOMYELITIS', 'EBOLA', 'CHOLERA', 
    'ZIKA VIRUS INFECTION', 'YELLOW FEVER', 'PLAGUE'
]

DIAG_TEMPOH_7D = [
    'GONORRHOEA', 'HIV/AIDS', 'HIV', 'VIRAL HEPATITIS', 'PERTUSSIS', 
    'LEPTOSPIROSIS', 'SYPHILIS', 'TUBERCULOSIS', 'RELAPSING FEVER', 
    'TYPHOID/PARATYPHOID', 'DYSENTRY', 'LEPROSY', 'VIRAL ENCEPHALITIS', 
    'TETANUS', 'TYPHUS', 'CHANCROID', 'WHOOPING COUGH', 'BRUCELLOSIS', 
    'MELIODOSIS', 'OTHER SPECIFIED VIRAL HEPATITIS'
]

NAVY = RGBColor(27, 54, 93)
LIGHT_GREY = RGBColor(211, 211, 211)  # #D3D3D3

# ---------------------------------------------------------
# Helper Functions: Epiweek Calculation
# ---------------------------------------------------------
def get_previous_epi_week():
    myt_zone = datetime.timezone(datetime.timedelta(hours=8))
    today = datetime.datetime.now(myt_zone).date()
    
    last_week_date = today - datetime.timedelta(days=7)
    jan_1 = datetime.date(last_week_date.year, 1, 1)
    jan_1_day = jan_1.isoweekday() % 7
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

def parse_excel_date(val):
    """Converts Excel serial number, string, or datetime object into a datetime.date."""
    if pd.isna(val) or val == '' or str(val).strip().upper() == 'NAN':
        return None
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.date() if hasattr(val, 'date') else val
    try:
        num = float(val)
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=num)).date()
    except Exception:
        pass
    try:
        return pd.to_datetime(val, errors='coerce').date()
    except Exception:
        return None

# ---------------------------------------------------------
# Table Formatting Helpers
# ---------------------------------------------------------
def set_cell_border(cell, color=NAVY, width=Pt(1.5)):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for line_type in ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']:
        ln = tcPr.find(qn(line_type))
        if ln is None:
            ln = OxmlElement(line_type)
            tcPr.append(ln)
        ln.set('w', str(int(width)))
        solidFill = OxmlElement('a:solidFill')
        srgbClr = OxmlElement('a:srgbClr')
        srgbClr.set('val', f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')
        solidFill.append(srgbClr)
        ln.append(solidFill)

def write_cell(cell, text, bold=True, align_left=False, is_red=False, size=12):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    lines = str(text).split('\n')
    for idx, line_str in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT if align_left else PP_ALIGN.CENTER
        run = p.add_run()
        run.text = line_str
        run.font.size = Pt(size)
        run.font.name = 'Calibri'
        run.font.bold = bold
        run.font.color.rgb = RGBColor(255, 0, 0) if is_red else NAVY
    return tf.paragraphs[0]

def format_cell_stat(tot, swasta, zero_as_dash=True):
    if tot == 0:
        return "-" if zero_as_dash else "0"
    if swasta > 0:
        return f"{tot} ({swasta})"
    return f"{tot}"

def fetch_google_slides_pptx(file_id):
    url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        return io.BytesIO(response.read())

# ---------------------------------------------------------
# Data Extraction & Processing
# ---------------------------------------------------------
@st.cache_data
def load_and_process_data(file, target_me, inclusion_tuple, exclusion_tuple):
    df = pd.read_excel(file)
    col_me = df.columns[5]    # Column F (0-indexed: 5)
    col_bc = df.columns[54]   # Column BC (Tarikh Notifikasi)
    col_bi = df.columns[60]   # Column BI (Tarikh Diagnosis)
    col_bp = df.columns[67]   # Column BP (Tempat Notifikasi)
    col_br = df.columns[69]   # Column BR (Status Notifikasi)
    col_bt = df.columns[71]   # Column BT (Kematian)
    col_cd = df.columns[81]   # Column CD (Kategori Tempat Notifikasi)
    col_dt = df.columns[123]  # Column DT (Daerah)
    col_dx = df.columns[127]  # Column DX (Diagnosis)
    col_dy = df.columns[128]  # Column DY (Subdiagnosis)

    # Base filter: official Selangor districts & numeric epiweek
    df_clean = df[df[col_dt].isin(VALID_DISTRICTS)].copy()
    df_clean[col_me] = pd.to_numeric(df_clean[col_me], errors='coerce')
    
    df_semasa = df_clean[df_clean[col_me] == target_me]
    df_kumulatif = df_clean[df_clean[col_me] <= target_me]
    
    # 1. Slide 2 Stats
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
            'pct_daftar_notif': f"{(daftar_notifikasi/total*100):.0f}%" if total else "0%",
            'pct_daftar_kes': f"{(daftar_kes/total*100):.0f}%" if total else "0%",
            'pct_abai': f"{(abai/total*100):.0f}%" if total else "0%",
            'pct_belum': f"{(belum/total*100):.2f}%" if total else "0.00%",
            'pct_batal': f"{(batal/total*100):.2f}%" if total else "0.00%",
        }
        
    stats_semasa = get_stats(df_semasa)
    stats_kumulatif = get_stats(df_kumulatif)

    # 2. Slide 3: Penyakit
    df_clean[col_dx] = df_clean[col_dx].astype(str).str.strip().str.upper()
    df_clean[col_dx] = df_clean[col_dx].replace({'MONKEYPOX': 'MPOX'})
    df_kumu_penyakit = df_clean[df_clean[col_me] <= target_me]
    
    diseases = df_kumu_penyakit[col_dx].unique()
    data_penyakit = []
    for dx in diseases:
        if pd.isna(dx) or dx == 'NAN': continue
        df_dx = df_kumu_penyakit[df_kumu_penyakit[col_dx] == dx]
        is_daftar_kes = df_dx[col_br].astype(str).str.strip().str.upper() == 'DAFTAR KES'
        is_mati = df_dx[col_bt].astype(str).str.strip().str.upper() == 'MATI'
        kumulatif_all_notif = len(df_dx)
        if kumulatif_all_notif == 0: continue
        semasa_daftar_kes = len(df_dx[(df_dx[col_me] == target_me) & is_daftar_kes])
        kumulatif_daftar_kes = len(df_dx[is_daftar_kes])
        kumulatif_mati = len(df_dx[is_mati])
        peratus = (kumulatif_daftar_kes / kumulatif_all_notif * 100)
        data_penyakit.append({
            'Penyakit': dx,
            'Semasa': semasa_daftar_kes,
            'Kumulatif': kumulatif_daftar_kes,
            'Mati': kumulatif_mati,
            'Peratus': peratus
        })
        
    df_penyakit = pd.DataFrame(data_penyakit)
    if not df_penyakit.empty:
        df_penyakit = df_penyakit.sort_values(by=['Kumulatif', 'Penyakit'], ascending=[False, True])

    # 3. Slide 4: District
    df_district_data = []
    for d in VALID_DISTRICTS:
        df_d = df_kumulatif[df_kumulatif[col_dt] == d]
        jml = len(df_d)
        counts = df_d[col_br].value_counts().to_dict()
        d_notif = counts.get('Daftar Notifikasi', 0)
        d_kes = counts.get('Daftar Kes', 0)
        d_abai = counts.get('Abai Notifikasi', 0)
        d_batal = counts.get('Batal Daftar', 0)
        d_belum = counts.get('Belum Ambil Tindakan', 0)
        df_district_data.append({
            'DAERAH': d,
            'Jumlah Notifikasi': jml,
            'Daftar Notifikasi': d_notif,
            'pct_notif': (d_notif / jml * 100) if jml else 0,
            'Daftar Kes': d_kes,
            'pct_kes': (d_kes / jml * 100) if jml else 0,
            'Abai Notifikasi': d_abai,
            'pct_abai': (d_abai / jml * 100) if jml else 0,
            'Batal Daftar': d_batal,
            'pct_batal': (d_batal / jml * 100) if jml else 0,
            'Belum Ambil Tindakan': d_belum,
            'pct_belum': (d_belum / jml * 100) if jml else 0
        })
    df_district = pd.DataFrame(df_district_data)
    if not df_district.empty:
        df_district = df_district.sort_values(by='Jumlah Notifikasi', ascending=False)

    district_names = [d[0] for d in DISTRICT_ABBR]

    # 4. Slide 5: Belum Ambil Tindakan
    df_belum = df_clean[
        (df_clean[col_me] == target_me) & 
        (df_clean[col_br].astype(str).str.strip().str.upper() == 'BELUM AMBIL TINDAKAN')
    ].copy()
    df_belum[col_dx] = df_belum[col_dx].astype(str).str.strip().str.upper()
    df_belum[col_dx] = df_belum[col_dx].replace({'MONKEYPOX': 'MPOX'})

    if not df_belum.empty:
        ct = pd.crosstab(df_belum[col_dx], df_belum[col_dt])
        ct = ct.reindex(columns=district_names, fill_value=0)
        ct['JUM'] = ct.sum(axis=1)
        ct = ct[ct['JUM'] > 0]
        ct = ct.sort_values(by=['JUM', ct.index.name or 'index'], ascending=[False, True]).reset_index()
    else:
        ct = pd.DataFrame(columns=[col_dx] + district_names + ['JUM'])

    # 5. Slide 6: Viral Hepatitis Subdiagnosis (Belum Ambil Tindakan)
    df_hep = df_clean[
        (df_clean[col_me] == target_me) & 
        (df_clean[col_br].astype(str).str.strip().str.upper() == 'BELUM AMBIL TINDAKAN') &
        (df_clean[col_dx].astype(str).str.strip().str.upper() == 'VIRAL HEPATITIS')
    ].copy()
    df_hep[col_dy] = df_hep[col_dy].fillna('').astype(str).str.strip()
    df_hep[col_dy] = df_hep[col_dy].apply(lambda x: 'Tiada subdiagnosis' if x == '' or x.upper() == 'NAN' else x)

    if not df_hep.empty:
        hep_ct = pd.crosstab(df_hep[col_dy], df_hep[col_dt])
        hep_ct = hep_ct.reindex(columns=district_names, fill_value=0)
        hep_ct['JUM'] = hep_ct.sum(axis=1)
        hep_ct = hep_ct[hep_ct['JUM'] > 0]
        hep_ct = hep_ct.sort_values(by=['JUM', hep_ct.index.name or 'index'], ascending=[False, True]).reset_index()
    else:
        hep_ct = pd.DataFrame(columns=[col_dy] + district_names + ['JUM'])

    # 6. Slide 7: Daftar Notifikasi (Inclusion List)
    df_dn = df_clean[
        (df_clean[col_me] == target_me) & 
        (df_clean[col_br].astype(str).str.strip().str.upper() == 'DAFTAR NOTIFIKASI') &
        (df_clean[col_dx].astype(str).str.strip().str.upper().isin(inclusion_tuple))
    ].copy()
    df_dn[col_dx] = df_dn[col_dx].astype(str).str.strip().str.upper()

    if not df_dn.empty:
        dn_ct = pd.crosstab(df_dn[col_dx], df_dn[col_dt])
        dn_ct = dn_ct.reindex(columns=district_names, fill_value=0)
        dn_ct['JUM'] = dn_ct.sum(axis=1)
        dn_ct = dn_ct[dn_ct['JUM'] > 0]
        dn_ct = dn_ct.sort_values(by=['JUM', dn_ct.index.name or 'index'], ascending=[False, True]).reset_index()
    else:
        dn_ct = pd.DataFrame(columns=[col_dx] + district_names + ['JUM'])

    # 7. Slide 8: Daftar Notifikasi (Exclusion List)
    df_dn_exc = df_clean[
        (df_clean[col_me] == target_me) & 
        (df_clean[col_br].astype(str).str.strip().str.upper() == 'DAFTAR NOTIFIKASI') &
        (~df_clean[col_dx].astype(str).str.strip().str.upper().isin(exclusion_tuple))
    ].copy()
    df_dn_exc[col_dx] = df_dn_exc[col_dx].astype(str).str.strip().str.upper()

    if not df_dn_exc.empty:
        dn_exc_ct = pd.crosstab(df_dn_exc[col_dx], df_dn_exc[col_dt])
        dn_exc_ct = dn_exc_ct.reindex(columns=district_names, fill_value=0)
        dn_exc_ct['JUM'] = dn_exc_ct.sum(axis=1)
        dn_exc_ct = dn_exc_ct[dn_exc_ct['JUM'] > 0]
        dn_exc_ct = dn_exc_ct.sort_values(by=['JUM', dn_exc_ct.index.name or 'index'], ascending=[False, True]).reset_index()
    else:
        dn_exc_ct = pd.DataFrame(columns=[col_dx] + district_names + ['JUM'])

    # 8. Slide 9: Viral Hepatitis Subdiagnosis (Daftar Notifikasi)
    df_hep_dn = df_clean[
        (df_clean[col_me] == target_me) & 
        (df_clean[col_br].astype(str).str.strip().str.upper() == 'DAFTAR NOTIFIKASI') &
        (df_clean[col_dx].astype(str).str.strip().str.upper() == 'VIRAL HEPATITIS')
    ].copy()
    df_hep_dn[col_dy] = df_hep_dn[col_dy].fillna('').astype(str).str.strip()
    df_hep_dn[col_dy] = df_hep_dn[col_dy].apply(lambda x: 'Tiada subdiagnosis' if x == '' or x.upper() == 'NAN' else x)

    if not df_hep_dn.empty:
        hep_dn_ct = pd.crosstab(df_hep_dn[col_dy], df_hep_dn[col_dt])
        hep_dn_ct = hep_dn_ct.reindex(columns=district_names, fill_value=0)
        hep_dn_ct['JUM'] = hep_dn_ct.sum(axis=1)
        hep_dn_ct = hep_dn_ct[hep_dn_ct['JUM'] > 0]
        hep_dn_ct = hep_dn_ct.sort_values(by=['JUM', hep_dn_ct.index.name or 'index'], ascending=[False, True]).reset_index()
    else:
        hep_dn_ct = pd.DataFrame(columns=[col_dy] + district_names + ['JUM'])

    # ---------------------------------------------------------
    # 9 & 10. Lewat Notifikasi (24 Jam & 7 Hari)
    # ---------------------------------------------------------
    df_lewat_base = df_clean[df_clean[col_me] == target_me].copy()
    
    # Filter out if Column BP contains 'PKD' or 'PEJABAT'
    bp_series = df_lewat_base[col_bp].fillna('').astype(str).str.upper()
    df_lewat_base = df_lewat_base[~bp_series.str.contains('PKD|PEJABAT')].copy()

    # Parse dates (Column BC - Tarikh Notifikasi, Column BI - Tarikh Diagnosis)
    parsed_bc = df_lewat_base[col_bc].apply(parse_excel_date)
    parsed_bi = df_lewat_base[col_bi].apply(parse_excel_date)

    diff_days = []
    for d_bc, d_bi in zip(parsed_bc, parsed_bi):
        if d_bc and d_bi:
            diff_days.append((d_bc - d_bi).days)
        else:
            diff_days.append(None)
            
    df_lewat_base['DIFF_DAYS'] = diff_days
    df_lewat_base['IS_SWASTA'] = df_lewat_base[col_cd].fillna('').astype(str).str.upper().str.contains('SWASTA')

    # Aggregator for Lewat slides
    def aggregate_lewat(df_subset):
        grouped = {}
        for _, r in df_subset.iterrows():
            dx = r[col_dx]
            dt = r[col_dt]
            is_swasta = r['IS_SWASTA']
            if dx not in grouped:
                grouped[dx] = {d: {'total': 0, 'swasta': 0} for d in VALID_DISTRICTS}
            if dt in grouped[dx]:
                grouped[dx][dt]['total'] += 1
                if is_swasta:
                    grouped[dx][dt]['swasta'] += 1

        res = []
        for dx, dist_dict in grouped.items():
            tot_row = sum(dist_dict[d]['total'] for d in VALID_DISTRICTS)
            swasta_row = sum(dist_dict[d]['swasta'] for d in VALID_DISTRICTS)
            if tot_row > 0:
                row_dict = {'DIAGNOSIS': dx, 'TOTAL': tot_row, 'SWASTA': swasta_row}
                for d in VALID_DISTRICTS:
                    row_dict[f"{d}_tot"] = dist_dict[d]['total']
                    row_dict[f"{d}_swasta"] = dist_dict[d]['swasta']
                res.append(row_dict)
                
        df_res = pd.DataFrame(res)
        if not df_res.empty:
            df_res = df_res.sort_values(by=['TOTAL', 'DIAGNOSIS'], ascending=[False, True]).reset_index(drop=True)
        return df_res

    # 24 Jam: exceed 2 days
    df_24h = df_lewat_base[
        df_lewat_base[col_dx].isin(DIAG_TEMPOH_24H) & 
        (df_lewat_base['DIFF_DAYS'] > 2)
    ].copy()
    lewat_24h_df = aggregate_lewat(df_24h)

    # 7 Hari: exceed 7 days
    df_7d = df_lewat_base[
        df_lewat_base[col_dx].isin(DIAG_TEMPOH_7D) & 
        (df_lewat_base['DIFF_DAYS'] > 7)
    ].copy()
    lewat_7d_df = aggregate_lewat(df_7d)

    return (
        len(df), stats_semasa, stats_kumulatif, df_penyakit, 
        df_district, ct, hep_ct, dn_ct, dn_exc_ct, hep_dn_ct,
        lewat_24h_df, lewat_7d_df
    )

# ---------------------------------------------------------
# PPTX Generator
# ---------------------------------------------------------
@st.cache_data
def generate_pptx(stats_semasa, stats_kumulatif, df_penyakit, df_district, df_belum_ct, df_hep_ct, df_dn_ct, df_dn_exc_ct, df_hep_dn_ct, lewat_24h_df, lewat_7d_df, epi_week, year):
    try:
        remote_buffer = fetch_google_slides_pptx(GOOGLE_SLIDES_ID)
        prs = Presentation(remote_buffer)
    except Exception as e:
        st.warning(f"Could not fetch Google Slides directly: {e}. Generating standalone slides.")
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

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

    # --- Slide 1: Title Slide ---
    if len(prs.slides) > 0:
        slide1 = prs.slides[0]
        for shape in list(slide1.shapes):
            sp = shape._element
            sp.getparent().remove(sp)
    else:
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])

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

    # --- Slide 2: Analisa e-Notifikasi ---
    slide_notif = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_notif.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_notif.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_notif.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(8), Inches(0.8))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Analisa e-Notifikasi"
    p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"ME {epi_week:02d} /{year}"
    p2.font.size = Pt(16); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows, cols = 7, 5
    table_shape = slide_notif.shapes.add_table(rows, cols, Inches(1.15), Inches(1.7), Inches(11.0), Inches(4.7))
    table = table_shape.table
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(2.2)
    table.columns[2].width = Inches(2.2)
    table.columns[3].width = Inches(2.2)
    table.columns[4].width = Inches(2.2)

    headers = ["Minggu Epid", f"ME {epi_week:02d} (Semasa)", "", f"Kumulatif Sehingga ME {epi_week:02d}", ""]
    for i, txt in enumerate(headers):
        write_cell(table.cell(0, i), txt, bold=True, size=17) 
    table.cell(0, 1).merge(table.cell(0, 2))
    table.cell(0, 3).merge(table.cell(0, 4))

    row_labels = ["", "Jumlah Notifikasi", "Daftar Notifikasi", "Daftar Kes", "Abai Notifikasi", "Belum Ambil Tindakan", "Batal Daftar"]
    for i in range(1, 7):
        write_cell(table.cell(i, 0), row_labels[i], bold=True, size=17) 

    write_cell(table.cell(1, 1), f"{stats_semasa['total']:,}", bold=True, size=17)
    write_cell(table.cell(1, 3), f"{stats_kumulatif['total']:,}", bold=True, size=17)
    write_cell(table.cell(2, 1), f"{stats_semasa['daftar_notifikasi']:,}", bold=True, size=17)
    write_cell(table.cell(2, 2), stats_semasa['pct_daftar_notif'], bold=True, size=17)
    write_cell(table.cell(2, 3), f"{stats_kumulatif['daftar_notifikasi']:,}", bold=True, size=17)
    write_cell(table.cell(2, 4), stats_kumulatif['pct_daftar_notif'], bold=True, size=17)
    write_cell(table.cell(3, 1), f"{stats_semasa['daftar_kes']:,}", bold=True, size=17)
    write_cell(table.cell(3, 2), stats_semasa['pct_daftar_kes'], bold=True, size=17)
    write_cell(table.cell(3, 3), f"{stats_kumulatif['daftar_kes']:,}", bold=True, size=17)
    write_cell(table.cell(3, 4), stats_kumulatif['pct_daftar_kes'], bold=True, size=17)
    write_cell(table.cell(4, 1), f"{stats_semasa['abai']:,}", bold=True, size=17)
    write_cell(table.cell(4, 2), stats_semasa['pct_abai'], bold=True, size=17)
    write_cell(table.cell(4, 3), f"{stats_kumulatif['abai']:,}", bold=True, size=17)
    write_cell(table.cell(4, 4), stats_kumulatif['pct_abai'], bold=True, size=17)
    write_cell(table.cell(5, 1), f"{stats_semasa['belum']:,}", bold=True, size=17)
    write_cell(table.cell(5, 2), stats_semasa['pct_belum'], bold=True, size=17)
    write_cell(table.cell(5, 3), f"{stats_kumulatif['belum']:,}", bold=True, size=17)
    write_cell(table.cell(5, 4), stats_kumulatif['pct_belum'], bold=True, size=17)
    write_cell(table.cell(6, 1), f"{stats_semasa['batal']:,}", bold=True, size=17)
    write_cell(table.cell(6, 2), stats_semasa['pct_batal'], bold=True, size=17)
    write_cell(table.cell(6, 3), f"{stats_kumulatif['batal']:,}", bold=True, size=17)
    write_cell(table.cell(6, 4), stats_kumulatif['pct_batal'], bold=True, size=17)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE 
            if i == 0 or j == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            if i == 1 and j in [2, 4]:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0, 0, 0)
                
    myt_zone = datetime.timezone(datetime.timedelta(hours=8))
    now = datetime.datetime.now(myt_zone)
    timestamp_str = now.strftime("%d/%m/%Y @ %I.%M%p").upper()
    
    txBox = slide_notif.shapes.add_textbox(Inches(0.1), Inches(6.9), Inches(10), Inches(0.4))
    p = txBox.text_frame.paragraphs[0]
    p.text = f"(Sumber : Sistem e-notifikasi, KKM muat turun pada ({timestamp_str}))"
    p.font.size = Pt(9); p.font.italic = True; p.font.name = 'Calibri'; p.font.bold = True
    
    bottom_banner = slide_notif.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.8), Inches(6.833), Inches(0.5))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE 

    # --- Slide 3: Bilangan Daftar Kes Mengikut Penyakit ---
    slide_penyakit = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_penyakit.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_penyakit.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_penyakit.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(8), Inches(0.8))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Bilangan Daftar Kes Mengikut Penyakit"
    p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"ME {epi_week:02d} / {year}"
    p2.font.size = Pt(16); p2.font.color.rgb = NAVY; p2.font.bold = True

    total_semasa = df_penyakit['Semasa'].sum() if not df_penyakit.empty else 0
    total_kumulatif = df_penyakit['Kumulatif'].sum() if not df_penyakit.empty else 0
    total_mati = df_penyakit['Mati'].sum() if not df_penyakit.empty else 0
    df_render = df_penyakit.head(19) if len(df_penyakit) > 19 else df_penyakit

    rows = len(df_render) + 2
    cols = 4
    table_shape = slide_penyakit.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(2.2)
    table.columns[2].width = Inches(2.7)
    table.columns[3].width = Inches(3.633)

    headers = ["Penyakit", f"ME {epi_week:02d}", f"≤ ME {epi_week:02d}/ {year}", "Peratus Daftar Kes\n(Bil Daftar Kes / Bil Notifikasi Kes Tersebut)"]
    for j, h in enumerate(headers):
        write_cell(table.cell(0, j), h, bold=True)

    for i, (_, row) in enumerate(df_render.iterrows()):
        r_idx = i + 1
        write_cell(table.cell(r_idx, 0), str(row['Penyakit']), bold=True, align_left=True) 
        write_cell(table.cell(r_idx, 1), f"{int(row['Semasa']):,}", bold=True) 
        
        p = write_cell(table.cell(r_idx, 2), f"{int(row['Kumulatif']):,}", bold=True) 
        if row['Mati'] > 0:
            run = p.add_run()
            run.text = f" ({int(row['Mati'])})"
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True; run.font.size = Pt(10); run.font.name = 'Calibri'
            
        write_cell(table.cell(r_idx, 3), f"{int(row['Peratus'])}%", bold=True) 
        
    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True) 
    write_cell(table.cell(r_idx, 1), f"{int(total_semasa):,}", bold=True)
    p = write_cell(table.cell(r_idx, 2), f"{int(total_kumulatif):,}", bold=True)
    if total_mati > 0:
        run = p.add_run()
        run.text = f" ({int(total_mati)})"
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True; run.font.size = Pt(10); run.font.name = 'Calibri'
    write_cell(table.cell(r_idx, 3), "", bold=True)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    txBox = slide_penyakit.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(6), Inches(0.4))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Sumber data adalah daripada sistem eNotifikasi"
    p.font.size = Pt(9); p.font.bold = True; p.font.name = 'Calibri'; p.font.color.rgb = RGBColor(0, 0, 0) 
    p2 = txBox.text_frame.add_paragraph()
    p2.text = "*(Mati)"
    p2.font.size = Pt(9); p.font.bold = True; p.font.color.rgb = RGBColor(255, 0, 0); p2.font.name = 'Calibri' 
    
    bottom_banner = slide_penyakit.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE 

    # --- Slide 4: Status pencapaian e-Notifikasi Mengikut Daerah ---
    slide_daerah = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_daerah.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_daerah.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_daerah.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(8), Inches(0.8))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Status pencapaian e-Notifikasi"
    p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"Sehingga ME {epi_week:02d} / {year}"
    p2.font.size = Pt(16); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(df_district) + 2
    cols = 12
    table_shape = slide_daerah.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.0))
    table = table_shape.table

    col_widths = [
        Inches(1.4), Inches(1.1), Inches(1.0), Inches(0.85),
        Inches(0.95), Inches(0.85), Inches(1.0), Inches(0.85),
        Inches(0.95), Inches(0.85), Inches(1.0), Inches(0.85)
    ]
    for j, w in enumerate(col_widths):
        table.columns[j].width = w

    headers = [
        "DAERAH", "Jumlah\nNotifikasi", "Daftar\nNotifikasi", "%", 
        "Daftar Kes", "%", "Abai\nNotifikasi", "%", 
        "Batal Daftar", "%", "Belum\nAmbil\nTindakan", "%"
    ]
    for j, h in enumerate(headers):
        write_cell(table.cell(0, j), h, bold=True, size=12)

    tot_jml = df_district['Jumlah Notifikasi'].sum() if not df_district.empty else 0
    tot_notif = df_district['Daftar Notifikasi'].sum() if not df_district.empty else 0
    tot_kes = df_district['Daftar Kes'].sum() if not df_district.empty else 0
    tot_abai = df_district['Abai Notifikasi'].sum() if not df_district.empty else 0
    tot_batal = df_district['Batal Daftar'].sum() if not df_district.empty else 0
    tot_belum = df_district['Belum Ambil Tindakan'].sum() if not df_district.empty else 0

    for i, (_, row) in enumerate(df_district.iterrows()):
        r_idx = i + 1
        write_cell(table.cell(r_idx, 0), str(row['DAERAH']), bold=True, size=12)
        write_cell(table.cell(r_idx, 1), f"{int(row['Jumlah Notifikasi']):,}", bold=True, size=12)
        write_cell(table.cell(r_idx, 2), f"{int(row['Daftar Notifikasi']):,}", bold=True, size=12)
        write_cell(table.cell(r_idx, 3), f"{row['pct_notif']:.2f}%", bold=True, size=12)
        write_cell(table.cell(r_idx, 4), f"{int(row['Daftar Kes']):,}", bold=True, size=12)
        write_cell(table.cell(r_idx, 5), f"{row['pct_kes']:.2f}%", bold=True, size=12)
        write_cell(table.cell(r_idx, 6), f"{int(row['Abai Notifikasi']):,}", bold=True, size=12)
        write_cell(table.cell(r_idx, 7), f"{row['pct_abai']:.2f}%", bold=True, size=12)
        write_cell(table.cell(r_idx, 8), f"{int(row['Batal Daftar']):,}", bold=True, size=12)
        write_cell(table.cell(r_idx, 9), f"{row['pct_batal']:.2f}%", bold=True, size=12)
        write_cell(table.cell(r_idx, 10), f"{int(row['Belum Ambil Tindakan']):,}", bold=True, size=12)
        write_cell(table.cell(r_idx, 11), f"{row['pct_belum']:.2f}%", bold=True, size=12)

    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=12)
    write_cell(table.cell(r_idx, 1), f"{int(tot_jml):,}", bold=True, size=12)
    write_cell(table.cell(r_idx, 2), f"{int(tot_notif):,}", bold=True, size=12)
    write_cell(table.cell(r_idx, 3), f"{(tot_notif/tot_jml*100):.2f}%" if tot_jml else "0.00%", bold=True, size=12)
    write_cell(table.cell(r_idx, 4), f"{int(tot_kes):,}", bold=True, size=12)
    write_cell(table.cell(r_idx, 5), f"{(tot_kes/tot_jml*100):.2f}%" if tot_jml else "0.00%", bold=True, size=12)
    write_cell(table.cell(r_idx, 6), f"{int(tot_abai):,}", bold=True, size=12)
    write_cell(table.cell(r_idx, 7), f"{(tot_abai/tot_jml*100):.2f}%" if tot_jml else "0.00%", bold=True, size=12)
    write_cell(table.cell(r_idx, 8), f"{int(tot_batal):,}", bold=True, size=12)
    write_cell(table.cell(r_idx, 9), f"{(tot_batal/tot_jml*100):.2f}%" if tot_jml else "0.00%", bold=True, size=12)
    write_cell(table.cell(r_idx, 10), f"{int(tot_belum):,}", bold=True, size=12)
    write_cell(table.cell(r_idx, 11), f"{(tot_belum/tot_jml*100):.2f}%" if tot_jml else "0.00%", bold=True, size=12)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i == 0 or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    bottom_banner = slide_daerah.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE 

    # --- Slide 5: Senarai Kes Belum Ambil Tindakan Mengikut Diagnosis ---
    slide_belum = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_belum.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_belum.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_belum.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(9.5), Inches(0.8))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Senarai Kes Belum Ambil Tindakan Mengikut Diagnosis"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"ME {epi_week:02d} / {year}"
    p2.font.size = Pt(16); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(df_belum_ct) + 3
    cols = 11
    table_shape = slide_belum.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(2.333)
    for j in range(1, 11):
        table.columns[j].width = Inches(1.0)

    h0 = ["DIAGNOSIS"] + [d[1] for d in DISTRICT_ABBR] + ["JUM"]
    for j, txt in enumerate(h0):
        write_cell(table.cell(0, j), txt, bold=True, size=16)

    h1 = [""] + [f"ME{epi_week:02d}"] * 10
    for j, txt in enumerate(h1):
        if j > 0:
            write_cell(table.cell(1, j), txt, bold=True, size=16)

    table.cell(0, 0).merge(table.cell(1, 0))
    write_cell(table.cell(0, 0), "DIAGNOSIS", bold=True, size=16)

    district_names = [d[0] for d in DISTRICT_ABBR]
    district_sums = {d: 0 for d in district_names}
    total_jum = 0

    col_dx_name = df_belum_ct.columns[0] if not df_belum_ct.empty else 'DIAGNOSIS'

    for i, (_, row) in enumerate(df_belum_ct.iterrows()):
        r_idx = i + 2
        write_cell(table.cell(r_idx, 0), str(row[col_dx_name]), bold=True, align_left=True, size=16)
        row_jum = 0
        for j, d_name in enumerate(district_names):
            val = int(row[d_name]) if d_name in row else 0
            write_cell(table.cell(r_idx, j + 1), str(val), bold=True, size=16)
            district_sums[d_name] += val
            row_jum += val
        jum_val = int(row['JUM']) if 'JUM' in row else row_jum
        write_cell(table.cell(r_idx, 10), str(jum_val), bold=True, size=16)
        total_jum += jum_val

    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=16)
    for j, d_name in enumerate(district_names):
        write_cell(table.cell(r_idx, j + 1), str(district_sums[d_name]), bold=True, size=16)
    write_cell(table.cell(r_idx, 10), str(total_jum), bold=True, size=16)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i in [0, 1] or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    bottom_banner = slide_belum.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Slide 6: Viral Hepatitis Subdiagnosis Belum Ambil Tindakan ---
    slide_hep = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_hep.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_hep.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_hep.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(9.5), Inches(0.8))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Senarai Kes Belum Ambil Tindakan Mengikut Diagnosis"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"ME {epi_week:02d} / {year}"
    p2.font.size = Pt(16); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(df_hep_ct) + 3
    cols = 11
    table_shape = slide_hep.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(2.333)
    for j in range(1, 11):
        table.columns[j].width = Inches(1.0)

    h0 = ["DIAGNOSIS"] + [d[1] for d in DISTRICT_ABBR] + ["JUM"]
    for j, txt in enumerate(h0):
        write_cell(table.cell(0, j), txt, bold=True, size=16)

    h1 = [""] + [f"ME{epi_week:02d}"] * 10
    for j, txt in enumerate(h1):
        if j > 0:
            write_cell(table.cell(1, j), txt, bold=True, size=16)

    table.cell(0, 0).merge(table.cell(1, 0))
    write_cell(table.cell(0, 0), "DIAGNOSIS", bold=True, size=16)

    district_sums = {d: 0 for d in district_names}
    total_jum = 0

    col_subdx_name = df_hep_ct.columns[0] if not df_hep_ct.empty else 'DIAGNOSIS'

    for i, (_, row) in enumerate(df_hep_ct.iterrows()):
        r_idx = i + 2
        write_cell(table.cell(r_idx, 0), str(row[col_subdx_name]), bold=True, align_left=True, size=16)
        row_jum = 0
        for j, d_name in enumerate(district_names):
            val = int(row[d_name]) if d_name in row else 0
            write_cell(table.cell(r_idx, j + 1), str(val), bold=True, size=16)
            district_sums[d_name] += val
            row_jum += val
        jum_val = int(row['JUM']) if 'JUM' in row else row_jum
        write_cell(table.cell(r_idx, 10), str(jum_val), bold=True, size=16)
        total_jum += jum_val

    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=16)
    for j, d_name in enumerate(district_names):
        write_cell(table.cell(r_idx, j + 1), str(district_sums[d_name]), bold=True, size=16)
    write_cell(table.cell(r_idx, 10), str(total_jum), bold=True, size=16)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i in [0, 1] or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    bottom_banner = slide_hep.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Slide 7: Bilangan Kes Yang Masih Berstatus Daftar Notifikasi (Inclusion List) ---
    slide_dn = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_dn.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_dn.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_dn.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(9.5), Inches(0.9))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Bilangan Kes Yang Masih Berstatus Daftar Notifikasi"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"Tempoh daftar kes ≤ 7 hari"
    p2.font.size = Pt(14); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(df_dn_ct) + 3
    cols = 11
    table_shape = slide_dn.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(2.333)
    for j in range(1, 11):
        table.columns[j].width = Inches(1.0)

    h0 = ["DIAGNOSIS"] + [d[1] for d in DISTRICT_ABBR] + ["JUM"]
    for j, txt in enumerate(h0):
        write_cell(table.cell(0, j), txt, bold=True, size=16)

    h1 = [""] + [f"ME{epi_week:02d}"] * 10
    for j, txt in enumerate(h1):
        if j > 0:
            write_cell(table.cell(1, j), txt, bold=True, size=16)

    table.cell(0, 0).merge(table.cell(1, 0))
    write_cell(table.cell(0, 0), "DIAGNOSIS", bold=True, size=16)

    district_names = [d[0] for d in DISTRICT_ABBR]
    district_sums = {d: 0 for d in district_names}
    total_jum = 0

    col_dn_dx_name = df_dn_ct.columns[0] if not df_dn_ct.empty else 'DIAGNOSIS'

    for i, (_, row) in enumerate(df_dn_ct.iterrows()):
        r_idx = i + 2
        write_cell(table.cell(r_idx, 0), str(row[col_dn_dx_name]), bold=True, align_left=True, size=16)
        row_jum = 0
        for j, d_name in enumerate(district_names):
            val = int(row[d_name]) if d_name in row else 0
            write_cell(table.cell(r_idx, j + 1), str(val), bold=True, size=16)
            district_sums[d_name] += val
            row_jum += val
        jum_val = int(row['JUM']) if 'JUM' in row else row_jum
        write_cell(table.cell(r_idx, 10), str(jum_val), bold=True, size=16)
        total_jum += jum_val

    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=16)
    for j, d_name in enumerate(district_names):
        write_cell(table.cell(r_idx, j + 1), str(district_sums[d_name]), bold=True, size=16)
    write_cell(table.cell(r_idx, 10), str(total_jum), bold=True, size=16)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i in [0, 1] or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    bottom_banner = slide_dn.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Slide 8: Bilangan Kes Yang Masih Berstatus Daftar Notifikasi (Exclusion List) ---
    slide_dn_exc = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_dn_exc.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_dn_exc.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_dn_exc.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(9.5), Inches(0.9))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Bilangan Kes Yang Masih Berstatus Daftar Notifikasi"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"Tempoh daftar kes ≤ 14 hari"
    p2.font.size = Pt(14); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(df_dn_exc_ct) + 3
    cols = 11
    table_shape = slide_dn_exc.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(2.333)
    for j in range(1, 11):
        table.columns[j].width = Inches(1.0)

    h0 = ["DIAGNOSIS"] + [d[1] for d in DISTRICT_ABBR] + ["JUM"]
    for j, txt in enumerate(h0):
        write_cell(table.cell(0, j), txt, bold=True, size=16)

    h1 = [""] + [f"ME{epi_week:02d}"] * 10
    for j, txt in enumerate(h1):
        if j > 0:
            write_cell(table.cell(1, j), txt, bold=True, size=16)

    table.cell(0, 0).merge(table.cell(1, 0))
    write_cell(table.cell(0, 0), "DIAGNOSIS", bold=True, size=16)

    district_sums = {d: 0 for d in district_names}
    total_jum = 0

    col_dn_exc_dx_name = df_dn_exc_ct.columns[0] if not df_dn_exc_ct.empty else 'DIAGNOSIS'

    for i, (_, row) in enumerate(df_dn_exc_ct.iterrows()):
        r_idx = i + 2
        write_cell(table.cell(r_idx, 0), str(row[col_dn_exc_dx_name]), bold=True, align_left=True, size=16)
        row_jum = 0
        for j, d_name in enumerate(district_names):
            val = int(row[d_name]) if d_name in row else 0
            write_cell(table.cell(r_idx, j + 1), str(val), bold=True, size=16)
            district_sums[d_name] += val
            row_jum += val
        jum_val = int(row['JUM']) if 'JUM' in row else row_jum
        write_cell(table.cell(r_idx, 10), str(jum_val), bold=True, size=16)
        total_jum += jum_val

    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=16)
    for j, d_name in enumerate(district_names):
        write_cell(table.cell(r_idx, j + 1), str(district_sums[d_name]), bold=True, size=16)
    write_cell(table.cell(r_idx, 10), str(total_jum), bold=True, size=16)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i in [0, 1] or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    bottom_banner = slide_dn_exc.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Slide 9: Viral Hepatitis Subdiagnosis Daftar Notifikasi ---
    slide_hep_dn = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_hep_dn.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_hep_dn.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_hep_dn.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(9.5), Inches(0.9))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Bilangan Kes Yang Masih Berstatus Daftar Notifikasi"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"Tempoh daftar kes ≤ 14 hari"
    p2.font.size = Pt(14); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(df_hep_dn_ct) + 3
    cols = 11
    table_shape = slide_hep_dn.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.0))
    table = table_shape.table
    table.columns[0].width = Inches(2.333)
    for j in range(1, 11):
        table.columns[j].width = Inches(1.0)

    h0 = ["DIAGNOSIS"] + [d[1] for d in DISTRICT_ABBR] + ["JUM"]
    for j, txt in enumerate(h0):
        write_cell(table.cell(0, j), txt, bold=True, size=16)

    h1 = [""] + [f"ME{epi_week:02d}"] * 10
    for j, txt in enumerate(h1):
        if j > 0:
            write_cell(table.cell(1, j), txt, bold=True, size=16)

    table.cell(0, 0).merge(table.cell(1, 0))
    write_cell(table.cell(0, 0), "DIAGNOSIS", bold=True, size=16)

    district_sums = {d: 0 for d in district_names}
    total_jum = 0

    col_hep_dn_dx_name = df_hep_dn_ct.columns[0] if not df_hep_dn_ct.empty else 'DIAGNOSIS'

    for i, (_, row) in enumerate(df_hep_dn_ct.iterrows()):
        r_idx = i + 2
        write_cell(table.cell(r_idx, 0), str(row[col_hep_dn_dx_name]), bold=True, align_left=True, size=16)
        row_jum = 0
        for j, d_name in enumerate(district_names):
            val = int(row[d_name]) if d_name in row else 0
            write_cell(table.cell(r_idx, j + 1), str(val), bold=True, size=16)
            district_sums[d_name] += val
            row_jum += val
        jum_val = int(row['JUM']) if 'JUM' in row else row_jum
        write_cell(table.cell(r_idx, 10), str(jum_val), bold=True, size=16)
        total_jum += jum_val

    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=16)
    for j, d_name in enumerate(district_names):
        write_cell(table.cell(r_idx, j + 1), str(district_sums[d_name]), bold=True, size=16)
    write_cell(table.cell(r_idx, 10), str(total_jum), bold=True, size=16)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i in [0, 1] or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    bottom_banner = slide_hep_dn.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Slide 10: Lewat Notifikasi 24 Jam ---
    slide_l24 = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_l24.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_l24.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_l24.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(9.5), Inches(0.9))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Lewat Notifikasi 24 Jam"
    p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"ME {epi_week:02d} / {year}"
    p2.font.size = Pt(16); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(lewat_24h_df) + 2
    cols = 11
    table_shape = slide_l24.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(4.5))
    table = table_shape.table
    table.columns[0].width = Inches(2.0)
    for j in range(1, 11):
        table.columns[j].width = Inches(1.033)

    h_lewat = ["Diagnosis", "GOMBAK", "HULU\nLANGAT", "HULU\nSELANGOR", "KLANG", "KUALA\nLANGAT", "KUALA\nSELANGOR", "PETALING", "SABAK\nBERNAM", "SEPANG", "JUMLAH"]
    for j, h in enumerate(h_lewat):
        write_cell(table.cell(0, j), h, bold=True, size=12)

    dist_sums_24 = {d: {'total': 0, 'swasta': 0} for d in VALID_DISTRICTS}
    total_lewat_24 = 0
    swasta_lewat_24 = 0

    for i, (_, r) in enumerate(lewat_24h_df.iterrows()):
        r_idx = i + 1
        write_cell(table.cell(r_idx, 0), str(r['DIAGNOSIS']), bold=True, align_left=True, size=12)
        for j, d in enumerate(VALID_DISTRICTS):
            tot = int(r[f"{d}_tot"])
            sw = int(r[f"{d}_swasta"])
            dist_sums_24[d]['total'] += tot
            dist_sums_24[d]['swasta'] += sw
            write_cell(table.cell(r_idx, j + 1), format_cell_stat(tot, sw, zero_as_dash=True), bold=True, size=12)
            
        r_tot = int(r['TOTAL'])
        r_sw = int(r['SWASTA'])
        total_lewat_24 += r_tot
        swasta_lewat_24 += r_sw
        write_cell(table.cell(r_idx, 10), format_cell_stat(r_tot, r_sw, zero_as_dash=True), bold=True, size=12)

    # JUMLAH Row
    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=12)
    for j, d in enumerate(VALID_DISTRICTS):
        tot = dist_sums_24[d]['total']
        sw = dist_sums_24[d]['swasta']
        write_cell(table.cell(r_idx, j + 1), format_cell_stat(tot, sw, zero_as_dash=False), bold=True, size=12)
    write_cell(table.cell(r_idx, 10), format_cell_stat(total_lewat_24, swasta_lewat_24, zero_as_dash=False), bold=True, size=12)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i == 0 or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    txBox = slide_l24.shapes.add_textbox(Inches(0.5), Inches(6.85), Inches(6), Inches(0.4))
    p = txBox.text_frame.paragraphs[0]
    p.text = "( ) Kemudahan Kesihatan Swasta"
    p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = RGBColor(0, 0, 0)

    bottom_banner = slide_l24.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Slide 11: Lewat Notifikasi 7 Hari ---
    slide_l7 = prs.slides.add_slide(prs.slide_layouts[6])

    if os.path.exists("logo.png"):
        slide_l7.shapes.add_picture("logo.png", Inches(0.6), Inches(0.3), width=Inches(1.5))

    line = slide_l7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(0.4), Inches(0.03), Inches(0.9))
    line.fill.solid(); line.fill.fore_color.rgb = NAVY; line.line.fill.background()

    txBox = slide_l7.shapes.add_textbox(Inches(2.4), Inches(0.35), Inches(9.5), Inches(0.9))
    p = txBox.text_frame.paragraphs[0]
    p.text = "Lewat Notifikasi 7 Hari"
    p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = NAVY
    
    p2 = txBox.text_frame.add_paragraph()
    p2.text = f"ME {epi_week:02d} / {year}"
    p2.font.size = Pt(16); p2.font.color.rgb = NAVY; p2.font.bold = True

    rows = len(lewat_7d_df) + 2
    cols = 11
    table_shape = slide_l7.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.333), Inches(4.5))
    table = table_shape.table
    table.columns[0].width = Inches(2.0)
    for j in range(1, 11):
        table.columns[j].width = Inches(1.033)

    for j, h in enumerate(h_lewat):
        write_cell(table.cell(0, j), h, bold=True, size=12)

    dist_sums_7 = {d: {'total': 0, 'swasta': 0} for d in VALID_DISTRICTS}
    total_lewat_7 = 0
    swasta_lewat_7 = 0

    for i, (_, r) in enumerate(lewat_7d_df.iterrows()):
        r_idx = i + 1
        write_cell(table.cell(r_idx, 0), str(r['DIAGNOSIS']), bold=True, align_left=True, size=12)
        for j, d in enumerate(VALID_DISTRICTS):
            tot = int(r[f"{d}_tot"])
            sw = int(r[f"{d}_swasta"])
            dist_sums_7[d]['total'] += tot
            dist_sums_7[d]['swasta'] += sw
            write_cell(table.cell(r_idx, j + 1), format_cell_stat(tot, sw, zero_as_dash=True), bold=True, size=12)
            
        r_tot = int(r['TOTAL'])
        r_sw = int(r['SWASTA'])
        total_lewat_7 += r_tot
        swasta_lewat_7 += r_sw
        write_cell(table.cell(r_idx, 10), format_cell_stat(r_tot, r_sw, zero_as_dash=True), bold=True, size=12)

    # JUMLAH Row
    r_idx = rows - 1
    write_cell(table.cell(r_idx, 0), "JUMLAH", bold=True, size=12)
    for j, d in enumerate(VALID_DISTRICTS):
        tot = dist_sums_7[d]['total']
        sw = dist_sums_7[d]['swasta']
        write_cell(table.cell(r_idx, j + 1), format_cell_stat(tot, sw, zero_as_dash=False), bold=True, size=12)
    write_cell(table.cell(r_idx, 10), format_cell_stat(total_lewat_7, swasta_lewat_7, zero_as_dash=False), bold=True, size=12)

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            set_cell_border(cell, NAVY)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i == 0 or j == 0 or i == len(table.rows) - 1:
                cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GREY
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    txBox = slide_l7.shapes.add_textbox(Inches(0.5), Inches(6.85), Inches(6), Inches(0.4))
    p = txBox.text_frame.paragraphs[0]
    p.text = "( ) Kemudahan Kesihatan Swasta"
    p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = RGBColor(0, 0, 0)

    bottom_banner = slide_l7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(6.9), Inches(6.833), Inches(0.4))
    bottom_banner.fill.solid(); bottom_banner.fill.fore_color.rgb = NAVY; bottom_banner.line.fill.background()
    p = bottom_banner.text_frame.paragraphs[0]
    p.text = "UNIT SURVELAN & KESIAPSIAGAAN, JABATAN KESIHATAN NEGERI SELANGOR"
    p.font.size = Pt(9); p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.RIGHT; p.font.name = 'Calibri'; p.font.bold = True
    bottom_banner.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# UI Runner
# ---------------------------------------------------------
st.title("📊 Selangor Epi Review Slide Generator")
st.write(f"Target Output: **ME {epi_week:02d} / {year}** (Malaysia Time UTC+8)")

st.divider()
st.subheader("1. Upload Data")
uploaded_file = st.file_uploader("Upload raw Excel data for Analisa e-Notifikasi", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner("Processing data and generating slides automatically..."):
        (
            total_rows, stats_semasa, stats_kumulatif, df_penyakit, 
            df_district, df_belum_ct, df_hep_ct, df_dn_ct, df_dn_exc_ct, 
            df_hep_dn_ct, lewat_24h_df, lewat_7d_df
        ) = load_and_process_data(
            uploaded_file, epi_week, INCLUSION_DIAGNOSES, EXCLUSION_DIAGNOSES_14
        )
        
        pptx_buffer = generate_pptx(
            stats_semasa, stats_kumulatif, df_penyakit, df_district, 
            df_belum_ct, df_hep_ct, df_dn_ct, df_dn_exc_ct, df_hep_dn_ct, 
            lewat_24h_df, lewat_7d_df, epi_week, year
        )
        
        st.success(f"Successfully processed {total_rows:,} records. Slide deck is ready!")

        filename = f"Selangor_Epi_Review_ME{epi_week:02d}_{year}.pptx"
        
        st.download_button(
            label="📥 Download Presentation (.pptx)",
            data=pptx_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
else:
    st.info("⚠️ Please upload the Excel file. The presentation will automatically generate once uploaded.")
