<<<<<<< SEARCH
# ---------------------------------------------------------
# UI Runner
# ---------------------------------------------------------
st.divider()

if uploaded_file:
=======
# ---------------------------------------------------------
# UI Runner
# ---------------------------------------------------------
st.title("📊 Selangor Epi Review Slide Generator")
st.write(f"Target Output: **ME {epi_week:02d} / {year}** (Malaysia Time UTC+8)")

st.divider()
st.subheader("1. Upload Data")
uploaded_file = st.file_uploader("Upload raw Excel data for Analisa e-Notifikasi", type=["xlsx", "xls"])

if uploaded_file:
>>>>>>> REPLACE
```

### What was fixed:
- Declared `uploaded_file = st.file_uploader(...)` immediately before `if uploaded_file:`.
- Re-added the app header and title block so the page displays cleanly and avoids the `NameError`.
