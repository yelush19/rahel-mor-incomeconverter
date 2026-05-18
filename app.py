import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="ממיר הכנסות | רחל מור",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700&display=swap');

/* ── Base ── */
html, body, .stApp {
    direction: rtl !important;
    font-family: 'Heebo', 'Segoe UI', sans-serif !important;
    background-color: #f4f6f8 !important;
    color: #1a2b3c !important;
}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 3px solid #27ae60 !important;
    box-shadow: -2px 0 8px rgba(0,0,0,0.06) !important;
    direction: rtl !important;
}
section[data-testid="stSidebar"] * { color: #2c3e50 !important; }
section[data-testid="stSidebar"] .stRadio label {
    font-size: 15px !important;
    font-weight: 500 !important;
}

/* ── Main content ── */
.main .block-container {
    padding: 2rem 3rem !important;
    direction: rtl !important;
    max-width: 1100px !important;
}

/* ── Headings ── */
h1 {
    color: #1a2b3c !important;
    font-weight: 700 !important;
    font-size: 1.9rem !important;
    border-bottom: 3px solid #27ae60 !important;
    padding-bottom: 10px !important;
    margin-bottom: 20px !important;
}
h2, h3 { color: #2c3e50 !important; font-weight: 600 !important; }
h4      { color: #555 !important; }

/* ── All text alignment — scoped, not global ── */
.main p, .main div, .main label, .main span, .main li,
.main td, .main th, .main input, .main textarea,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span {
    direction: rtl !important;
    text-align: right !important;
    font-family: 'Heebo', 'Segoe UI', sans-serif !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    background: #ffffff !important;
    border: 2px dashed #27ae60 !important;
    border-radius: 12px !important;
    padding: 20px !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
}

/* ── Primary button ── */
.stButton > button {
    background: linear-gradient(135deg, #27ae60, #2ecc71) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Heebo', sans-serif !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    padding: 10px 28px !important;
    box-shadow: 0 3px 10px rgba(39,174,96,0.25) !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #219a55, #27ae60) !important;
    box-shadow: 0 5px 15px rgba(39,174,96,0.35) !important;
    transform: translateY(-1px) !important;
}

/* ── Download button ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #27ae60, #2ecc71) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 17px !important;
    font-weight: 700 !important;
    padding: 14px 28px !important;
    width: 100% !important;
    box-shadow: 0 4px 15px rgba(39,174,96,0.3) !important;
    letter-spacing: 0.3px !important;
}
.stDownloadButton > button:hover {
    box-shadow: 0 6px 20px rgba(39,174,96,0.45) !important;
    transform: translateY(-1px) !important;
}

/* ── Metrics ── */
[data-testid="stMetric"] {
    background: #ffffff !important;
    border: 1px solid #e0e7ef !important;
    border-top: 3px solid #27ae60 !important;
    border-radius: 10px !important;
    padding: 16px !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06) !important;
}
[data-testid="stMetricValue"] { color: #27ae60 !important; font-size: 2rem !important; font-weight: 700 !important; }
[data-testid="stMetricLabel"] { color: #7f8c8d !important; font-size: 13px !important; }

/* ── Cards / Expander ── */
[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid #e0e7ef !important;
    border-radius: 10px !important;
    box-shadow: 0 2px 6px rgba(0,0,0,0.05) !important;
}

/* ── Divider ── */
hr { border-color: #e0e7ef !important; }

/* ── Input fields ── */
input, textarea, select {
    background: #ffffff !important;
    border: 1px solid #d0d9e3 !important;
    color: #1a2b3c !important;
    border-radius: 6px !important;
}

/* ── Tabs ── */
[data-testid="stTabs"] button {
    color: #7f8c8d !important;
    font-family: 'Heebo', sans-serif !important;
    font-size: 15px !important;
    font-weight: 500 !important;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #27ae60 !important;
    border-bottom: 3px solid #27ae60 !important;
    font-weight: 700 !important;
}

/* ── Info / Success / Error ── */
[data-testid="stAlert"] { direction: rtl !important; border-radius: 8px !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid #e0e7ef !important;
    border-radius: 8px !important;
    box-shadow: 0 2px 6px rgba(0,0,0,0.05) !important;
}
</style>
""", unsafe_allow_html=True)

from utils.db import (
    get_all_clients, add_client, delete_client, get_next_account_number,
    import_clients_from_df, get_account_columns, add_account_column,
)
from utils.converter import convert_income_file, create_excel_output, detect_month_label

# ── Sidebar ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 20px 0 10px 0;'>
        <div style='font-size:2rem;'>📊</div>
        <div style='font-size:1.1rem; font-weight:700; color:#1a2b3c;'>ממיר הכנסות</div>
        <div style='font-size:0.8rem; color:#888; margin-top:2px;'>רחל מור, עו"ד</div>
    </div>
    <hr style='border-color:#e0e7ef; margin:10px 0 20px 0;'>
    """, unsafe_allow_html=True)

    if "page" not in st.session_state:
        st.session_state.page = "עיבוד"

    btn_style_active   = "background:#27ae60;color:white;border:none;border-radius:8px;padding:12px 20px;width:100%;font-size:15px;font-weight:700;cursor:pointer;margin-bottom:8px;font-family:'Heebo',sans-serif;"
    btn_style_inactive = "background:#f4f6f8;color:#2c3e50;border:1px solid #e0e7ef;border-radius:8px;padding:12px 20px;width:100%;font-size:15px;font-weight:500;cursor:pointer;margin-bottom:8px;font-family:'Heebo',sans-serif;"

    if st.button("📊  עיבוד חודשי", use_container_width=True,
                 type="primary" if st.session_state.page == "עיבוד" else "secondary"):
        st.session_state.page = "עיבוד"
        st.rerun()

    if st.button("⚙️  הגדרות", use_container_width=True,
                 type="primary" if st.session_state.page == "הגדרות" else "secondary"):
        st.session_state.page = "הגדרות"
        st.rerun()

page = st.session_state.page

# ════════════════════════════════════════════════════════════
#  PAGE 1 — Monthly Processing
# ════════════════════════════════════════════════════════════
if page == "עיבוד":
    st.title("📊 ממיר הכנסות → חשבשבת")

    uploaded_files = st.file_uploader(
        "העלי קובץ הכנסות חודשי (.xlsx) — אפשר כמה קבצים יחד",
        type=["xlsx"],
        accept_multiple_files=True,
        help="לדוגמה: רחל_מור_הכנסות_3_2026.xlsx",
    )

    if not uploaded_files:
        st.info("⬆️ העלי קובץ אחד או יותר כדי להתחיל")
        st.stop()

    for uploaded in uploaded_files:
        st.divider()
        df = pd.read_excel(uploaded, header=None)
        month_label = detect_month_label(uploaded.name, df)
        st.success(f"✅ **{uploaded.name}** | חודש: **{month_label}**")

    st.divider()

    if st.button("🚀 עבד והכן קבצי ייבוא", type="primary", use_container_width=True):
        clients_dict = get_all_clients()
        extra_cols   = get_account_columns()

        all_unmatched    = []
        all_unknown_cols = []

        # First pass — collect all issues across all files
        for uploaded in uploaded_files:
            df = pd.read_excel(uploaded, header=None)
            _, _, unmatched, unknown_cols = convert_income_file(df, clients_dict, extra_cols)
            for c in unmatched:
                if c not in all_unmatched:
                    all_unmatched.append(c)
            for c in unknown_cols:
                if c not in all_unknown_cols:
                    all_unknown_cols.append(c)

        # ── Handle unknown columns ──────────────────────────
        if all_unknown_cols:
            st.error("⚠️ עמודות לא מוכרות — הגדירי חשבון לכל אחת")
            for col_name in all_unknown_cols:
                with st.form(f"col_form_{col_name}"):
                    st.markdown(f"#### עמודה: `{col_name}`")
                    c1, c2 = st.columns([2, 1])
                    acct   = c1.number_input("מספר חשבון", min_value=1000, max_value=9999, step=1)
                    exempt = c2.checkbox('פטור ממע"מ')
                    if st.form_submit_button("💾 שמור"):
                        add_account_column(col_name, int(acct), exempt)
                        st.success("נשמר! לחצי שוב על 'עבד'")
                        st.rerun()

        # ── Handle unmatched clients ────────────────────────
        if all_unmatched:
            st.error(f"⚠️ {len(all_unmatched)} לקוחות לא נמצאו באינדקס")
            next_acct = get_next_account_number()
            for i, client_name in enumerate(all_unmatched):
                with st.form(f"client_form_{client_name}"):
                    st.markdown(f"**{client_name}**")
                    acct = st.number_input(
                        "מספר חשבון (3000–3999)",
                        min_value=3000, max_value=3999, step=1,
                        value=next_acct + i,
                    )
                    if st.form_submit_button("💾 שמור"):
                        add_client(client_name, int(acct))
                        st.success("נשמר! לחצי שוב על 'עבד'")
                        st.rerun()

        # ── All good — generate one file per month ───────────
        if not all_unmatched and not all_unknown_cols:
            clients_dict = get_all_clients()
            extra_cols   = get_account_columns()

            for uploaded in uploaded_files:
                df          = pd.read_excel(uploaded, header=None)
                month_label = detect_month_label(uploaded.name, df)

                invoice_rows, receipt_rows, _, _ = convert_income_file(df, clients_dict, extra_cols)

                c1, c2, c3 = st.columns(3)
                c1.metric(f"חשבוניות — {month_label}", len(invoice_rows))
                c2.metric(f"קבלות — {month_label}",    len(receipt_rows))
                c3.metric('סה"כ',                       len(invoice_rows) + len(receipt_rows))

                excel_data  = create_excel_output(invoice_rows, receipt_rows, month_label)
                output_name = f"Hashavshevet_{month_label.replace('.', '_')}.xlsx"

                st.download_button(
                    label=f"⬇️ הורד — {month_label}",
                    data=excel_data,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                    key=f"dl_{month_label}",
                )


# ════════════════════════════════════════════════════════════
#  PAGE 2 — Settings
# ════════════════════════════════════════════════════════════
elif page == "הגדרות":
    st.title("⚙️ הגדרות")
    tab1, tab2 = st.tabs(["👥 לקוחות", "🔢 חשבונות / עמודות"])

    # ── Tab 1: Clients ──────────────────────────────────────
    with tab1:
        st.subheader("ייבוא ראשוני — אינדקס לקוחות")
        idx_file = st.file_uploader(
            "CLients_Index_Rachel.xlsx (מאזן בוחן)",
            type=["xlsx"],
            key="idx_uploader",
        )
        if idx_file:
            if st.button("📥 ייבא לקוחות"):
                df_idx = pd.read_excel(idx_file, header=None)
                count  = import_clients_from_df(df_idx)
                st.success(f"יובאו {count} לקוחות ✅")
                st.rerun()

        st.divider()
        st.subheader("הוסף לקוח חדש")
        with st.form("add_client_form"):
            c1, c2 = st.columns([3, 1])
            new_name = c1.text_input("שם לקוח")
            new_acct = c2.number_input("חשבון", min_value=3000, max_value=3999, step=1)
            if st.form_submit_button("➕ הוסף"):
                if new_name.strip():
                    add_client(new_name.strip(), int(new_acct))
                    st.success(f"נוסף: {new_name} → {new_acct}")
                    st.rerun()

        st.divider()
        st.subheader("כל הלקוחות")
        clients = get_all_clients()
        if clients:
            df_display = pd.DataFrame(
                [(acct, name) for name, acct in sorted(clients.items(), key=lambda x: x[1])],
                columns=["מספר חשבון", "שם לקוח"],
            )
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            st.caption(f'סה"כ: {len(clients)} לקוחות')
        else:
            st.info("אין לקוחות — ייבאי אינדקס למעלה")

    # ── Tab 2: Account Columns ──────────────────────────────
    with tab2:
        st.subheader("חשבונות מובנים")
        st.table(pd.DataFrame([
            {"עמודה": "שכט",     "חשבון": 5000, 'מע"מ': "חייב 18%"},
            {"עמודה": "נוטריון", "חשבון": 5004, 'מע"מ': "חייב 18%"},
            {"עמודה": "הוצאות",  "חשבון": 5002, 'מע"מ': "פטור"},
        ]))

        st.divider()
        st.subheader("עמודות מותאמות אישית")
        extra = get_account_columns()
        if extra:
            for col, cfg in extra.items():
                vat_label = 'פטור ממע"מ' if cfg["vat_exempt"] else 'חייב מע"מ'
                st.write(f"**{col}** → חשבון {cfg['account']} | {vat_label}")
        else:
            st.info("אין עדיין — תופיע כאן אם תהיה עמודה לא מוכרת בקובץ")
