import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="ממיר הכנסות | רחל מור",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@400;500;600;700&display=swap');

.main .block-container {
    direction: rtl;
    font-family: 'Heebo', sans-serif;
    background: #f4f6f8;
    padding: 2rem 2.5rem;
    max-width: 1100px;
}
.main h1 { color: #1a2b3c; font-weight: 700; border-bottom: 3px solid #27ae60; padding-bottom: 8px; }
.main h2, .main h3 { color: #2c3e50; font-weight: 600; }
.main p, .main label, .main div.stMarkdown { direction: rtl; text-align: right; font-family: 'Heebo', sans-serif; }

.stButton > button {
    background: linear-gradient(135deg, #27ae60, #2ecc71) !important;
    color: white !important; border: none !important; border-radius: 8px !important;
    font-family: 'Heebo', sans-serif !important; font-size: 15px !important;
    font-weight: 600 !important; box-shadow: 0 3px 10px rgba(39,174,96,0.25) !important;
}
.stButton > button[kind="secondary"] {
    background: transparent !important; color: #aab8c2 !important;
    border: none !important; box-shadow: none !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #27ae60, #2ecc71) !important;
    color: white !important; border: none !important; border-radius: 10px !important;
    font-size: 17px !important; font-weight: 700 !important; width: 100% !important;
    box-shadow: 0 4px 15px rgba(39,174,96,0.3) !important;
    font-family: 'Heebo', sans-serif !important;
}
[data-testid="stMetric"] {
    background: white; border: 1px solid #e0e7ef; border-top: 3px solid #27ae60;
    border-radius: 10px; padding: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
[data-testid="stMetricValue"] { color: #27ae60 !important; font-weight: 700 !important; }
[data-testid="stFileUploader"] { background: white; border: 2px dashed #27ae60 !important; border-radius: 12px; }
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #27ae60 !important; border-bottom: 3px solid #27ae60 !important; font-weight: 700 !important;
}
section[data-testid="stSidebar"] > div { background: #1a2b3c; padding-top: 1rem; }
section[data-testid="stSidebar"] .stButton > button {
    background: transparent !important; color: #c9d1d9 !important;
    border: none !important; border-radius: 8px !important;
    font-size: 15px !important; box-shadow: none !important;
    padding: 10px 16px !important; width: 100% !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(39,174,96,0.15) !important; color: #2ecc71 !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #27ae60 !important; color: white !important;
}
</style>
""", unsafe_allow_html=True)

from utils.db import (
    get_all_clients, get_clients_full, add_client, delete_client, get_next_account_number,
    import_clients_from_df, get_account_columns, add_account_column,
)
from utils.converter import convert_income_file, create_excel_output, detect_month_label

with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding:20px 0 16px; color:white; font-family:Heebo,sans-serif;'>
        <div style='font-size:2.2rem;'>📊</div>
        <div style='font-size:1rem; font-weight:700; margin-top:4px;'>ממיר הכנסות</div>
        <div style='font-size:0.78rem; color:#8b9bb4; margin-top:2px;'>רחל מור, עו"ד</div>
    </div>
    <hr style='border-color:#2d3f55; margin:0 12px 16px;'>
    """, unsafe_allow_html=True)

    if "page" not in st.session_state:
        st.session_state.page = "עיבוד"

    if st.button("📊   עיבוד חודשי", use_container_width=True,
                 type="primary" if st.session_state.page == "עיבוד" else "secondary"):
        st.session_state.page = "עיבוד"
        st.rerun()

    if st.button("⚙️   הגדרות", use_container_width=True,
                 type="primary" if st.session_state.page == "הגדרות" else "secondary"):
        st.session_state.page = "הגדרות"
        st.rerun()

page = st.session_state.page

if page == "עיבוד":
    st.title("📊 ממיר הכנסות ← חשבשבת")

    uploaded_files = st.file_uploader(
        "העלי קובץ הכנסות חודשי (.xlsx) — אפשר כמה קבצים יחד",
        type=["xlsx"], accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("⬆️ העלי קובץ אחד או יותר כדי להתחיל")
        st.session_state.processing = False
        st.stop()

    for uploaded in uploaded_files:
        df = pd.read_excel(uploaded, header=None)
        month_label = detect_month_label(uploaded.name, df)
        st.success(f"✅ **{uploaded.name}** | חודש: **{month_label}**")

    st.divider()

    if "processing" not in st.session_state:
        st.session_state.processing = False

    if st.button("🚀 עבד והכן קבצי ייבוא", type="primary", use_container_width=True):
        st.session_state.processing = True

    if st.session_state.processing:
        clients_dict = get_all_clients()
        extra_cols   = get_account_columns()
        all_unmatched, all_unknown_cols = [], []

        for uploaded in uploaded_files:
            df = pd.read_excel(uploaded, header=None)
            _, _, unmatched, unknown_cols = convert_income_file(df, clients_dict, extra_cols)
            for c in unmatched:
                if c not in all_unmatched: all_unmatched.append(c)
            for c in unknown_cols:
                if c not in all_unknown_cols: all_unknown_cols.append(c)

        if all_unknown_cols:
            st.error("⚠️ עמודות לא מוכרות — הגדירי חשבון")
            for col_name in all_unknown_cols:
                with st.form(f"col_{col_name}"):
                    st.markdown(f"**עמודה: {col_name}**")
                    c1, c2 = st.columns([2, 1])
                    acct   = c1.number_input("חשבון", min_value=1000, max_value=9999, step=1)
                    exempt = c2.checkbox('פטור ממע"מ')
                    if st.form_submit_button("💾 שמור"):
                        add_account_column(col_name, int(acct), exempt)
                        st.rerun()

        if all_unmatched:
            st.error(f"⚠️ {len(all_unmatched)} לקוחות לא נמצאו — הזיני מספר חשבון")
            next_acct = get_next_account_number()
            for i, client_name in enumerate(all_unmatched):
                with st.form(f"cl_{client_name}"):
                    st.markdown(f"**{client_name}**")
                    acct = st.number_input("מספר חשבון", min_value=3000, max_value=3999,
                                           step=1, value=next_acct + i)
                    if st.form_submit_button("💾 שמור"):
                        add_client(client_name, int(acct))
                        st.rerun()

        if not all_unmatched and not all_unknown_cols:
            clients_dict = get_all_clients()
            extra_cols   = get_account_columns()
            for uploaded in uploaded_files:
                df          = pd.read_excel(uploaded, header=None)
                month_label = detect_month_label(uploaded.name, df)
                invoice_rows, receipt_rows, _, _ = convert_income_file(df, clients_dict, extra_cols)
                st.divider()
                c1, c2, c3 = st.columns(3)
                c1.metric("חשבוניות", len(invoice_rows))
                c2.metric("קבלות",    len(receipt_rows))
                c3.metric("חודש",     month_label)
                excel_data = create_excel_output(invoice_rows, receipt_rows, month_label)
                st.download_button(
                    label=f"⬇️ הורד ייבוא — {month_label}",
                    data=excel_data,
                    file_name=f"Hashavshevet_{month_label.replace('.', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, type="primary", key=f"dl_{month_label}",
                )

elif page == "הגדרות":
    st.title("⚙️ הגדרות")
    tab1, tab2 = st.tabs(["👥 לקוחות", "🔢 חשבונות"])

    with tab1:
        st.subheader("ייבוא מאינדקס")
        idx_file = st.file_uploader("CLients_Index_Rachel.xlsx", type=["xlsx"], key="idx")
        if idx_file:
            if st.button("📥 ייבא לקוחות"):
                count = import_clients_from_df(pd.read_excel(idx_file, header=None))
                st.success(f"יובאו {count} לקוחות ✅")
                st.rerun()
        st.divider()
        st.subheader("הוסף לקוח חדש")
        with st.form("add_client"):
            c1, c2, c3 = st.columns([3, 1, 2])
            name   = c1.text_input("שם לקוח")
            acct   = c2.number_input("חשבון", min_value=3000, max_value=3999,
                                     step=1, value=get_next_account_number())
            id_num = c3.text_input("ת.ז / מס.ע.מ")
            if st.form_submit_button("➕ הוסף"):
                if name.strip():
                    add_client(name.strip(), int(acct), id_num.strip())
                    st.rerun()
        st.divider()
        st.subheader("רשימת לקוחות")
        rows = get_clients_full()
        if rows:
            df_show = pd.DataFrame(rows)[["account_number", "name", "id_number"]]
            df_show.columns = ["חשבון", "שם לקוח", "ת.ז / מס.ע.מ"]
            df_show["ת.ז / מס.ע.מ"] = df_show["ת.ז / מס.ע.מ"].fillna("").astype(str)
            st.dataframe(df_show, use_container_width=True, hide_index=True)
            st.caption(f'סה"כ {len(rows)} לקוחות')
        else:
            st.info("אין לקוחות — ייבאי אינדקס למעלה")

    with tab2:
        st.subheader("חשבונות מובנים")
        st.table(pd.DataFrame([
            {"עמודה": "שכט",     "חשבון": 5000, 'מע"מ': "חייב 18%"},
            {"עמודה": "נוטריון", "חשבון": 5004, 'מע"מ': "חייב 18%"},
            {"עמודה": "הוצאות",  "חשבון": 5002, 'מע"מ': "פטור"},
        ]))
        st.divider()
        st.subheader("עמודות מותאמות")
        extra = get_account_columns()
        if extra:
            for col, cfg in extra.items():
                st.write(f"**{col}** → {cfg['account']} | {'פטור' if cfg['vat_exempt'] else 'חייב'} מע\"מ")
        else:
            st.info("אין עדיין")
