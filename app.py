import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="ממיר הכנסות | רחל מור",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
body, .stApp, p, div, label, span, h1, h2, h3, h4 {
    direction: rtl !important;
    text-align: right !important;
    font-family: 'Segoe UI', Arial, sans-serif !important;
}
section[data-testid="stSidebar"] { direction: rtl; }
.stDownloadButton > button {
    background-color: #28a745 !important;
    color: white !important;
    font-size: 18px !important;
    padding: 14px !important;
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

from utils.db import (
    get_all_clients, add_client, delete_client,
    import_clients_from_df, get_account_columns, add_account_column,
)
from utils.converter import convert_income_file, create_excel_output, detect_month_label

# ── Navigation ──────────────────────────────────────────────
page = st.sidebar.radio("ניווט", ["📊 עיבוד חודשי", "⚙️ הגדרות"])

# ════════════════════════════════════════════════════════════
#  PAGE 1 — Monthly Processing
# ════════════════════════════════════════════════════════════
if page == "📊 עיבוד חודשי":
    st.title("📊 ממיר הכנסות → חשבשבת")

    uploaded = st.file_uploader(
        "העלי קובץ הכנסות חודשי (.xlsx)",
        type=["xlsx"],
        help="לדוגמה: רחל_מור_הכנסות_3_2026.xlsx",
    )

    if not uploaded:
        st.info("⬆️ העלי קובץ כדי להתחיל")
        st.stop()

    df = pd.read_excel(uploaded, header=None)
    month_label = detect_month_label(uploaded.name, df)

    st.success(f"✅ נטען: **{uploaded.name}** | חודש: **{month_label}**")

    with st.expander("👁️ תצוגה מקדימה"):
        try:
            preview = df.copy()
            preview.columns = df.iloc[0].tolist()
            st.dataframe(preview.iloc[1:].reset_index(drop=True), use_container_width=True)
        except Exception:
            st.dataframe(df, use_container_width=True)

    st.divider()

    if st.button("🚀 עבד והכן קובץ ייבוא", type="primary", use_container_width=True):
        with st.spinner("מעבד..."):
            clients_dict = get_all_clients()
            extra_cols   = get_account_columns()

            invoice_rows, receipt_rows, unmatched, unknown_cols = convert_income_file(
                df, clients_dict, extra_cols
            )

        # ── Handle unknown columns ──────────────────────────
        if unknown_cols:
            st.error("⚠️ עמודות לא מוכרות — הגדירי חשבון לכל אחת")
            for col_name in unknown_cols:
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
        if unmatched:
            st.error(f"⚠️ {len(unmatched)} לקוחות לא נמצאו באינדקס")
            for client_name in unmatched:
                with st.form(f"client_form_{client_name}"):
                    st.markdown(f"**{client_name}**")
                    acct = st.number_input(
                        "מספר חשבון (3000–3999)",
                        min_value=3000, max_value=3999, step=1,
                    )
                    if st.form_submit_button("💾 שמור"):
                        add_client(client_name, int(acct))
                        st.success(f"נשמר! לחצי שוב על 'עבד'")
                        st.rerun()

        # ── All good — show download ─────────────────────────
        if not unmatched and not unknown_cols:
            c1, c2, c3 = st.columns(3)
            c1.metric("שורות חשבוניות", len(invoice_rows))
            c2.metric("שורות קבלות",    len(receipt_rows))
            c3.metric('סה"כ שורות',     len(invoice_rows) + len(receipt_rows))

            excel_data  = create_excel_output(invoice_rows, receipt_rows, month_label)
            output_name = f"Hashavshevet_{month_label.replace('.', '_')}.xlsx"

            st.success("✅ הקובץ מוכן!")
            st.download_button(
                label=f"⬇️ הורד קובץ ייבוא — {month_label}",
                data=excel_data,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )


# ════════════════════════════════════════════════════════════
#  PAGE 2 — Settings
# ════════════════════════════════════════════════════════════
elif page == "⚙️ הגדרות":
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
