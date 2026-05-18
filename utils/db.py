import streamlit as st
from supabase import create_client, Client
import pandas as pd


@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["supabase"]["SUPABASE_URL"]
    key = st.secrets["supabase"]["SUPABASE_KEY"]
    return create_client(url, key)


def get_all_clients() -> dict:
    sb = get_supabase()
    result = sb.table("clients").select("name, account_number").order("account_number").execute()
    return {row["name"]: row["account_number"] for row in result.data}


def add_client(name: str, account_number: int):
    sb = get_supabase()
    sb.table("clients").upsert({"name": name, "account_number": account_number}).execute()


def delete_client(name: str):
    sb = get_supabase()
    sb.table("clients").delete().eq("name", name).execute()


def get_next_account_number() -> int:
    sb = get_supabase()
    result = sb.table("clients").select("account_number").order("account_number", desc=True).limit(1).execute()
    if result.data:
        return result.data[0]["account_number"] + 1
    return 3001


def import_clients_from_df(df: pd.DataFrame) -> int:
    """Import clients from מאזן בוחן Excel file."""
    sb = get_supabase()
    clients = []

    for _, row in df.iterrows():
        account = row.iloc[2] if len(row) > 2 else None
        name    = row.iloc[3] if len(row) > 3 else None

        if pd.isna(account) or pd.isna(name):
            continue

        account_str = str(account).replace(".0", "").strip()
        # Client accounts: 3000–3999
        if not (account_str.isdigit() and len(account_str) == 4 and account_str.startswith("3")):
            continue

        name_str = str(name).strip()
        if not name_str or name_str.isdigit() or len(name_str) < 2:
            continue

        try:
            clients.append({"name": name_str, "account_number": int(float(account))})
        except Exception:
            pass

    if clients:
        sb.table("clients").upsert(clients).execute()

    return len(clients)


def get_account_columns() -> dict:
    sb = get_supabase()
    result = sb.table("account_columns").select("*").execute()
    return {
        row["column_name"]: {
            "account":    row["account_number"],
            "vat_exempt": row["is_vat_exempt"],
        }
        for row in result.data
    }


def add_account_column(column_name: str, account_number: int, is_vat_exempt: bool):
    sb = get_supabase()
    sb.table("account_columns").upsert({
        "column_name":   column_name,
        "account_number": account_number,
        "is_vat_exempt":  is_vat_exempt,
    }).execute()
