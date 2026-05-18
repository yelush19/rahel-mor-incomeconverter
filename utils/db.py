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


def get_clients_full() -> list:
    """Returns list of {account_number, name, id_number}"""
    sb = get_supabase()
    result = sb.table("clients").select("account_number, name, id_number").order("account_number").execute()
    return result.data


def add_client(name: str, account_number: int, id_number: str = ""):
    sb = get_supabase()
    sb.table("clients").delete().eq("account_number", account_number).execute()
    sb.table("clients").insert(
        {"name": name, "account_number": account_number, "id_number": id_number}
    ).execute()


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
    sb = get_supabase()

    # מצא שורת כותרת
    header_row = None
    for i in range(min(10, len(df))):
        row_vals = [str(v) for v in df.iloc[i].tolist()]
        if any('חשבון' in v for v in row_vals):
            header_row = i
            break

    if header_row is not None:
        headers = [str(v) for v in df.iloc[header_row].tolist()]
        def find_col(keywords):
            for kw in keywords:
                for j, h in enumerate(headers):
                    if kw in h:
                        return j
            return None
        col_account = find_col(['חשבון'])
        col_name    = find_col(['שם חשבון', 'שם'])
        col_id      = find_col(['מס.ע.מ', 'ע.מ', 'ת.ז'])
        data_start  = header_row + 1
    else:
        col_account, col_name, col_id = 2, 3, None
        data_start = 0

    clients = []
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        try:
            account = row.iloc[col_account] if col_account is not None else None
            name    = row.iloc[col_name]    if col_name    is not None else None
            id_num  = row.iloc[col_id]      if col_id      is not None else None
        except Exception:
            continue

        if pd.isna(account) or pd.isna(name):
            continue

        account_str = str(account).replace(".0", "").strip()
        if not (account_str.isdigit() and len(account_str) == 4 and account_str.startswith("3")):
            continue

        name_str = str(name).strip()
        if not name_str or name_str.isdigit() or len(name_str) < 2:
            continue

        id_str = ""
        if id_num is not None and not pd.isna(id_num):
            id_str = str(id_num).replace(".0", "").strip()
            if id_str in ("0", "nan", "None"):
                id_str = ""

        try:
            clients.append({
                "name":           name_str,
                "account_number": int(float(account)),
                "id_number":      id_str,
            })
        except Exception:
            pass

    if clients:
        sb.table("clients").delete().neq("account_number", 0).execute()
        sb.table("clients").insert(clients).execute()

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
    sb.table("account_columns").delete().eq("column_name", column_name).execute()
    sb.table("account_columns").insert({
        "column_name":    column_name,
        "account_number": account_number,
        "is_vat_exempt":  is_vat_exempt,
    }).execute()
