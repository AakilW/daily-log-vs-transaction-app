import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Daily Log vs Transaction Report", layout="wide")
st.title("Daily Log vs Transaction Report Comparator")

daily_file = st.file_uploader("Upload Daily Log Report", type=["xlsx", "xls"])
txn_file = st.file_uploader("Upload Transaction Report", type=["xlsx", "xls"])

if daily_file and txn_file:
    # Read file bytes once
    daily_bytes = BytesIO(daily_file.read())
    txn_bytes = BytesIO(txn_file.read())

    # Read only required columns to reduce memory
    daily_df = pd.read_excel(daily_bytes, dtype=str, engine="openpyxl", usecols=lambda x: True)
    txn_df = pd.read_excel(txn_bytes, dtype=str, engine="openpyxl", usecols=lambda x: True)

    # Normalize column names
    daily_df.columns = daily_df.columns.str.strip()
    txn_df.columns = txn_df.columns.str.strip()

    # Identify likely column names
    acc_col_daily = [c for c in daily_df.columns if "acc" in c.lower()][0]
    code_col_daily = [c for c in daily_df.columns if "code" in c.lower()][0]
    acc_col_txn = [c for c in txn_df.columns if "inv" in c.lower() or "acc" in c.lower()][0]
    code_col_txn = [c for c in txn_df.columns if "code" in c.lower()][0]
    name_col_txn = next((c for c in txn_df.columns if "name" in c.lower()), code_col_txn)

    # Convert to string and strip
    for c in [acc_col_daily, code_col_daily]:
        daily_df[c] = daily_df[c].astype(str).str.strip()
    for c in [acc_col_txn, code_col_txn]:
        txn_df[c] = txn_df[c].astype(str).str.strip()

    # Build maps
    daily_map = daily_df.groupby(acc_col_daily)[code_col_daily].apply(set).to_dict()
    txn_map = txn_df.groupby(acc_col_txn)[code_col_txn].apply(set).to_dict()
    txn_name_map = txn_df.set_index([acc_col_txn, code_col_txn])[name_col_txn].to_dict()

    # Build cleaned list
    cleaned_rows, notes = [], []
    for _, r in daily_df.iterrows():
        acc = r[acc_col_daily]
        code = r[code_col_daily]
        row = r.to_dict()
        if acc in txn_map:
            if code in txn_map[acc]:
                row["Test Name"] = txn_name_map.get((acc, code), "")
                cleaned_rows.append(row)
            else:
                notes.append({"Accession": acc, "Action": f"Removed code {code}"})
        else:
            notes.append({"Accession": acc, "Action": "Accession not in Transaction (kept)"})
            cleaned_rows.append(row)

    # Add missing from transaction
    for acc, codes in txn_map.items():
        missing = codes - daily_map.get(acc, set())
        for c in missing:
            template = {k: "" for k in daily_df.columns}
            template[acc_col_daily] = acc
            template[code_col_daily] = c
            template["Test Name"] = txn_name_map.get((acc, c), "")
            cleaned_rows.append(template)
            notes.append({"Accession": acc, "Action": f"Added missing code {c}"})

    cleaned_df = pd.DataFrame(cleaned_rows)
    notes_df = pd.DataFrame(notes)

    tab1, tab2, tab3 = st.tabs(["Cleaned Log", "Notes", "Reconciliation"])

    with tab1:
        st.write("Cleaned Log Preview", cleaned_df.head(20))
        out = BytesIO()
        cleaned_df.to_excel(out, index=False, engine="openpyxl")
        st.download_button("Download Cleaned File", out.getvalue(), "Cleaned_Daily_Log.xlsx")

    with tab2:
        st.write(notes_df if not notes_df.empty else "No discrepancies found")

    with tab3:
        daily_acc = set(daily_map.keys())
        txn_acc = set(txn_map.keys())
        recon = pd.DataFrame({
            "Metric": ["Daily Count", "Txn Count", "Missing in Txn", "Missing in Daily"],
            "Count": [len(daily_acc), len(txn_acc), len(daily_acc - txn_acc), len(txn_acc - daily_acc)]
        })
        st.table(recon)
