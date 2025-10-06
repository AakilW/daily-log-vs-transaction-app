import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Daily Log vs Transaction Report", layout="wide")
st.title("⚡ Fast Daily Log vs Transaction Comparator")

@st.cache_data(show_spinner=False)
def load_excel(file, usecols):
    return pd.read_excel(file, dtype=str, usecols=usecols, engine="openpyxl")

@st.cache_data(show_spinner=False)
def process_files(daily_df, txn_df):
    daily_df = daily_df.fillna("").astype(str)
    txn_df = txn_df.fillna("").astype(str)

    # Filter required columns only
    daily_subset = daily_df[["Accession", "Test Code"]].copy()
    txn_subset = txn_df[["Inv.#", "Code", "TestName"]].copy()

    # Strip and normalize
    for df in [daily_subset, txn_subset]:
        for c in df.columns:
            df[c] = df[c].str.strip()

    # Build maps using vectorized operations
    daily_map = daily_subset.groupby("Accession")["Test Code"].agg(set).to_dict()
    txn_map = txn_subset.groupby("Inv.#")["Code"].agg(set).to_dict()
    txn_test_map = txn_subset.set_index(["Inv.#", "Code"])["TestName"].to_dict()

    # Vectorized merge to detect matches
    merged = daily_df.merge(
        txn_subset.rename(columns={"Inv.#": "Accession", "Code": "Test Code"}),
        on=["Accession", "Test Code"],
        how="outer",
        indicator=True,
        suffixes=("", "_txn")
    )

    # Fill missing test names
    merged["Test Name"] = merged.apply(
        lambda r: txn_test_map.get((r["Accession"], r["Test Code"]), r.get("Test Name")),
        axis=1
    )

    # Notes for discrepancies
    notes = []
    for acc, codes in daily_map.items():
        if acc not in txn_map:
            notes.append({"Accession": acc, "Action": "Accession not in Transaction Report"})
        else:
            missing_codes = codes - txn_map[acc]
            for c in missing_codes:
                notes.append({"Accession": acc, "Action": f"Removed code {c}"})
    for acc, codes in txn_map.items():
        missing_codes = codes - daily_map.get(acc, set())
        for c in missing_codes:
            notes.append({"Accession": acc, "Action": f"Added missing code {c}"})

    cleaned = merged.drop(columns=["_merge", "TestName_txn"], errors="ignore")

    # Format all date-like columns fast
    date_cols = [c for c in cleaned.columns if any(x in c.lower() for x in ["date", "dob", "collected", "received", "billed"])]
    for c in date_cols:
        cleaned[c] = pd.to_datetime(cleaned[c], errors="coerce").dt.strftime("%m/%d/%Y")

    return cleaned, pd.DataFrame(notes)

# File upload
daily_file = st.file_uploader("Upload Daily Log Report", type=["xlsx"])
txn_file = st.file_uploader("Upload Transaction Report", type=["xlsx"])

if daily_file and txn_file:
    with st.spinner("Processing..."):
        daily_df = load_excel(daily_file, ["Accession", "Test Code"])
        txn_df = load_excel(txn_file, ["Inv.#", "Code", "TestName"])
        cleaned_df, notes_df = process_files(daily_df, txn_df)

    st.success("✅ Cleaned file ready!")

    st.download_button(
        "📥 Download Cleaned Daily Log",
        data=cleaned_df.to_excel(BytesIO(), index=False),
        file_name="Cleaned_Daily_Log.xlsx"
    )

    st.download_button(
        "📥 Download Discrepancy Notes",
        data=notes_df.to_excel(BytesIO(), index=False),
        file_name="Discrepancy_Notes.xlsx"
    )
