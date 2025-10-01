import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Daily Log vs Transaction Report", layout="wide")

st.title("🔍 Daily Log vs Transaction Report Comparator")

# Upload files
daily_file = st.file_uploader("Upload Daily Log Report", type=["xlsx", "xls"])
txn_file = st.file_uploader("Upload Transaction Report", type=["xlsx", "xls"])

if daily_file and txn_file:
    # Read files
    daily_df = pd.read_excel(daily_file, dtype=str)
    txn_df = pd.read_excel(txn_file, dtype=str)

    tab1, tab2, tab3 = st.tabs(["📑 Cleaned Daily Log", "📝 Discrepancy Notes", "📊 Reconciliation"])

    # ========== TAB 1: CLEANED DAILY LOG ==========
    with tab1:
        st.subheader("Preview of Uploaded Reports")
        st.write("**Daily Log Report**", daily_df.head())
        st.write("**Transaction Report**", txn_df.head())

        # Extract required columns
        daily_subset = daily_df[["Accession", "Test Code"]].dropna()
        txn_subset = txn_df[["Inv.#", "Code", "TestName"]].dropna()

        # Normalize values
        daily_subset["Test Code"] = daily_subset["Test Code"].astype(str).str.strip()
        txn_subset["Code"] = txn_subset["Code"].astype(str).str.strip()
        daily_subset["Accession"] = daily_subset["Accession"].astype(str).str.strip()
        txn_subset["Inv.#"] = txn_subset["Inv.#"].astype(str).str.strip()

        # Build maps
        daily_map = daily_subset.groupby("Accession")["Test Code"].apply(set).to_dict()
        txn_map = txn_subset.groupby("Inv.#")["Code"].apply(set).to_dict()
        txn_testname_map = txn_subset.set_index(["Inv.#", "Code"])["TestName"].to_dict()

        notes = []
        cleaned_rows = []

        # Iterate through Daily Log
        for idx, row in daily_df.iterrows():
            accession = str(row["Accession"]).strip()
            code = str(row.get("Test Code", "")).strip()

            row_dict = row.to_dict()

            if accession in txn_map:
                if code in txn_map[accession]:
                    # ✅ Code exists in both
                    # Add Test Name from transaction
                    row_dict["Test Name"] = txn_testname_map.get((accession, code), row_dict.get("Test Name", None))
                    cleaned_rows.append(row_dict)
                else:
                    # ❌ Code not in txn
                    notes.append({"Accession": accession, "Action": f"Removed code {code}"})
            else:
                # Accession missing in txn → keep
                notes.append({"Accession": accession, "Action": "Accession not in Transaction Report (kept)"})
                cleaned_rows.append(row_dict)

        # Add missing codes from Transaction Report
        for accession, txn_codes in txn_map.items():
            daily_codes = daily_map.get(accession, set())
            missing_codes = txn_codes - daily_codes
            if missing_codes:
                for code in missing_codes:
                    # Take full row template from first row of daily_df with same accession (if exists)
                    if accession in daily_df["Accession"].astype(str).values:
                        template_row = daily_df[daily_df["Accession"].astype(str) == accession].iloc[0].to_dict()
                    else:
                        template_row = {col: None for col in daily_df.columns}
                        template_row["Accession"] = accession

                    template_row["Test Code"] = code
                    template_row["Test Name"] = txn_testname_map.get((accession, code), None)

                    cleaned_rows.append(dict(template_row))
                    notes.append({"Accession": accession, "Action": f"Added missing code {code}"})

    # Final cleaned DataFrame
    cleaned_daily_df = pd.DataFrame(cleaned_rows)
    
    # ✅ Format all date columns as MM/DD/YYYY
    for col in cleaned_daily_df.columns:
        if "date" in col.lower():
            cleaned_daily_df[col] = pd.to_datetime(
                cleaned_daily_df[col], errors="coerce"
            ).dt.strftime("%m/%d/%Y")
    
    st.subheader("📑 Cleaned Daily Log Report")
    st.dataframe(cleaned_daily_df.head())
    
    # Download option
    output = BytesIO()
    cleaned_daily_df.to_excel(output, index=False)
    st.download_button(
        label="📥 Download Cleaned Daily Log Report",
        data=output.getvalue(),
        file_name="Cleaned_Daily_Log_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # ========== TAB 2: DISCREPANCY NOTES ==========
    with tab2:
        st.subheader("📝 Discrepancy Notes")
        notes_df = pd.DataFrame(notes)
        if not notes_df.empty:
            st.dataframe(notes_df)
        else:
            st.success("No discrepancies found ✅")

    # ========== TAB 3: RECONCILIATION ==========
    with tab3:
        st.subheader("📊 Reconciliation Breakdown")

        # --- Accession Level ---
        st.markdown("### 🔹 Accession Level Reconciliation")
        daily_accessions = set(daily_df["Accession"].astype(str))
        txn_accessions = set(txn_df["Inv.#"].astype(str))

        missing_in_txn = daily_accessions - txn_accessions
        missing_in_daily = txn_accessions - daily_accessions

        accession_recon = pd.DataFrame({
            "Metric": ["Total in Daily Log", "Total in Transaction", "Missing in Transaction", "Missing in Daily Log"],
            "Count": [len(daily_accessions), len(txn_accessions), len(missing_in_txn), len(missing_in_daily)]
        })
        st.table(accession_recon)

        # --- CPT Level ---
        st.markdown("### 🔹 CPT Level Reconciliation (per Accession)")
        daily_cpt = daily_subset.groupby("Accession")["Test Code"].apply(set)
        txn_cpt = txn_subset.groupby("Inv.#")["Code"].apply(set)

        recon_rows = []
        for acc in sorted(daily_accessions | txn_accessions):
            d_codes = daily_cpt.get(acc, set())
            t_codes = txn_cpt.get(acc, set())
            missing_in_txn_codes = d_codes - t_codes
            missing_in_daily_codes = t_codes - d_codes
            recon_rows.append({
                "Accession": acc,
                "Daily Codes": ", ".join(d_codes) if d_codes else None,
                "Txn Codes": ", ".join(t_codes) if t_codes else None,
                "Missing in Transaction": ", ".join(missing_in_txn_codes) if missing_in_txn_codes else None,
                "Missing in Daily Log": ", ".join(missing_in_daily_codes) if missing_in_daily_codes else None
            })

        cpt_recon_df = pd.DataFrame(recon_rows)
        st.dataframe(cpt_recon_df)

