import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Worklist Comparison Checker", layout="wide")
st.title("📑 Worklist Comparison Checker")
st.caption("Upload yerday and today file to automatically compare the file contact number")
st.divider()

st.markdown("""
Compare yesterday's and today's worklists to identify changes in customer contact information.  
**Unique Identifier:** CUST_ID  
""")

COLUMNS = ["CUST_ID", "CUST_NAME", "OFC", "HOME", "MOBILE_NO"]


def format_mobile(mobile):
    if pd.isna(mobile):
        return ''
    mobile = str(mobile).strip().replace(".0", "")
    mobile = mobile.lstrip("0")
    return "0" + mobile if mobile else ""


def load_file(file):
    try:
        df = pd.read_excel(file, dtype={"MOBILE_NO": str})
        df.columns = df.columns.str.strip().str.upper()

        missing = [c for c in COLUMNS if c not in df.columns]
        if missing:
            st.error(f"Missing columns: {', '.join(missing)}")
            return None

        df = df[COLUMNS].copy()
        df["MOBILE_NO"] = df["MOBILE_NO"].apply(format_mobile)
        df["CUST_ID"] = df["CUST_ID"].astype(str).str.strip()
        df["CUST_ID"] = df["CUST_ID"].apply(lambda x: x if x.startswith("0") else "0" + x)


        return df

    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None


def compare_worklists(yesterday, today):

    merged = yesterday.merge(
        today,
        on="CUST_ID",
        how="outer",
        suffixes=("_OLD", "_NEW"),
        indicator=True
    )

    changes = []
    changed_columns_global = set()

    for _, row in merged.iterrows():

        if row["_merge"] == "left_only":
            changes.append({
                "CUST_ID": row["CUST_ID"],
                "CHANGE_TYPE": "DELETED"
            })

        elif row["_merge"] == "right_only":
            changes.append({
                "CUST_ID": row["CUST_ID"],
                "CHANGE_TYPE": "NEW"
            })

        else:
            changed_cols = []

            for col in ["CUST_NAME", "OFC", "HOME", "MOBILE_NO"]:
                old_val = row[f"{col}_OLD"]
                new_val = row[f"{col}_NEW"]

                # Normalize NaN to empty string
                old_val = "" if pd.isna(old_val) else str(old_val).strip()
                new_val = "" if pd.isna(new_val) else str(new_val).strip()

                # 🔴 SPECIAL RULE:
                # If OLD has value but NEW is blank → IGNORE
                if old_val != "" and new_val == "":
                    continue

                if old_val != new_val:
                    changed_cols.append(col)
                    changed_columns_global.add(col)

            if changed_cols:
                changes.append({
                    "CUST_ID": row["CUST_ID"],
                    "CHANGE_TYPE": "UPDATED",
                    "COLUMNS_CHANGED": ", ".join(changed_cols),
                    "COLUMNS_CHANGED_LIST": changed_cols
                })

    changes_df = pd.DataFrame(changes)

    summary = {
        "total_records_yesterday": len(yesterday),
        "total_records_today": len(today),
        "new_customers": (changes_df["CHANGE_TYPE"] == "NEW").sum() if not changes_df.empty else 0,
        "updated_customers": (changes_df["CHANGE_TYPE"] == "UPDATED").sum() if not changes_df.empty else 0,
        "deleted_customers": (changes_df["CHANGE_TYPE"] == "DELETED").sum() if not changes_df.empty else 0,
        "total_changes": len(changes_df)
    }

    return changes_df, summary, changed_columns_global


def generate_updated(today, changes_df, changed_columns):

    updated = changes_df[changes_df["CHANGE_TYPE"] == "UPDATED"]

    if updated.empty:
        return pd.DataFrame(columns=["CHANGE_TYPE", "CUST_ID"])

    today_filtered = today[today["CUST_ID"].isin(updated["CUST_ID"])].copy()

    change_map = {
        row["CUST_ID"]: "UPDATED_" + "_".join(row["COLUMNS_CHANGED_LIST"])
        for _, row in updated.iterrows()
    }

    today_filtered["CHANGE_TYPE"] = today_filtered["CUST_ID"].map(change_map)

    columns = ["CHANGE_TYPE", "CUST_ID"] + sorted(changed_columns)

    return today_filtered[columns]


col1, col2 = st.columns(2)

with col1:
    yesterday_file = st.file_uploader("Upload Yesterday File", type=["xlsx", "xls"])

with col2:
    today_file = st.file_uploader("Upload Today File", type=["xlsx", "xls"])

if yesterday_file and today_file:

    yesterday_df = load_file(yesterday_file)
    today_df = load_file(today_file)

    if yesterday_df is not None and today_df is not None:

        changes_df, summary, changed_columns = compare_worklists(yesterday_df, today_df)

        # Summary
        st.header("📊 Summary")
        c1, c2, c3, c4, c5 = st.columns(5)

        c1.metric("Yesterday", summary["total_records_yesterday"])
        c2.metric("Today", summary["total_records_today"])
        c3.metric("New", summary["new_customers"])
        c4.metric("Updated", summary["updated_customers"])
        c5.metric("Deleted", summary["deleted_customers"])

        # Detailed Changes
        if not changes_df.empty:
            st.header("🔄 Detailed Changes")
            st.dataframe(changes_df, use_container_width=True)

        else:
            st.success("No changes detected.")

        st.header("📝 Updated Records Only")
        updated_df = generate_updated(today_df, changes_df, changed_columns)

        if not updated_df.empty:

            st.dataframe(updated_df, use_container_width=True)

            # Excel Export
            buffer = BytesIO()

            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                updated_df.to_excel(writer, index=False, sheet_name="Updated")

                ws = writer.sheets["Updated"]

                # Force MOBILE_NO as text
                if "MOBILE_NO" in updated_df.columns:
                    col_idx = updated_df.columns.get_loc("MOBILE_NO") + 1
                    for row in ws.iter_rows(min_row=2,
                                            min_col=col_idx,
                                            max_col=col_idx):
                        for cell in row:
                            cell.number_format = "@"

            buffer.seek(0)

            st.download_button(
                "📥 Download Updated (Excel)",
                buffer,
                f"updated_{datetime.now().strftime('%m%d%Y')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.info("No updated records found.")
