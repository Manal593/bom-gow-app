
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="BOM GOW Calculator", layout="wide")
st.title("📊 BOM GOW Calculator")
st.markdown("Upload your BOM Excel file. The app will calculate Remaining and GOW values, display results interactively, and allow file download.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Sheet1', header=1)

    df.columns = [
        "Project", "Item Code", "Item Description", "Size", "Material",
        "P-BOM Quantity mtr", "P-BOM Cm-M", "C-BOM Description", "C-BOM Material",
        "C-BOM Quantity mtr", "C-BOM Cm-M", "PO Quantity mtr", "Approved Quantity mtr",
        "Approved Quantity cm-mtr", "Remaining C-BOM quantity mtr", "Remaining C-BOM Cm-M",
        "GOW Quantity cm-mtr", "Remarks for GOW"
    ]

    # Drop empty rows
    df = df.dropna(subset=["Size", "Material"])

    # Convert numerics
    numeric_cols = [
        "P-BOM Quantity mtr", "P-BOM Cm-M", "C-BOM Quantity mtr", "C-BOM Cm-M",
        "Approved Quantity mtr", "Approved Quantity cm-mtr"
    ]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce')

    # Recalculate columns
    df["Remaining C-BOM quantity mtr"] = df["C-BOM Quantity mtr"] - df["P-BOM Quantity mtr"]
    df["Remaining C-BOM Cm-M"] = df["C-BOM Cm-M"] - df["P-BOM Cm-M"]
    df["GOW Quantity cm-mtr"] = (df["P-BOM Cm-M"] - df["C-BOM Cm-M"]).clip(lower=0)

    st.subheader("📋 Calculated BOM Table")

    # Filters
    col1, col2 = st.columns(2)
    with col1:
        material_filter = st.multiselect("Filter by Material", options=df["Material"].unique())
    with col2:
        size_filter = st.multiselect("Filter by Size", options=sorted(df["Size"].unique()))

    df_filtered = df.copy()
    if material_filter:
        df_filtered = df_filtered[df_filtered["Material"].isin(material_filter)]
    if size_filter:
        df_filtered = df_filtered[df_filtered["Size"].isin(size_filter)]

    st.dataframe(df_filtered.style
                 .applymap(lambda x: "background-color: #ffcccc" if isinstance(x, (int, float)) and x > 0 else "")
                 , use_container_width=True)

    # Summary stats
    st.subheader("📈 Summary")
    total_gow = df["GOW Quantity cm-mtr"].sum()
    num_items_with_gow = (df["GOW Quantity cm-mtr"] > 0).sum()
    total_rows = len(df)

    st.metric("Total GOW (cm-mtr)", f"{total_gow:.0f}")
    st.metric("Items with GOW > 0", f"{num_items_with_gow}/{total_rows}")

    # Download button
    st.subheader("⬇️ Download Updated File")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Updated BOM')
        writer.save()
    st.download_button(
        label="Download as Excel",
        data=output.getvalue(),
        file_name="Updated_BOM_with_GOW.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Done calculating! Scroll to explore data or apply filters.")
