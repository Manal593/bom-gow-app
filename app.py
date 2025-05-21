import streamlit as st
import pandas as pd
import io

st.title("GOW Quantity Calculator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Sheet1', header=1)

    # Fill missing values for Remaining C-BOM and GOW calculations
    df["Remaining C-BOM quantity mtr"] = df["C-BOM Quantity\nmtr"] - df["P-BOM\nQuantity\nmtr"]
    df["Remaining C-BOM Cm-M"] = df["C-BOM \nCm-M"] - df["P-BOM\nCm-M"]
    df["GOW Quantity\ncm-mtr"] = df["Approved Quantity\ncm-mtr"] - df["C-BOM \nCm-M"]
    df["GOW Quantity\ncm-mtr"] = df["GOW Quantity\ncm-mtr"].apply(lambda x: x if x > 0 else 0)

    # Filter rows where GOW Quantity > 0
    gow_rows = df[df["GOW Quantity\ncm-mtr"] > 0]
    st.subheader(f"Items with GOW > 0: {len(gow_rows)}/{len(df)}")
    st.dataframe(gow_rows)

    # Convert updated DataFrame to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    # Download button
    st.download_button(
        label="⬇️ Download Updated File",
        data=output,
        file_name="updated_GOW.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

