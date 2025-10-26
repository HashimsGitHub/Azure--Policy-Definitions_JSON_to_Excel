import streamlit as st
import pandas as pd
import json
from io import BytesIO

st.set_page_config(page_title="Azure Policy Viewer", layout="wide")
st.title("📜 Azure Policy Definitions Viewer")

uploaded_file = st.file_uploader("Upload your Azure Policy JSON file", type=["json"])

if uploaded_file:
    try:
        data = json.load(uploaded_file)

        # Flatten the list of policy definitions
        df = pd.json_normalize(
            data,
            sep='_',
            record_path=None,
            meta=[
                'name',
                'id',
                'type',
                ['properties', 'displayName'],
                ['properties', 'description'],
                ['properties', 'policyType'],
                ['properties', 'mode']
            ],
            errors='ignore'
        )

        st.subheader("📊 Policy Definitions Table")
        st.dataframe(df, use_container_width=True)

        # Convert to Excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='PolicyDefinitions')
                workbook = writer.book
                worksheet = writer.sheets['PolicyDefinitions']
                header_format = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'valign': 'top',
                    'fg_color': '#D7E4BC', 'border': 1
                })
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 30)
            output.seek(0)
            return output

        excel_data = to_excel(df)

        st.download_button(
            label="📥 Download Excel File",
            data=excel_data,
            file_name="AzurePolicyDefinitions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")
