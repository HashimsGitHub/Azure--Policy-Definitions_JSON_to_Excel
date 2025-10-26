import streamlit as st
import pandas as pd
import json
from io import BytesIO

st.set_page_config(page_title="Azure Policy Explorer", layout="wide")
st.title("📘 Azure Policy Definitions Explorer")

uploaded_file = st.file_uploader("Upload Azure Policy JSON file", type=["json"])

if uploaded_file:
    try:
        data = json.load(uploaded_file)

        # Flatten each policy definition
        records = []
        for item in data:
            props = item.get("properties", {})
            rule = props.get("policyRule", {})
            parameters = props.get("parameters", {})

            # Flatten parameters into readable string
            param_str = "\n".join([
                f"{key}: {val.get('type', '')} - {val.get('metadata', {}).get('description', '')}"
                for key, val in parameters.items()
            ]) if parameters else ""

            # Safely stringify mixed-type policyRule logic
            rule_if = json.dumps(rule.get("if", {}), indent=2)
            rule_then = json.dumps(rule.get("then", {}), indent=2)

            records.append({
                "Display Name": props.get("displayName"),
                "Description": props.get("description"),
                "Policy Type": props.get("policyType"),
                "Mode": props.get("mode"),
                "Name": item.get("name"),
                "ID": item.get("id"),
                "Type": item.get("type"),
                "Parameters": param_str,
                "Rule - IF": rule_if,
                "Rule - THEN": rule_then
            })

        df = pd.DataFrame(records)

        st.subheader("📊 All Azure Policy Definitions")
        st.dataframe(df, use_container_width=True)

        # Excel export function
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
                    worksheet.set_column(col_num, col_num, 40)
            output.seek(0)
            return output

        excel_data = to_excel(df)

        st.download_button(
            label="📥 Download Full Excel File",
            data=excel_data,
            file_name="AzurePolicyDefinitions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")
