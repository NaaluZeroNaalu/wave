import streamlit as st
from structure_and_finishing1 import *
from structure_and_finishing2 import *
from structure_and_finishing3 import *
from structure_and_finishing4 import *
from io import BytesIO
import pandas as pd
import urllib.parse
import ibm_boto3
from ibm_botocore.client import Config
import io



def get_cos_files():
    try:
        response = st.session_state.cos_client.list_objects_v2(Bucket="projectreport")
        files = [obj['Key'] for obj in response.get('Contents', []) if obj['Key'].endswith('.xlsx')]
        if not files:
            print("No .json files found in the bucket 'ozonetell'. Please ensure JSON files are uploaded.")
        return files
    except Exception as e:
        print(f"Error fetching COS files: {e}")
        return []
    

st.title("Multiple Excel File Sheet Name Viewer")

# Convert to Excel in memory
def to_excel(df):
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Combined')
        processed_data = output.getvalue()
        return processed_data
    except Exception as e:
        st.error(f"Error generating Excel file: {str(e)}")
        return None

# uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

# if uploaded_files:
#     averagedf = None
#     averagedf2 = None
#     averagedf3 = None
#     veridia = pd.DataFrame(Getprecentage(uploaded_files))
#     for file in uploaded_files:
#         if "Structure Work Tracker EWS LIG P4 all towers" in file.name:
#             st.write("Processing first file")
#             averagedf = CountingProcess(file)
#         elif "Structure Work Tracker Tower G & Tower H" in file.name:
#             st.write("Processing second file")
#             averagedf2 = CountingProcess2(file)
#         elif "Structure Work Tracker Tower 6 & Tower 7" in file.name:
#             st.write("Processing third file")
#             averagedf3 = CountingProcess3(file)

files = get_cos_files()    

averagedf = None
averagedf2 = None
averagedf3 = None
veridia = pd.DataFrame(Getprecentage(files))
st.write(veridia)
for file in files:
    if "Structure Work Tracker EWS LIG P4 all towers" in file:
        response = st.session_state.cos_client.get_object(Bucket="projectreport", Key=file)
        st.write("Processing first file")
        averagedf = CountingProcess(io.BytesIO(response['Body'].read()))
    elif "Structure Work Tracker Tower G & Tower H" in file:
        st.write("Processing second file")
        response = st.session_state.cos_client.get_object(Bucket="projectreport", Key=file)
        averagedf2 = CountingProcess2(io.BytesIO(response['Body'].read()))
    elif "Structure Work Tracker Tower 6 & Tower 7" in file:
        st.write("Processing third file")
        response = st.session_state.cos_client.get_object(Bucket="projectreport", Key=file)
        averagedf3 = CountingProcess3(io.BytesIO(response['Body'].read()))
        

    st.write(veridia)
    if averagedf is not None:
        st.write(averagedf)
    if averagedf2 is not None:
        st.write(averagedf2)
    if averagedf3 is not None:
        st.write(averagedf3)

    # Combine DataFrames
    combined_dfs = [df for df in [veridia, averagedf, averagedf2, averagedf3] if df is not None and not df.empty]
    if combined_dfs:
        combined_df = pd.concat(combined_dfs, ignore_index=True)
        st.write("### 🧾 Combined Data", combined_df)
        st.session_state.structure_and_finishingdf = combined_df
        st.session_state.structure_and_finishing = to_excel(combined_df)
    #     if excel_data:
    #         st.download_button(
    #             label="📥 Download Combined Excel",
    #             data=excel_data,
    #             file_name="Combined_Tracker_Data.xlsx",
    #             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    #         )
    # else:
    #     st.warning("No data to combine.")