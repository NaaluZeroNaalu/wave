import streamlit as st
import pandas as pd
import io
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import urllib.parse
import ibm_boto3
from ibm_botocore.client import Config
import io

COS_API_KEY = "VVATXn7oVk51VbC-cCE1DJJEA_UCSAMu1r6ushqG2GQ9"
COS_SERVICE_INSTANCE_ID = "crn:v1:bluemix:public:cloud-object-storage:global:a/fddc2a92db904306b413ed706665c2ff:e99c3906-0103-4257-bcba-e455e7ced9b7:bucket:projectreport"
COS_ENDPOINT = "https://s3.us-south.cloud-object-storage.appdomain.cloud"
COS_BUCKET = "projectreport"



st.session_state.cos_client = ibm_boto3.client(
    's3',
    ibm_api_key_id=COS_API_KEY,
    ibm_service_instance_id=COS_SERVICE_INSTANCE_ID,
    config=Config(signature_version='oauth'),
    endpoint_url=COS_ENDPOINT
)

if 'ncr' not in st.session_state:
    st.session_state.ncr = None

if 'ncrdf' not in st.session_state:
    st.session_state.ncrdf = None

if 'file' not in st.session_state:
    st.session_state.file = None

if 'filedf' not in st.session_state:
    st.session_state.filedf = None

if 'structure_and_finishing' not in st.session_state:
    st.session_state.structure_and_finishing = None

if 'safdf' not in st.session_state:
    st.session_state.safdf = None

if 'shedule' not in st.session_state:
    st.session_state.shedule = None

if 'sheduledf' not in st.session_state:
    st.session_state.shedulef = None

if 'safety' not in st.session_state:
    st.session_state.safety = None

if 'safetydf' not in st.session_state:
    st.session_state.safetydf = None

if 'house' not in st.session_state:
    st.session_state.house = None

if 'housedf' not in st.session_state:
    st.session_state.housedf = None

def create_combined_excel():
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    # Dictionary mapping session state keys to sheet names
    reports = {
        'ncr': 'NCR_Report',
        'structure_and_finishing': 'Structure_and_Finishing',
        'shedule': 'Schedule',
        'safety': 'Safety',
        'house': 'House'
    }
    
    for key, sheet_name in reports.items():
        if st.session_state.get(key) is not None:
            # Convert BytesIO to DataFrame
            df = pd.read_excel(st.session_state[key])
            ws = wb.create_sheet(sheet_name)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

st.divider()
if st.session_state.ncr == None:
    st.write("Run to get output")
else:
    st.write("NCR Report")
    # st.write(st.session_state.ncrdf)
    for section_name, section_data in st.session_state.ncrdf.items():
        st.subheader(section_name)
        
        sites_data = section_data["Sites"]
        df = pd.DataFrame.from_dict(sites_data, orient='index')
        df.index.name = "Site"
        df.reset_index(inplace=True)
        
        st.dataframe(df, use_container_width=True)
    st.download_button(
            label="📥 Download Combined Excel Report",
            data=st.session_state.ncr,
            file_name=f"NCR_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    

st.divider()
if st.session_state.structure_and_finishing == None:
    st.write("Run to get output")
else:
    st.write("structure_and_finishing")
    st.write(st.session_state.structure_and_finishingdf)
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.structure_and_finishing,
            file_name=f"Structure_and_finishing_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )  
   

st.divider()
if st.session_state.shedule == None:
    st.write("Run to get output")
else:
    st.write("Shedule Report")
    st.write(st.session_state.sheduledf)  
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.shedule,
            file_name=f"Shedule_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )  
   
st.divider()
if st.session_state.safety == None:
    st.write("Run to get output")
else:
    st.write("Safety Report")
    # st.write(st.session_state.safetydf)  
    section_name = "Safety"
    section = st.session_state.safetydf[section_name]

    # Convert nested site counts to DataFrame
    df = pd.DataFrame.from_dict(section["Sites"], orient='index')
    df.index.name = "Site"
    df.reset_index(inplace=True)

    # Display section title and table
    st.subheader(section_name)
    st.dataframe(df, use_container_width=True)
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.safety,
            file_name=f"Safety_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )  
    
    

st.divider()
if st.session_state.house == None:
    st.write("Run to get output")
else:
    st.write("House Report")
    # st.write(st.session_state.housedf) 
    df_error = pd.DataFrame([{"Type": "Error", "Message": st.session_state.housedf["error"]}])
    st.dataframe(df_error, use_container_width=True)
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.house,
            file_name=f"House_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ) 
   

# if any(st.session_state.get(key) is not None for key in ['ncr', 'structure_and_finishing', 'shedule', 'safety', 'house']):
#     st.download_button(
#         label="📥 Download All Reports",
#         data=create_combined_excel(),
#         file_name="All_Reports.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#     )

if any(st.session_state.get(key) is not None for key in ['ncr', 'structure_and_finishing', 'shedule', 'safety', 'house']):
    
    # Read the file from disk as bytes
    file_path = "combined_report.xlsx"  # <-- Change this to your actual file path
    try:
        with open(file_path, "rb") as f:
            file_bytes = f.read()

        # Download button for the existing file
        st.download_button(
            label="📥 Download All Reports",
            data=file_bytes,
            file_name="All_Reports.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except FileNotFoundError:
        st.error(f"File not found at path: {file_path}")