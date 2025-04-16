import streamlit as st
import pandas as pd
import io
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows



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

if st.session_state.ncr == None:
    st.write("Run to get output")
else:
    st.write("NCR Report")
    st.dataframe(st.session_state.ncrdf)
    st.download_button(
            label="📥 Download Combined Excel Report",
            data=st.session_state.ncr,
            file_name=f"NCR_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    

if st.session_state.structure_and_finishing == None:
    st.write("Run to get output")
else:
    st.write("structure_and_finishing")
    # st.write(st.session_state.structure_and_finishing)
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.structure_and_finishing,
            file_name=f"Structure_and_finishing_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )  
   

if st.session_state.shedule == None:
    st.write("Run to get output")
else:
    st.write("Shedule Report")
    # st.write(st.session_state.shedule)  
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.shedule,
            file_name=f"Shedule_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )  
   
if st.session_state.safety == None:
    st.write("Run to get output")
else:
    st.write("Safety Report")
    st.write(st.session_state.safety)  
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.safety,
            file_name=f"Safety_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )  
    
    

if st.session_state.house == None:
    st.write("Run to get output")
else:
    st.write("House Report")
    # st.write(st.session_state.house) 
    st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.house,
            file_name=f"House_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ) 
   

if any(st.session_state.get(key) is not None for key in ['ncr', 'structure_and_finishing', 'shedule', 'safety', 'house']):
    st.download_button(
        label="📥 Download All Reports",
        data=create_combined_excel(),
        file_name="All_Reports.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )