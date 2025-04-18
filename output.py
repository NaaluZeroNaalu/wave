import streamlit as st

pages = {

    "REPORTS":[
        st.Page("dashboard.py",title="Dashboard",icon=":material/dashboard:"),
        st.Page("ncr.py",title="NCR Report",icon=":material/dataset:"),
        st.Page("structure_and_finishing_main.py",title="Overall Project Completion - Structure and Finishing",icon=":material/view_timeline:"),
        st.Page("shedule_report.py",title="Flatwise report",icon=":material/apartment:"),
        st.Page("safety.py",title="Safety report",icon=":material/engineering:"),
        st.Page("house.py",title="House report",icon=":material/home:"),
    ]
}


st.navigation(pages).run()
