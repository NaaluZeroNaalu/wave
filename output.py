import streamlit as st

pages = {

    "REPORTS":[
        st.Page("dashboard.py",title="Dashboard",icon=":material/dashboard:"),
        st.Page("ncr.py",title="NCR Report",icon=":material/dataset:"),
<<<<<<< HEAD
        st.Page("structure_and_finishing_main.py",title="Overall Project Completion - Structure and Finishing",icon=":material/view_timeline:"),
        st.Page("shedule_report.py",title="Flatwise report",icon=":material/apartment:"),
        st.Page("safety.py",title="Safety report",icon=":material/engineering:"),
        st.Page("house.py",title="House report",icon=":material/home:"),
=======
        st.Page("structure_and_finishing_main.py",title="Structure and Finishing",icon=":material/view_timeline:"),
        st.Page("shedule_report.py",title="Flatwise shedule report",icon=":material/apartment:"),
        st.Page("Safety.py",title="Safety report",icon=":material/engineering:"),
        st.Page("House.py",title="House report",icon=":material/home:"),
        st.Page("testcos.py",title="Test",icon=":material/home:")
>>>>>>> f74a99153372cb94db7f6c9b7edef174a21ee8b8
    ]
}


st.navigation(pages).run()
