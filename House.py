import streamlit as st
import requests
import json
import urllib.parse
import urllib3
import certifi
import pandas as pd  
from bs4 import BeautifulSoup
from datetime import datetime
import re
import logging
import os
from dotenv import load_dotenv
import io
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# WatsonX configuration
WATSONX_API_URL = os.getenv("WATSONX_API_URL")
MODEL_ID = os.getenv("MODEL_ID")
PROJECT_ID = os.getenv("PROJECT_ID")
API_KEY = os.getenv("API_KEY")

# Check environment variables
if not all([API_KEY, WATSONX_API_URL, MODEL_ID, PROJECT_ID]):
    st.error("❌ Required environment variables (API_KEY, WATSONX_API_URL, MODEL_ID, PROJECT_ID) missing!")
    logger.error("Missing one or more required environment variables")
    st.stop()

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# API Endpoints
LOGIN_URL = "https://dms.asite.com/apilogin/"
SEARCH_URL = "https://adoddleak.asite.com/commonapi/formsearchapi/search"
IAM_TOKEN_URL = "https://iam.cloud.ibm.com/identity/token"

# Function to generate access token with retry
def get_access_token(API_KEY):
    session = requests.Session()
    retry_strategy = Retry(
        total=3,
        backoff_factor=5,
        status_forcelist=[500, 502, 503, 504],
        allowed_methods=["POST"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)

    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
    data = {"grant_type": "urn:ibm:params:oauth:grant-type:apikey", "apikey": API_KEY}
    try:
        response = session.post(IAM_TOKEN_URL, headers=headers, data=data, verify=certifi.where(), timeout=120)
        if response.status_code == 200:
            token_info = response.json()
            logger.info("Access token generated successfully")
            return token_info['access_token']
        else:
            logger.error(f"Failed to get access token: {response.status_code} - {response.text}")
            st.error(f"❌ Failed to get access token: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        logger.error(f"Exception getting access token: {str(e)}")
        st.error(f"❌ Error getting access token: {str(e)}")
        return None

# Login Function
def login_to_asite(email, password):
    headers = {"Accept": "application/json", "Content-Type": "application/x-www-form-urlencoded"}
    payload = {"emailId": email, "password": password}
    response = requests.post(LOGIN_URL, headers=headers, data=payload, verify=certifi.where(), timeout=50)
    if response.status_code == 200:
        try:
            session_id = response.json().get("UserProfile", {}).get("Sessionid")
            logger.info(f"Login successful, Session ID: {session_id}")
            return session_id
        except json.JSONDecodeError:
            logger.error("JSONDecodeError during login")
            st.error("❌ Failed to parse login response")
            return None
    logger.error(f"Login failed: {response.status_code}")
    st.error(f"❌ Login failed: {response.status_code}")
    return None

# Fetch Data Function
def fetch_project_data(session_id, project_name, form_name, record_limit=1000):
    headers = {"Accept": "application/json", "Content-Type": "application/x-www-form-urlencoded", "Cookie": f"ASessionID={session_id}"}
    all_data = []
    start_record = 1
    total_records = None

    with st.spinner("Fetching data from Asite..."):
        while True:
            search_criteria = {"criteria": [{"field": "ProjectName", "operator": 1, "values": [project_name]}, {"field": "FormName", "operator": 1, "values": [form_name]}], "recordStart": start_record, "recordLimit": record_limit}
            search_criteria_str = json.dumps(search_criteria)
            encoded_payload = f"searchCriteria={urllib.parse.quote(search_criteria_str)}"
            response = requests.post(SEARCH_URL, headers=headers, data=encoded_payload, verify=certifi.where(), timeout=50)

            try:
                response_json = response.json()
                if total_records is None:
                    total_records = response_json.get("responseHeader", {}).get("results-total", 0)
                all_data.extend(response_json.get("FormList", {}).get("Form", []))
                st.info(f"🔄 Fetched {len(all_data)} / {total_records} records")
                if start_record + record_limit - 1 >= total_records:
                    break
                start_record += record_limit
            except Exception as e:
                logger.error(f"Error fetching data: {str(e)}")
                st.error(f"❌ Error fetching data: {str(e)}")
                break

    return {"responseHeader": {"results": len(all_data), "total_results": total_records}}, all_data, encoded_payload

# Process JSON Data
def process_json_data(json_data):
    data = []
    for item in json_data:
        form_details = item.get('FormDetails', {})
        created_date = form_details.get('FormCreationDate', None)
        expected_close_date = form_details.get('UpdateDate', None)
        form_status = form_details.get('FormStatus', None)
        
        discipline = None
        description = None
        custom_fields = form_details.get('CustomFields', {}).get('CustomField', [])
        for field in custom_fields:
            if field.get('FieldName') == 'CFID_DD_DISC':
                discipline = field.get('FieldValue', None)
            elif field.get('FieldName') == 'CFID_RTA_DES':
                description = BeautifulSoup(field.get('FieldValue', None) or '', "html.parser").get_text()

        days_diff = None
        if created_date and expected_close_date:
            try:
                created_date_obj = datetime.strptime(created_date.split('#')[0], "%d-%b-%Y")
                expected_close_date_obj = datetime.strptime(expected_close_date.split('#')[0], "%d-%b-%Y")
                days_diff = (expected_close_date_obj - created_date_obj).days
            except Exception as e:
                logger.error(f"Error calculating days difference: {str(e)}")
                days_diff = None

        data.append([days_diff, created_date, expected_close_date, description, form_status, discipline])

    df = pd.DataFrame(data, columns=['Days', 'Created Date (WET)', 'Expected Close Date (WET)', 'Description', 'Status', 'Discipline'])
    df['Created Date (WET)'] = pd.to_datetime(df['Created Date (WET)'].str.split('#').str[0], format="%d-%b-%Y", errors='coerce')
    df['Expected Close Date (WET)'] = pd.to_datetime(df['Expected Close Date (WET)'].str.split('#').str[0], format="%d-%b-%Y", errors='coerce')
    logger.debug(f"DataFrame columns after processing: {df.columns.tolist()}")  # Debug column names
    if df.empty:
        logger.warning("DataFrame is empty after processing")
        st.warning("⚠️ No data processed. Check the API response.")
    return df

# Generate NCR Report
def generate_ncr_report(df, start_date=None, end_date=None):
    with st.spinner("Generating Housekeeping NCR Report with WatsonX..."):
        # Filter for HES discipline and Days > 21
        filtered_df = df[
            (df['Discipline'] == 'HES') &
            (df['Days'] > 21)
        ].copy()

        if filtered_df.empty:
            return {"error": "No HES records found with duration > 21 days"}, ""

        filtered_df.loc[:, 'Created Date (WET)'] = filtered_df['Created Date (WET)'].astype(str)
        filtered_df.loc[:, 'Expected Close Date (WET)'] = filtered_df['Expected Close Date (WET)'].astype(str)

        processed_data = filtered_df.to_dict(orient="records")
        
        cleaned_data = []
        for record in processed_data:
            cleaned_record = {
                "Description": str(record.get("Description", "")),
                "Created Date (WET)": str(record.get("Created Date (WET)", "")),
                "Expected Close Date (WET)": str(record.get("Expected Close Date (WET)", "")),
                "Status": str(record.get("Status", "")),
                "Days": record.get("Days", 0),
                "Tower": "External Development"
            }

            description = cleaned_record["Description"]
            if any(phrase in description.lower() for phrase in ["veridia clubhouse", "veridia-clubhouse", "veridia club"]):
                cleaned_record["Tower"] = "Veridia-Club"
                logger.debug(f"Matched 'Veridia Clubhouse' variant in description: {description}, setting Tower to Veridia-Club")
            else:
                tower_match = re.search(r"(Tower|T)\s*-?\s*(\d+|2021|28)", description, re.IGNORECASE)
                cleaned_record["Tower"] = f"Veridia- Tower-{tower_match.group(2).zfill(2)}" if tower_match else None
                logger.debug(f"Tower match result: {tower_match}, Tower set to {cleaned_record['Tower']}")

            cleaned_data.append(cleaned_record)
            logger.debug(f"Processed record: {json.dumps(cleaned_record, indent=2)}")

        cleaned_data = [record for record in cleaned_data if record["Tower"] is not None]

        st.write(f"Debug - Total records to process: {len(cleaned_data)}")
        logger.debug(f"Processed data: {json.dumps(cleaned_data, indent=2)}")

        if not cleaned_data:
            return {"Housekeeping": {"Sites": {}, "Grand_Total": 0}}, ""

        access_token = get_access_token(API_KEY)
        if not access_token:
            return {"error": "Failed to obtain access token"}, ""

        result = {"Housekeeping": {"Sites": {}, "Grand_Total": 0}}
        
        chunk_size = 200
        for i in range(0, len(cleaned_data), chunk_size):
            chunk = cleaned_data[i:i + chunk_size]
            st.write(f"Processing chunk {i // chunk_size + 1}: Records {i} to {min(i + chunk_size, len(cleaned_data))}")
            logger.debug(f"Chunk data sent to WatsonX: {json.dumps(chunk, indent=2)}")

            prompt = (
                "IMPORTANT: YOU MUST RETURN ONLY A SINGLE VALID JSON OBJECT. DO NOT INCLUDE ANY TEXT, EXPLANATION, OR MULTIPLE RESPONSES.\n\n"
                "Task: For Housekeeping NCRs, count ALL records in the provided data by site ('Tower' field) where 'Days' > 21 and 'Discipline' is 'HES'. Use the 'Tower' values exactly as they appear in the data (e.g., 'Veridia-Club', 'Veridia- Tower 02', 'External Development') and assign each count to the 'HES' key.\n\n"
                "REQUIRED OUTPUT FORMAT (exactly this structure with actual counts):\n"
                "{\n"
                '  "Housekeeping": {\n'
                '    "Sites": {\n'
                '      "Site_Name1": {\n'
                '        "HES": number\n'
                '      },\n'
                '      "Site_Name2": {\n'
                '        "HES": number\n'
                '      }\n'
                '    },\n'
                '    "Grand_Total": number\n'
                '  }\n'
                '}\n\n'
                f"Data: {json.dumps(chunk)}\n"
                f"Return the result strictly as a single JSON object—no code, no explanations, no string literal like this ```, only the JSON."
            )

            payload = {"input": prompt, "parameters": {"decoding_method": "greedy", "max_new_tokens": 8100, "min_new_tokens": 0, "temperature": 0.01}, "model_id": MODEL_ID, "project_id": PROJECT_ID}
            headers = {"Accept": "application/json", "Content-Type": "application/json", "Authorization": f"Bearer {access_token}"}

            try:
                response = requests.post(WATSONX_API_URL, headers=headers, json=payload, verify=certifi.where(), timeout=120)
                logger.info(f"WatsonX API response status code: {response.status_code}")
                logger.debug(f"WatsonX API response text: {response.text}")
                st.write(f"Debug - Response status code: {response.status_code}")

                if response.status_code == 200:
                    api_result = response.json()
                    generated_text = api_result.get("results", [{}])[0].get("generated_text", "").strip()
                    st.write(f"Debug - Raw response: {generated_text}")
                    logger.debug(f"Parsed generated text: {generated_text}")

                    if generated_text:
                        json_match = re.search(r'({[\s\S]*?})(?=\n\{|\Z)', generated_text)
                        if json_match:
                            json_str = json_match.group(1)
                            try:
                                parsed_json = json.loads(json_str)
                                chunk_result = parsed_json.get("Housekeeping", {})
                                chunk_sites = chunk_result.get("Sites", {})
                                chunk_grand_total = chunk_result.get("Grand_Total", 0)
                                logger.debug(f"Parsed JSON result: {json.dumps(parsed_json, indent=2)}")
                                logger.info(f"WatsonX results for comparison: {json.dumps(parsed_json, indent=2)}")
                                
                                # Merge chunk results into final result
                                for site, values in chunk_sites.items():
                                    if site not in result["Housekeeping"]["Sites"]:
                                        result["Housekeeping"]["Sites"][site] = {"HES": 0}
                                    result["Housekeeping"]["Sites"][site]["HES"] += values.get("HES", 0)
                                result["Housekeeping"]["Grand_Total"] += chunk_grand_total
                            except json.JSONDecodeError as e:
                                logger.error(f"JSONDecodeError: {str(e)} - Raw response: {generated_text}")
                                st.error(f"❌ Failed to parse JSON: {str(e)}")
                                st.write("Falling back to local count due to JSON parsing error")
                        else:
                            logger.error("No valid JSON found in response")
                            st.error("❌ No valid JSON found in response")
                            st.write("Falling back to local count due to invalid JSON response")
                    else:
                        logger.error("Empty generated_text from WatsonX")
                        st.error("❌ Empty response from WatsonX")
                        st.write("Falling back to local count due to empty WatsonX response")
                else:
                    error_msg = f"❌ WatsonX API error: {response.status_code} - {response.text}"
                    st.error(error_msg)
                    logger.error(error_msg)
                    st.write("Falling back to local count due to WatsonX API error")
            except Exception as e:
                error_msg = f"❌ Exception during WatsonX call: {str(e)}"
                st.error(error_msg)
                logger.error(error_msg)
                st.write("Falling back to local count due to exception")

        return result, json.dumps(result)

def generate_consolidated_ncr_excel(combined_result, report_title="Housekeeping NCR: April"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        title_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'yellow',
            'border': 1,
            'font_size': 12
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        })
        
        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        site_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })
        
        worksheet = workbook.add_worksheet('Housekeeping NCR Report')
        
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 12)
        
        data = combined_result.get("Housekeeping", {}).get("Sites", {})
        
        standard_sites = [
            "External Development",
            "Veridia- Tower 01",
            "Veridia- Tower 02",
            "Veridia- Tower 03",
            "Veridia- Tower 04",
            "Veridia- Tower 05",
            "Veridia- Tower 06",
            "Veridia- Tower 07",
            "Veridia-Club",
            "Veridia-Commercial"
        ]
        
        def normalize_site_name(site):
            if site in ["Veridia-Club", "Veridia-Commercial", "External Development"]:
                return site
            match = re.search(r'(?:tower|t)[- ]?(\d+|2021|28)', site, re.IGNORECASE)
            if match:
                num = match.group(1).zfill(2) if match.group(1) in ['01', '02', '03', '04', '05', '06', '07'] else match.group(1)
                return f"Veridia- Tower {num}"
            return site

        site_mapping = {k: normalize_site_name(k) for k in data.keys()}
        
        sorted_sites = sorted(standard_sites)
        
        worksheet.merge_range('A1:B1', report_title, title_format)
        
        row = 1
        worksheet.write(row, 0, 'Site', header_format)
        worksheet.write(row, 1, 'Knows the awareness of housekeeping', header_format)
        
        row = 2
        for site in sorted_sites:
            worksheet.write(row, 0, site, site_format)
            original_key = next((k for k, v in site_mapping.items() if v == site), None)
            value = data.get(original_key, {}).get("HES", 0) if original_key else 0
            worksheet.write(row, 1, value, cell_format)
            row += 1
        
        output.seek(0)
        return output

# Streamlit UI
st.title("Asite Housekeeping NCR Reporter")

if "df" not in st.session_state:
    st.session_state["df"] = None

st.sidebar.title("🔒 Asite Login")
email = st.sidebar.text_input("Email", "impwatson@gadieltechnologies.com", key="email_input")
password = st.sidebar.text_input("Password", "Srihari@790$", type="password", key="password_input")
if st.sidebar.button("Login"):
    session_id = login_to_asite(email, password)
    if session_id:
        st.session_state["session_id"] = session_id
        st.sidebar.success("✅ Login Successful")

st.sidebar.title("📂 Project Data")
project_name = st.sidebar.text_input("Project Name", "Wave Oakwood, Wave City")
form_name = st.sidebar.text_input("Form Name", "Non Conformity Report")
if "session_id" in st.session_state and st.sidebar.button("Fetch Data"):
    header, data, payload = fetch_project_data(st.session_state["session_id"], project_name, form_name)
    st.json(header)
    if data:
        df = process_json_data(data)
        st.session_state["df"] = df
        st.dataframe(df)
        st.success("✅ Data fetched and processed successfully!")

if st.session_state["df"] is not None:
    df = st.session_state["df"]
    
    st.sidebar.title("📋 Housekeeping NCR Report")
    start_date = st.sidebar.date_input("Start Date", df['Expected Close Date (WET)'].min())
    end_date = st.sidebar.date_input("End Date", df['Expected Close Date (WET)'].max())
    
    if st.sidebar.button("Generate Housekeeping NCR Report"):
        report_title = f"Housekeeping NCR: {end_date.strftime('%B')}"
        
        result, raw = generate_ncr_report(df, start_date, end_date)

        st.subheader("Housekeeping NCR Report (JSON)")
        st.json(result)
        st.session_state.housedf = result
        
        st.session_state.house = generate_consolidated_ncr_excel(result, report_title)
        # st.download_button(
        #     label="📥 Download Excel Report",
        #     data=excel_file,
        #     file_name=f"Housekeeping_NCR_Report_{end_date.strftime('%B')}.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # )