import streamlit as st
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import camelot
import io
import tempfile
import re

# Function to extract PDF links and associated dates from the website
def get_pdf_links():
    url = import streamlit as st
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import camelot
import io
import tempfile
import re

# Function to extract PDF links and associated dates from the website
def get_pdf_links():
    url = "https://grid-india.in/en/reports/daily-psp-report"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    links = soup.find_all("a", href=True)
    pdf_links = []

    for link in links:
        href = link['href']
        if href.endswith(".pdf") and "PSP" in href:
            full_url = href if href.startswith("http") else f"https://grid-india.in{href}"
            match = re.search(r'/(\d{4})/(\d{2})/(\d{2})\.(\d{2})\.(\d{2})_', full_url)
            if match:
                date_str = f"20{match.group(5)}-{match.group(4)}-{match.group(3)}"
                try:
                    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
                    pdf_links.append((date_obj, full_url))
                except ValueError:
                    continue
    return pdf_links

# Function to extract the last table from the last page of a PDF using Camelot
def extract_last_table_from_pdf(pdf_url):
    try:
        response = requests.get(pdf_url)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(response.content)
            tmp_file_path = tmp_file.name

        # Use Camelot to read the last page
        tables = camelot.read_pdf(tmp_file_path, pages='end', flavor='stream')
        if tables and tables.n > 0:
            return tables[-1].df  # Return the last table
    except Exception as e:
        st.warning(f"Failed to process {pdf_url}: {e}")
    return None

# Streamlit UI
st.title("Daily PSP Report Extractor")

st.markdown("""
This app scrapes PDF reports from [grid-india.in](https://grid-india.in/en/reports/daily-psge of each PDF within a selected date range,  
and compiles them into a downloadable Excel file.
""")

# Date range selection
start_date = st.date_input("Start Date", datetime.today() - timedelta(days=7))
end_date = st.date_input("End Date", datetime.today())

if start_date > end_date:
    st.error("Start date must be before end date.")
else:
    if st.button("Fetch and Process Reports"):
        with st.spinner("Fetching PDF links..."):
            pdf_links = get_pdf_links()
            filtered_links = [url for date, url in pdf_links if start_date <= date <= end_date]

        if not filtered_links:
            st.warning("No PDF reports found for the selected date range.")
        else:
            all_tables = []
            for pdf_url in filtered_links:
                st.info(f"Processing: {pdf_url}")
                table_df = extract_last_table_from_pdf(pdf_url)
                if table_df is not None:
                    table_df.insert(0, "Source PDF", pdf_url)
                    all_tables.append(table_df)

            if all_tables:
                final_df = pd.concat(all_tables, ignore_index=True)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name="Combined Data")
                st.success("Processing complete!")
                st.download_button(
                    label="Download Excel File",
                    data=output.getvalue(),
                    file_name="combined_psp_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No tables could be extracted from the selected PDFs.")

    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    links = soup.find_all("a", href=True)
    pdf_links = []

    for link in links:
        href = link['href']
        if href.endswith(".pdf") and "PSP" in href:
            full_url = href if href.startswith("http") else f"https://grid-india.in{href}"
            match = re.search(r'/(\d{4})/(\d{2})/(\d{2})\.(\d{2})\.(\d{2})_', full_url)
            if match:
                date_str = f"20{match.group(5)}-{match.group(4)}-{match.group(3)}"
                try:
                    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
                    pdf_links.append((date_obj, full_url))
                except ValueError:
                    continue
    return pdf_links

# Function to extract the last table from the last page of a PDF using Camelot
def extract_last_table_from_pdf(pdf_url):
    try:
        response = requests.get(pdf_url)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(response.content)
            tmp_file_path = tmp_file.name

        # Use Camelot to read the last page
        tables = camelot.read_pdf(tmp_file_path, pages='end', flavor='stream')
        if tables and tables.n > 0:
            return tables[-1].df  # Return the last table
    except Exception as e:
        st.warning(f"Failed to process {pdf_url}: {e}")
    return None

# Streamlit UI
st.title("Daily PSP Report Extractor")

st.markdown("""
This app scrapes PDF reports from [grid-india.in](https://grid-india.in/en/reports/daily-psge of each PDF within a selected date range,  
and compiles them into a downloadable Excel file.
""")

# Date range selection
start_date = st.date_input("Start Date", datetime.today() - timedelta(days=7))
end_date = st.date_input("End Date", datetime.today())

if start_date > end_date:
    st.error("Start date must be before end date.")
else:
    if st.button("Fetch and Process Reports"):
        with st.spinner("Fetching PDF links..."):
            pdf_links = get_pdf_links()
            filtered_links = [url for date, url in pdf_links if start_date <= date <= end_date]

        if not filtered_links:
            st.warning("No PDF reports found for the selected date range.")
        else:
            all_tables = []
            for pdf_url in filtered_links:
                st.info(f"Processing: {pdf_url}")
                table_df = extract_last_table_from_pdf(pdf_url)
                if table_df is not None:
                    table_df.insert(0, "Source PDF", pdf_url)
                    all_tables.append(table_df)

            if all_tables:
                final_df = pd.concat(all_tables, ignore_index=True)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name="Combined Data")
                st.success("Processing complete!")
                st.download_button(
                    label="Download Excel File",
                    data=output.getvalue(),
                    file_name="combined_psp_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No tables could be extracted from the selected PDFs.")
