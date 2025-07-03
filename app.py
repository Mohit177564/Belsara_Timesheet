import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import pandas as pd
import os
import time
import glob
from datetime import timedelta
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import tempfile

# --- Paths ---
EXPORT_FILE = os.path.join(os.getcwd(), "clients_list.csv")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads_temp1")
FINAL_DIR = os.path.join(os.getcwd(), "consolidated_reports")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(FINAL_DIR, exist_ok=True)


# --- Helper Functions ---
def parse_time(t):
    """Convert time strings to timedelta."""
    try:
        t = str(t).strip().lower()
        h, m = 0, 0
        if ':' in t:
            h, m = map(int, t.split(':'))
        elif '.' in t:
            parts = t.split('.')
            h = int(parts[0])
            m = int(float('0.' + parts[1]) * 60)
        elif t.isdigit():
            h = int(t)
        return timedelta(hours=h, minutes=m)
    except:
        return timedelta(0)


def format_td(td):
    """Format timedelta to 'Xh Ym'."""
    total_minutes = int(td.total_seconds() // 60)
    return f"{total_minutes // 60}h {total_minutes % 60}m"


def style_excel(file_path):
    """Apply styling to Excel file."""
    wb = load_workbook(file_path)
    ws = wb.active
    ws.title = "Summary"

    ws.freeze_panes = ws['A2']
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(border_style="thin", color="000000")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if cell.row % 2 == 0:
                cell.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

    wb.save(file_path)


# --- Streamlit App UI ---
st.set_page_config(page_title="Belsara Timesheet Downloader App", layout="wide")
st.title("📊 Timesheet Universal App")

col1, col2 = st.columns(2)

# ================================================
# 🚀 Download Timesheets Section
# ================================================
with col1:
    st.header("🔽 Download Timesheets")
    username = st.text_input("👤 Username")
    password = st.text_input("🔑 Password", type="password")
    from_date = st.text_input("📅 FROM Date (DD/MM/YYYY)")
    to_date = st.text_input("📅 TO Date (DD/MM/YYYY)")
    headless_mode = st.toggle("Run in Headless Mode", value=True)
    start_download = st.button("🚀 Start Download")

    if start_download and username and password and from_date and to_date:
        st.info("Starting automation... please wait.")
        df_clients = pd.read_csv(EXPORT_FILE)
        df_clients["Status"] = ""
        clients = df_clients['Client'].dropna().tolist()
        progress_text = st.empty()
        progress_bar = st.progress(0)

        for idx, client in enumerate(clients, 1):
            progress_text.text(f"🔄 Processing [{idx}/{len(clients)}]: {client}")
            progress_bar.progress(idx / len(clients))
            try:
                options = Options()
                options.use_chromium = True
                options.add_argument("--start-maximized")
                if headless_mode:
                    options.add_argument("--headless=new")
                    options.add_argument("--disable-gpu")
                options.add_experimental_option("prefs", {
                    "download.default_directory": DOWNLOAD_DIR,
                    "download.prompt_for_download": False,
                    "safebrowsing.enabled": True
                })

                driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
                wait = WebDriverWait(driver, 15)

                # Login
                driver.get("https://timesheet.outsourcinghubindia.com")
                wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='l-login']/form/div/div[1]/div/input"))).send_keys(username)
                wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='l-login']/form/div/div[2]/div/input"))).send_keys(password)
                wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='l-login']/form/div/button"))).click()

                # Navigate
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/header/div[2]/div/div/div/ul/li[5]/a"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[1]/a"))).click()

                # Fill client & dates
                process_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div/div/form/div/div[1]/div/div/div/button")))
                process_dropdown.click()
                search_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div/div/form/div/div[1]/div/div/div/div/div/input")))
                search_input.clear()
                search_input.send_keys(client)
                time.sleep(0.5)
                client_elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.dropdown-menu.inner span.text")))
                for element in client_elements:
                    if element.text.strip().lower() == client.lower():
                        element.click()
                        break
                wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div/div/form/div/div[3]/div/div[1]/div/input"))).send_keys(from_date)
                wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div/div/form/div/div[3]/div/div[2]/div/input"))).send_keys(to_date)
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Search']"))).click()

                try:
                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table//tbody/tr")))
                except:
                    st.warning(f"⚠️ No data for {client}. Skipping.")
                    df_clients.loc[df_clients['Client'] == client, 'Status'] = 'Skipped - No Data'
                    continue

                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Export XLS')]"))).click()
                for _ in range(10):
                    time.sleep(1)
                    files = glob.glob(os.path.join(DOWNLOAD_DIR, "task_*.xlsx"))
                    if files:
                        latest = max(files, key=os.path.getctime)
                        new_name = os.path.join(DOWNLOAD_DIR, f"{client}.xlsx")
                        os.rename(latest, new_name)
                        st.success(f"✅ Downloaded: {client}.xlsx")
                        df_clients.loc[df_clients['Client'] == client, 'Status'] = 'Downloaded'
                        break
                else:
                    st.error(f"❌ Download failed for {client}")
                    df_clients.loc[df_clients['Client'] == client, 'Status'] = 'Failed Download'
            except Exception as e:
                st.error(f"❌ Error processing '{client}': {e}")
                df_clients.loc[df_clients['Client'] == client, 'Status'] = 'Error'
            finally:
                driver.quit()
                time.sleep(1)

        progress_text.text("🎉 All clients processed!")
        progress_bar.progress(1.0)
        st.success("✅ Finished downloading all client reports.")

# ================================================
# 📊 Consolidate Excel Files Section
# ================================================
with col2:
    st.header("📊 Consolidate Excel Files")
    uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)
    if uploaded_files:
        all_data = []
        for file in uploaded_files:
            try:
                df = pd.read_excel(file, header=3)
                possible_time_cols = [c for c in df.columns if 'time' in c.lower()]
                time_col = next((c for c in possible_time_cols if 'hrs' in c.lower() or 'hour' in c.lower()), None)
                if not time_col or not {'Employee', 'Process Name', 'Work Type'}.issubset(df.columns):
                    st.warning(f"⚠️ Skipping {file.name}: Missing required columns")
                    continue
                df[time_col] = df[time_col].apply(parse_time)
                df.rename(columns={time_col: 'Time Worked'}, inplace=True)
                all_data.append(df)
            except Exception as e:
                st.warning(f"⚠️ Skipping {file.name}: {e}")

        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)
            summary = (
                combined_df.groupby(['Employee', 'Process Name', 'Work Type'])['Time Worked']
                .sum().reset_index()
            )
            summary['Total Time Worked'] = summary['Time Worked'].apply(format_td)
            summary = summary.drop(columns=['Time Worked'])
            st.dataframe(summary)

            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            with pd.ExcelWriter(tmp_file.name, engine='openpyxl') as writer:
                summary.to_excel(writer, index=False)
            style_excel(tmp_file.name)

            with open(tmp_file.name, "rb") as f:
                st.download_button(
                    "⬇️ Download Styled Excel",
                    f,
                    file_name="Consolidated_Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("❌ No valid files found.")
