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
import platform
import tempfile

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# PATHS
# =============================================================================
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads_temp")
FINAL_DIR    = os.path.join(os.getcwd(), "consolidated_reports")
EXPORT_FILE  = os.path.join(os.getcwd(), "clients_list.csv")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(FINAL_DIR, exist_ok=True)


# =============================================================================
# HELPERS
# =============================================================================
def parse_time(t):
    try:
        t = str(t).strip().lower()
        if not t:
            return timedelta(0)

        if ":" in t:
            h, m = map(int, t.split(":"))
        elif "." in t:
            h, m = int(float(t)), int(round((float(t) % 1) * 60))
        else:
            h, m = int(float(t)), 0
        return timedelta(hours=h, minutes=m)
    except Exception:
        return timedelta(0)


def format_td(td):
    mins = int(td.total_seconds() // 60)
    return f"{mins // 60}h {mins % 60}m"


def style_excel(path):
    wb = load_workbook(path)
    ws = wb.active
    ws.title = "Summary"
    ws.freeze_panes = ws["A2"]
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill("solid", fgColor="4F81BD")
    header_font = Font(color="FFFFFF", bold=True)
    thin       = Side(style="thin", color="000000")

    for c in ws[1]:
        c.fill, c.font = header_fill, header_font
        c.alignment    = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="DCE6F1")
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for col in ws.columns:
        length = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = length + 2

    wb.save(path)


def safe_rename(src, dest):
    base, ext = os.path.splitext(dest)
    counter   = 1
    final     = dest
    while os.path.exists(final):
        final = f"{base}_{counter}{ext}"
        counter += 1
    os.rename(src, final)
    return final


# =============================================================================
# EDGE DRIVER (factory with anti-bot patches)
# =============================================================================
def create_edge_driver(headless=True):
    opts = Options()
    opts.use_chromium = True
    opts.add_argument("--start-maximized")

    # ⭐ 1 — strip Selenium flags
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    # ⭐ 2 — typical fingerprint blockers
    opts.add_argument("--disable-blink-features=AutomationControlled")

    # ⭐ 3 — spoof regular Chrome UA
    opts.add_argument(
        '--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"'
    )

    # ⭐ 4 — legacy headless (less detectable)
    if headless:
        opts.add_argument("--headless")          # not --headless=new
        opts.add_argument("--window-size=1920,1080")

    # ⭐ 5 — download prefs
    opts.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
        },
    )

    driver = webdriver.Edge(
        service=EdgeService(EdgeChromiumDriverManager().install()),
        options=opts,
    )

    # ⭐ 6 — patch navigator.webdriver at runtime
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"},
    )

    # Allow downloads in headless
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": DOWNLOAD_DIR},
    )

    return driver


# =============================================================================
# PAGE 1 — Download Timesheets
# =============================================================================
def download_timesheets():
    st.title("📥 Download Timesheets")
    if st.button("⬅️ Back"):
        st.session_state.page = "home"
        st.rerun()

    username      = st.text_input("Username")
    password      = st.text_input("Password", type="password")
    from_date     = st.text_input("FROM Date (DD/MM/YYYY)")
    to_date       = st.text_input("TO Date (DD/MM/YYYY)")
    headless_mode = st.toggle("Headless Mode", value=True)

    if not st.button("🚀 Start") or not all([username, password, from_date, to_date]):
        return

    try:
        df_clients = pd.read_csv(EXPORT_FILE)
    except FileNotFoundError:
        st.error("clients_list.csv missing.")
        return

    if "Client" not in df_clients.columns:
        st.error("'Client' column missing.")
        return

    df_clients["Status"] = ""
    clients              = df_clients["Client"].dropna().astype(str).tolist()

    progress = st.progress(0)
    log_box  = st.empty()

    for idx, client in enumerate(clients, 1):
        log_box.write(f"**{idx}/{len(clients)} — {client}**")
        driver = None
        try:
            driver = create_edge_driver(headless=headless_mode)
            wait   = WebDriverWait(driver, 30)

            # Login
            driver.get("https://timesheet.outsourcinghubindia.com")
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#l-login input[name='username']"))).send_keys(username)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#l-login input[name='password']"))).send_keys(password)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#l-login button[type='submit']"))).click()

            # Navigate
            wait.until(EC.element_to_be_clickable((By.XPATH, "//header//li[5]/a"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='card-body']/a"))).click()

            # Select client
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.bootstrap-select > button"))).click()
            search_in = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.bs-searchbox input")))
            search_in.clear(); search_in.send_keys(client); time.sleep(1)

            found = False
            for span in driver.find_elements(By.CSS_SELECTOR, "ul.dropdown-menu.inner span.text"):
                if span.text.strip().lower() == client.lower():
                    span.click(); found = True; break
            if not found:
                df_clients.loc[df_clients["Client"] == client, "Status"] = "Not found"
                continue

            # Dates
            wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='From Date']"))).clear()
            driver.find_element(By.XPATH, "//input[@placeholder='From Date']").send_keys(from_date)
            wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='To Date']"))).clear()
            driver.find_element(By.XPATH, "//input[@placeholder='To Date']").send_keys(to_date)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Search']"))).click()

            if not driver.find_elements(By.XPATH, "//table//tbody/tr"):
                df_clients.loc[df_clients["Client"] == client, "Status"] = "No data"
                continue

            # Detect new download
            before = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Export XLS')]"))).click()

            for _ in range(45):
                time.sleep(1)
                after = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
                new   = after - before
                if new:
                    final = safe_rename(max(new, key=os.path.getctime),
                                        os.path.join(DOWNLOAD_DIR, f"{client}.xlsx"))
                    df_clients.loc[df_clients["Client"] == client, "Status"] = "OK"
                    break
            else:
                df_clients.loc[df_clients["Client"] == client, "Status"] = "DL fail"

        except Exception as e:
            st.error(f"{client}: {e}")
            df_clients.loc[df_clients["Client"] == client, "Status"] = "Error"

        finally:
            if driver: driver.quit()
            progress.progress(idx / len(clients))

    # Save status
    status_csv = os.path.join(FINAL_DIR, "clients_with_status.csv")
    df_clients.to_csv(status_csv, index=False)
    st.success(f"Done. Status → {status_csv}")


# =============================================================================
# PAGE 2 — Consolidate XLSX
# =============================================================================
def consolidate_excels():
    st.title("📊 Consolidate Excel Files")
    if st.button("⬅️ Back"):
        st.session_state.page = "home"
        st.rerun()

    files = st.file_uploader("Upload XLSX files", type="xlsx", accept_multiple_files=True)
    if not files:
        return

    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f, header=3)
            df.columns = [c.strip() for c in df.columns]
            req = {"Employee", "Process Name", "Work Type"}
            if not req.issubset(df.columns): continue
            time_col = next((c for c in df.columns if "time" in c.lower()), None)
            if not time_col: continue
            df[time_col] = df[time_col].apply(parse_time)
            df.rename(columns={time_col: "Time Worked"}, inplace=True)
            dfs.append(df)
        except Exception as e:
            st.error(f"{f.name}: {e}")

    if not dfs:
        st.error("No valid sheets.")
        return

    summary = (
        pd.concat(dfs)
        .groupby(["Employee", "Process Name", "Work Type"])["Time Worked"]
        .sum().reset_index()
    )
    summary["Total"] = summary["Time Worked"].apply(format_td)
    summary.drop(columns="Time Worked", inplace=True)
    st.dataframe(summary, use_container_width=True)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with pd.ExcelWriter(tmp.name, engine="openpyxl") as w:
        summary.to_excel(w, index=False)
    style_excel(tmp.name)

    with open(tmp.name, "rb") as f:
        st.download_button("⬇️ Download", f, "Consolidated_Summary.xlsx")


# =============================================================================
# ROUTER
# =============================================================================
if "page" not in st.session_state:
    st.session_state.page = "home"

if st.session_state.page == "home":
    st.title("📑 Timesheet Automation & Consolidation")
    col1, col2 = st.columns(2)
    if col1.button("📥 Download Timesheets"):  st.session_state.page = "download";     st.rerun()
    if col2.button("📊 Consolidate Excel"):    st.session_state.page = "consolidate";  st.rerun()
elif st.session_state.page == "download":
    download_timesheets()
else:
    consolidate_excels()
