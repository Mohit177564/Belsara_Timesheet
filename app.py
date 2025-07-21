import streamlit as st
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import os, time, glob, tempfile
from datetime import timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ──────────────── Paths ────────────────
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads_temp")
FINAL_DIR    = os.path.join(os.getcwd(), "consolidated_reports")
EXPORT_FILE  = os.path.join(os.getcwd(), "clients_list.csv")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(FINAL_DIR, exist_ok=True)


# ──────────────── Helpers ────────────────
def parse_time(t):
    t = str(t).strip().lower()
    if not t: return timedelta(0)
    if ":" in t:
        h, m = map(int, t.split(":"))
    elif "." in t:
        h, m = int(float(t)), int(round((float(t) % 1) * 60))
    else:
        h, m = int(float(t)), 0
    return timedelta(hours=h, minutes=m)

def format_td(td):
    mins = int(td.total_seconds() // 60)
    return f"{mins // 60}h {mins % 60}m"

def style_excel(path):
    wb = load_workbook(path); ws = wb.active; ws.title = "Summary"
    ws.freeze_panes = ws["A2"]; ws.auto_filter.ref = ws.dimensions
    hdr_fill = PatternFill("solid", fgColor="4F81BD")
    hdr_font = Font(color="FFFFFF", bold=True)
    thin = Side("thin", color="000000")
    for c in ws[1]: c.fill, c.font = hdr_fill, hdr_font
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.row % 2 == 0:
                c.fill = PatternFill("solid", fgColor="DCE6F1")
            c.border = Border(top=thin, left=thin, right=thin, bottom=thin)
    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            max(len(str(c.value)) if c.value else 0 for c in col) + 2
    wb.save(path)

def safe_rename(src, dest):
    base, ext = os.path.splitext(dest)
    i = 1
    while os.path.exists(dest):
        dest = f"{base}_{i}{ext}"; i += 1
    os.rename(src, dest); return dest


# ──────────────── Chrome driver (uc) ────────────────
def create_driver(headless=True):
    opts = uc.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
    # Downloads
    opts.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
        },
    )
    driver = uc.Chrome(options=opts, use_subprocess=True)
    # Enable downloads in headless
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": DOWNLOAD_DIR},
    )
    return driver


# ──────────────── Page: Download Timesheets ────────────────
def page_download():
    st.title("📥 Download Timesheets")
    if st.button("⬅ Back"): st.session_state.page = "home"; st.rerun()

    un = st.text_input("Username")
    pw = st.text_input("Password", type="password")
    fr = st.text_input("FROM (DD/MM/YYYY)")
    to = st.text_input("TO (DD/MM/YYYY)")
    headless = st.toggle("Headless", value=True)

    if not st.button("🚀 Start") or not all([un, pw, fr, to]): return

    try:
        df = pd.read_csv(EXPORT_FILE)
    except FileNotFoundError:
        st.error("clients_list.csv missing"); return
    if "Client" not in df.columns:
        st.error("'Client' column missing"); return

    df["Status"] = ""
    clients = df["Client"].dropna().astype(str).tolist()
    prog    = st.progress(0.0); log = st.empty()

    for idx, client in enumerate(clients, 1):
        log.write(f"**{idx}/{len(clients)} → {client}**")
        driver = None
        try:
            driver = create_driver(headless=headless)
            wait = WebDriverWait(driver, 30)

            # Login
            driver.get("https://timesheet.outsourcinghubindia.com")
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#l-login input[name='username']"))).send_keys(un)
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#l-login input[name='password']"))).send_keys(pw)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#l-login button[type='submit']"))).click()

            # Navigate → Export page
            wait.until(EC.element_to_be_clickable((By.XPATH, "//header//li[5]/a"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='card-body']/a"))).click()

            # Select client
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.bootstrap-select > button"))).click()
            s_in = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.bs-searchbox input")))
            s_in.clear(); s_in.send_keys(client); time.sleep(1)
            found = False
            for span in driver.find_elements(By.CSS_SELECTOR, "ul.dropdown-menu.inner span.text"):
                if span.text.strip().lower() == client.lower():
                    span.click(); found = True; break
            if not found:
                df.loc[df.Client == client, "Status"] = "NotFound"; continue

            # Dates
            driver.find_element(By.XPATH, "//input[@placeholder='From Date']").clear()
            driver.find_element(By.XPATH, "//input[@placeholder='From Date']").send_keys(fr)
            driver.find_element(By.XPATH, "//input[@placeholder='To Date']").clear()
            driver.find_element(By.XPATH, "//input[@placeholder='To Date']").send_keys(to)
            driver.find_element(By.XPATH, "//button[text()='Search']").click()

            if not driver.find_elements(By.XPATH, "//table//tbody/tr"):
                df.loc[df.Client == client, "Status"] = "NoData"; continue

            before = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
            driver.find_element(By.XPATH, "//button[contains(text(),'Export XLS')]").click()

            for _ in range(50):
                time.sleep(1)
                new = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))) - before
                if new:
                    safe_rename(max(new, key=os.path.getctime), os.path.join(DOWNLOAD_DIR, f"{client}.xlsx"))
                    df.loc[df.Client == client, "Status"] = "OK"; break
            else:
                df.loc[df.Client == client, "Status"] = "DLFail"

        except Exception as e:
            st.error(f"{client}: {e}")
            df.loc[df.Client == client, "Status"] = "Error"

        finally:
            if driver: driver.quit()
            prog.progress(idx / len(clients))

    status_path = os.path.join(FINAL_DIR, "clients_with_status.csv")
    df.to_csv(status_path, index=False)
    st.success(f"All done → {status_path}")


# ──────────────── Page: Consolidate ────────────────
def page_consolidate():
    st.title("📊 Consolidate Excel")
    if st.button("⬅ Back"): st.session_state.page = "home"; st.rerun()

    files = st.file_uploader("Upload XLSX exports", type="xlsx", accept_multiple_files=True)
    if not files: return

    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f, header=3); df.columns = [c.strip() for c in df.columns]
            if not {"Employee","Process Name","Work Type"}.issubset(df.columns): continue
            tcol = next((c for c in df.columns if "time" in c.lower()), None)
            if not tcol: continue
            df[tcol] = df[tcol].apply(parse_time)
            df.rename(columns={tcol:"Time"}, inplace=True)
            dfs.append(df)
        except Exception as e:
            st.error(f"{f.name}: {e}")

    if not dfs: st.error("No valid files."); return

    summary = (pd.concat(dfs)
               .groupby(["Employee","Process Name","Work Type"])["Time"]
               .sum().reset_index())
    summary["Total"] = summary["Time"].apply(format_td); summary.drop(columns="Time", inplace=True)
    st.dataframe(summary, use_container_width=True)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with pd.ExcelWriter(tmp.name, engine="openpyxl") as w: summary.to_excel(w, index=False)
    style_excel(tmp.name)
    with open(tmp.name,"rb") as f:
        st.download_button("⬇ Download", f, "Consolidated_Summary.xlsx")


# ──────────────── Router ────────────────
if "page" not in st.session_state: st.session_state.page = "home"

if st.session_state.page == "home":
    st.title("📑 Timesheet Automation & Consolidation")
    c1,c2 = st.columns(2)
    if c1.button("📥 Download Timesheets"):  st.session_state.page="download"; st.rerun()
    if c2.button("📊 Consolidate Excel"):    st.session_state.page="consolidate"; st.rerun()
elif st.session_state.page == "download":
    page_download()
else:
    page_consolidate()
