from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.common.exceptions import TimeoutException

import os
from datetime import datetime
from plyer import notification
import pandas as pd
import numpy as np
import time

import multiprocessing

# Define the download directory
downloads_dir = os.path.join(os.getcwd(), "downloads")
print(f"Downloads directory: {downloads_dir}")

def get_latest_file_path(directory, extension=".xlsx"):
    files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(extension)]
    if not files:
        return None
    return max(files, key=os.path.getctime)

# Ensure the download directory exists
if not os.path.exists(downloads_dir):
    os.makedirs(downloads_dir)

def configure_options():
    options = Options()
    for arg in ["--allow-running-insecure-content", "--disable-web-security", "--unsafely-treat-insecure-origin-as-secure=http://hffsuk02"]:
        options.add_argument(arg)
    prefs = {
        "download.default_directory": downloads_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-features=InsecureDownloadWarnings")
    return options

def is_file_downloaded(directory, initial_files, timeout=60):
    """Check if a new file has been downloaded in the directory."""
    elapsed_time = 0
    while elapsed_time < timeout:
        current_files = set(os.listdir(directory))
        new_files = current_files - initial_files
        if new_files:
            new_file = new_files.pop()
            new_file_path = os.path.join(directory, new_file)
            if new_file_path.endswith(".xlsx"):
                return new_file_path
        time.sleep(1)
        elapsed_time += 1
    return None

def itemlistscraper():
    try:
        item_url = "http://hffsuk02/Reports/report/ReportsUK/Item/ItemListMDeptWC"
        options = configure_options()
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 30)  # Increased timeout to 30 seconds

        driver.get(item_url)
        driver.fullscreen_window()
        time.sleep(5)

        # Switch to frame if present
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        if frames:
            driver.switch_to.frame(frames[0]) 

        print("Navigated to URL successfully")

        # Viewing report
        print("Looking for dropdown button")
        dropdown_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='ReportViewerControl_ctl04_ctl03_ctl01']")))
        print("Dropdown button found")
        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_button)
        time.sleep(2)
        print("Clicking dropdown button")
        driver.execute_script("arguments[0].click();", dropdown_button)
        print("Dropdown button clicked")
        time.sleep(4)

        # Get page source for debugging
        page_source = driver.page_source
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(page_source)
        print("Saved page source to page_source.html for debugging")

        # Looking for 'Select All' button
        print("Looking for 'Select All' button")
        select_all_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl04_ctl03_divDropDown_ctl00")))
        driver.execute_script("arguments[0].scrollIntoView(true);", select_all_button)
        time.sleep(2)
        print("Clicking 'Select All' button")
        driver.execute_script("arguments[0].click();", select_all_button)
        time.sleep(2)
        print("Selection completed")

        # Viewing report
        print("Looking for 'View Report' button")
        view_report_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl04_ctl00")))
        driver.execute_script("arguments[0].scrollIntoView(true);", view_report_button)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", view_report_button)
        time.sleep(18)

        # Clicking dropdown button for Excel download
        print("Looking for dropdown button for Excel download")
        excel_dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_ButtonImg")))
        driver.execute_script("arguments[0].scrollIntoView(true);", excel_dropdown_button)
        time.sleep(3)
        driver.execute_script("arguments[0].click();", excel_dropdown_button)
        time.sleep(3)

        # Looking for Excel download button
        print("Looking for Excel download button")
        excel_download_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_Menu > div:nth-child(2) > a")))
        driver.execute_script("arguments[0].scrollIntoView(true);", excel_download_button)
        time.sleep(2)
        print("Clicking Excel download button")
        driver.execute_script("arguments[0].click();", excel_download_button)
        print("Download initiated")

        # Wait for the file to be downloaded
        initial_files = set(os.listdir(downloads_dir))
        latest_file = is_file_downloaded(downloads_dir, initial_files)

        driver.quit()

        if latest_file:
            df = pd.read_excel(latest_file)
            return df
        else:
            return None
    except Exception as e:
        print(f"Error in itemlistscraper: {e}")
        if 'driver' in locals():
            driver.quit()
        return None
    

def codedatescraper():
    try:
        codate_url = "http://hffsuk02/Reports/report/ReportsUK/Customer/CoDate2-X"
        options = configure_options()
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

        driver.get(codate_url)
        driver.fullscreen_window()
        time.sleep(10)

        # Switch to frame if present
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        if frames:
            driver.switch_to.frame(frames[0]) 

        print("Navigated to URL successfully")

        print("Viewing report")
        dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_ButtonImg")))
        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_button)
        time.sleep(5)
        print("Clicking dropdown button")
        driver.execute_script("arguments[0].click();", dropdown_button)
        time.sleep(10)  # Increased waiting time to ensure dropdown menu appears

        print("Looking for Excel download button")
        excel_download_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_Menu > div:nth-child(2) > a")))
        driver.execute_script("arguments[0].scrollIntoView(true);", excel_download_button)
        time.sleep(1)
        print("Clicking Excel download button")
        driver.execute_script("arguments[0].click();", excel_download_button)
        time.sleep(15)  # Increased waiting time to ensure download starts
        print("Download initiated")

        # Wait for the file to be downloaded
        latest_file = None
        attempt_time = 0
        while not latest_file and attempt_time < 60:
            latest_file = get_latest_file_path(downloads_dir)
            time.sleep(1)
            attempt_time += 1

        driver.quit()

        if latest_file:
            df = pd.read_excel(latest_file)
            return df
        else:
            return None
    except Exception as e:
        print(f"Error in codedatescraper: {e}")
        if 'driver' in locals():
            driver.quit()
        return None

def file_analysis(codate_df, itemlist_df):
    """
    Analyze the code date and item list dataframes.

    This function will:
    1. Add minutes of job for each cell in the itemlist_df.
    2. Use columns 'Buyer' and 'Quantity' from itemlist_df from both DataFrames to calculate the total quantity of each buyer.

    Args: 
        codate_df (pd.DataFrame): Code date DataFrame.
        itemlist_df (pd.DataFrame): Item list DataFrame.

    Returns:
        pd.DataFrame: DataFrame with total quantity per buyer. And total quantity per part.

    """
    # Convert input data to DataFrames
    df_codate = pd.DataFrame(codate_df)
    df_itemlist = pd.DataFrame(itemlist_df)
    
    # Check if 'CustID' column exists before dropping NaN values
    if 'CustID' in df_codate.columns:
        df_codate_cleaned = df_codate.dropna(subset=['CustID'])
    else:
        print("'CustID' column is missing in the codate DataFrame.")
        df_codate_cleaned = df_codate
    # drop rows with missing 'Quantity' in the itemlist DataFrame
    df_itemlist_cleaned = df_itemlist.dropna(subset=['Quantity'])
    df_itemlist_cleaned['MinutesOfJob'] = df_itemlist_cleaned['Quantity'] * 10  # Adjust the multiplier as needed
    df_codate_cleaned['MinutesOfJob'] = df_codate_cleaned['WCRMins'] * 10  # Adjust the multiplier as needed
    total_quantity_per_buyer = df_itemlist_cleaned.groupby('Buyer')['Quantity'].sum().reset_index()
    total_quantity_per_part = df_itemlist_cleaned.groupby('Parent')['Quantity'].sum().reset_index()
    total_quantity_per_buyer_codate = df_codate_cleaned.groupby('Buyer')['WCRMins'].sum().reset_index()
    total_quantity_per_part_codate = df_codate_cleaned.groupby('Item Number')['WCRMins'].sum().reset_index()
    total_quantity_per_buyer['Alert'] = total_quantity_per_buyer['Quantity'].apply(lambda x: 'Alert' if x > 4000 else '')
    total_quantity_per_part['Alert'] = total_quantity_per_part['Quantity'].apply(lambda x: 'Alert' if x > 4000 else '')
    total_quantity_per_buyer_codate['Alert'] = total_quantity_per_buyer_codate['WCRMins'].apply(lambda x: 'Alert' if x > 4000 else '')
    total_quantity_per_part_codate['Alert'] = total_quantity_per_part_codate['WCRMins'].apply(lambda x: 'Alert' if x > 4000 else '')
    total_quantity_per_buyer = total_quantity_per_buyer.sort_values(by = 'Quantity', ascending=False)
    total_quantity_per_part = total_quantity_per_part.sort_values(by = 'Quantity', ascending=False)
    total_quantity_per_buyer_codate = total_quantity_per_buyer_codate.sort_values(by = 'WCRMins', ascending=False)
    total_quantity_per_part_codate = total_quantity_per_part_codate.sort_values(by = 'WCRMins', ascending=False)
    # Print data processing results for verification
    print("Code date DataFrame after cleaning:")
    print(df_codate_cleaned)
    print("Item list DataFrame after cleaning and adding minutes of job:")
    print(df_itemlist_cleaned)
    print("Total quantity per buyer:")
    print(total_quantity_per_buyer)
    print(total_quantity_per_part)
    print(total_quantity_per_buyer_codate)
    print(total_quantity_per_part_codate)

    # export to Excel sheet
    with pd.ExcelWriter('Item Breakdown.xlsx') as writer:
        total_quantity_per_buyer.to_excel(writer, sheet_name='Total Quantity per Buyer', index=False)
        total_quantity_per_part.to_excel(writer, sheet_name='Total Quantity per Part', index=False)
        total_quantity_per_buyer_codate.to_excel(writer, sheet_name='Total Quantity per Buyer CoDate', index=False)
        total_quantity_per_part_codate.to_excel(writer, sheet_name='Total Quantity per Part CoDate', index=False)

    os.startfile('Item Breakdown.xlsx')
    
    return total_quantity_per_buyer, total_quantity_per_part, total_quantity_per_buyer_codate, total_quantity_per_part_codate

#-------------------------multiprocessing below (2 cores), confusing as shit----------------------------------#

if __name__ == "__main__":
    #retry logic for when it inevitably fails
    attempts = 3
    for attempt in range(attempts):
        codate_df = codedatescraper()
        if codate_df is not None:
            print(f"CoDate DataFrame loaded.")
            print(codate_df)
            break
        else:
            print(f"Attempt {attempt + 1} for CoDate scraper failed.")
        time.sleep(5)  
    #else:
        print("Failed to load CoDate DataFrame after 3 attempts.")

    #retry logic for when it inevitably fails... again
    for attempt in range(attempts):
        itemlist_df = itemlistscraper()
        if itemlist_df is not None:
            print(f"ItemList DataFrame loaded.")
            print(itemlist_df)
            break
        else:
            print(f"Attempt {attempt + 1} for ItemList scraper failed.")
        time.sleep(5)  
    else:
        print("Failed to load ItemList DataFrame after 3 attempts. Find Sean and harass him")
        
    if codate_df is not None and itemlist_df is not None:
        # Perform data processing here
        print("Data processing completed successfully. Variables stored in codate_df and itemlist_df.")
        file_analysis(codate_df, itemlist_df)
    else:
        print("Data processing failed. Check the logs for more information.")