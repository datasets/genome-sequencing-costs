import re
import csv
import os
import xlrd
import requests
import urllib.request

from bs4 import BeautifulSoup

# Base URL of the NHGRI sequencing cost data page
base_url = "https://www.genome.gov/about-genomics/fact-sheets/DNA-Sequencing-Costs-Data"
download_folder = 'archive'
out_path = 'data/sequencing_costs.csv'
os.makedirs(download_folder, exist_ok=True)

def get_latest_file_url_and_date():
    # Fetch and parse the NHGRI webpage
    response = requests.get(base_url)
    soup = BeautifulSoup(response.text, "html.parser")
    
    # Search for the latest file link
    file_link = None
    date = None
    for link in soup.find_all('a', href=True):
        if "Sequencing_Cost_Data_Table" in link['href']:
            file_link = link['href']
            # Extract the date from the link text
            date_match = re.search(r'(\d{4})', link.text)
            date = date_match.group(0) if date_match else "unknown"
            break
            
    # Formulate the complete URL if the link is relative
    if file_link and not file_link.startswith("http"):
        file_link = f"https://www.genome.gov{file_link}"
    
    return file_link, date

def execute():
    source, date = get_latest_file_url_and_date()
    
    in_path = f'{download_folder}/sequencing_costs_{date}.xls'
    
    if source:
        urllib.request.urlretrieve(source, in_path)
    else:
        print("Failed to retrieve the latest data file.")
        return
    
    workbook = xlrd.open_workbook(in_path)
    sheet = workbook.sheet_by_index(0)

    # Prepare data to write to CSV
    header = ['Date', 'Cost per Mb', 'Cost per Genome']
    records = []

    for row_idx in range(1, sheet.nrows):  # Assuming row 0 is header
        row = sheet.row_values(row_idx)
        date_cell = xlrd.xldate.xldate_as_datetime(row[0], workbook.datemode)
        date_str = f"{date_cell.year}-{str(date_cell.month).zfill(2)}"
        cost_per_mb = f"{float(row[1]):.3f}"
        cost_per_genome = f"{float(row[2]):.3f}"
        records.append([date_str, cost_per_mb, cost_per_genome])

    with open(out_path, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(records)

    print(f"Data successfully downloaded and saved to {out_path}")

execute()