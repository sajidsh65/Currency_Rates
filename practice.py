from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Currency Rates"


try:
    source = requests.get("https://x-rates.com/table/?from=PKR&amount=1")
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text, "html.parser")
    
    # Find the header row
    header_row = soup.find("thead").find("tr")
    headers = [header.text for header in header_row.find_all("th")]
    
    pkRupees = headers[0]  # Adjust indices as necessary
    pkCurr = headers[1]
    otherCurr = headers[2]
    
    sheet.append([pkRupees,pkCurr,otherCurr])
    
    # Find the data rows
    data_rows = soup.find("tbody").find_all("tr")
    
    # Loop through each row and extract data
    for row in data_rows:
        cells = row.find_all("td")  # Extract all cells in the current row
        if len(cells) >= 3:  # Ensure there are enough cells
            name = cells[0].text.strip()  # First cell for name
            pkRate = cells[1].a.text.strip()  # Second cell for PKR rate
            otherRate = cells[2].a.text.strip()  # Third cell for other rate
            sheet.append([name, pkRate, otherRate])
    
except Exception as e:
    print(f"An error occurred: {e}")

excel.save("Currency Rates.xlsx")