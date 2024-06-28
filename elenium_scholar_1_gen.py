import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import openpyxl

# Load the Excel file
workbook = openpyxl.load_workbook('CoICT Google Scholar.xlsx')

# Select the specific sheet
sheet = workbook['CoICT']  # Replace 'CoICT' with your actual sheet name

registered_hyperlinks = []

# Iterate through each row and check the status column
for row in sheet.iter_rows(min_row=2, min_col=1):
    status_cell = row[5]  # Assuming the status column is at index 5 (column F)

    if status_cell.value == 'Registered' and status_cell.hyperlink:
        hyperlink_address = status_cell.hyperlink.target + '&view_op=list_works&sortby=pubdate'
        name_cell = row[1]
        name_cell = name_cell.value
        registered_hyperlinks.append([name_cell, hyperlink_address])

# Close the workbook
workbook.close()

# Configure Chrome options for the website
brave_path = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
options = Options()
options.binary_location = brave_path
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

# Open the Chrome WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Initialize an empty list to store paper details
paper_details = []

# Iterate through the registered hyperlinks
for hyperlink in registered_hyperlinks:
    # Navigate to the registered hyperlink
    print(hyperlink[0])
    print(hyperlink)
    driver.get(hyperlink[1])

    # Find the name element
    #name_element = driver.find_element(By.XPATH, '//div[@id="gsc_prf_in"]')
    #name = name_element.text
    name = hyperlink[0]
    # Find all <a> tags with class 'gsc_a_at'
    elements = driver.find_elements(By.XPATH, '//a[@class="gsc_a_at"]')
    try:
        citation_element = driver.find_element(
            By.XPATH,
            '//td//a[@class="gsc_rsb_f gs_ibl" and contains(text(), "Citations")]'
        )
        citation_value = citation_element.find_element(
            By.XPATH,
            './parent::td/following-sibling::td[@class="gsc_rsb_std"]')
        citation_value = citation_value.text
    except:
        citation_value = 'N/A'

    try:
        h_index_element = driver.find_element(
            By.XPATH,
            '//td//a[@class="gsc_rsb_f gs_ibl" and contains(text(), "h-index")]'
        )
        h_index_value = h_index_element.find_element(
            By.XPATH,
            './parent::td/following-sibling::td[@class="gsc_rsb_std"]')
        h_index_value = h_index_value.text
    except:
        h_index_value = 'N/A'

    try:
        i_ten_index_element = driver.find_element(
            By.XPATH,
            '//td//a[@class="gsc_rsb_f gs_ibl" and contains(text(), "i10-index")]'
        )
        i_ten_index_value = i_ten_index_element.find_element(
            By.XPATH,
            './parent::td/following-sibling::td[@class="gsc_rsb_std"]')
        i_ten_index_value = i_ten_index_value.text
    except:
        i_ten_index_value = 'N/A'

    # Find all <span> tags with class 'gsc_a_h gsc_a_hc gs_ibl'
    span_elements = driver.find_elements(
        By.XPATH, '//span[@class="gsc_a_h gsc_a_hc gs_ibl"]')

    # Specify the title you want to stop at
    stop_year = '2022'
    # Flag to indicate whether stop title is found
    stop_found = False
    new_pub_checker = False
    for i, span in enumerate(span_elements):
        year_of_publication = span.text if span else "N/A"
        if year_of_publication == stop_year:
            print(year_of_publication)
            new_pub_checker = True
            break  # Break this specific the loop if the stop year is found

    # If stop year is not found, exit the loop
    if not new_pub_checker:
        print(f"Stop year {stop_year} not found.")
        # You can add further actions here if needed, such as logging or raising an exception
        continue

    # Iterate through the elements
    for i, element in enumerate(elements):
        year_span = span_elements[i] if i < len(span_elements) else None
        year_of_publication = year_span.text if year_span else "N/A"

        if year_of_publication == stop_year:
            stop_found = True
            break  # Stop the loop if the stop title is found

        # Extract title and link information
        title = element.text
        link = element.get_attribute('href')

        # Append paper details to the list
        paper_details.append({
            'NAME': name,
            'TITLE': title,
            'YEAR': year_of_publication,
            'LINK': link,
            'CITATIONS': citation_value,
            'H_INDEX': h_index_value,
            'I10_INDEX': i_ten_index_value
        })



# Close the WebDriver
driver.quit()

# Save the paper details to a CSV file
csv_file = "research_papers.csv"
with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.DictWriter(file, fieldnames=['NAME', 'TITLE', 'LINK', 'YEAR', 'CITATIONS', 'H_INDEX', 'I10_INDEX'])
    writer.writeheader()  # Write the header row
    writer.writerows(paper_details)  # Write the paper details

print(f"Research paper details saved to {csv_file}")
