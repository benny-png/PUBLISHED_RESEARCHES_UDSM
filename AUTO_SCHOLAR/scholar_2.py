import re
import csv
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# Initialize Chrome options for the website
brave_path = "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
options = Options()
options.binary_location = brave_path
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

# Open the Chrome WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Function to process authors
def process_authors(match):
    first_group = match.group(1)[0]
    if match.group(3):
        second_group = match.group(2)[0]
        return f"{match.group(3)} {first_group}. {second_group}."
    else:
        return f"{match.group(2)} {first_group}."

# Function to scrape paper details
def scrape_paper_details(url):
    # Navigate to the URL
    driver.get(url)
    time.sleep(1)
    # Find Authors
    try:
        authors_field = driver.find_element(
            By.XPATH,
            '//div[@class="gsc_oci_field" and contains(text(), "Authors")]')
        authors_value = authors_field.find_element(
            By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
        authors_text = authors_value.text
        formatted_authors = re.sub(r'(\w+)\s+(\w+)(?:\s+(\w+))?',
                                   process_authors, authors_text)
    except:
        authors_field = driver.find_element(
            By.XPATH,
            '//div[@class="gsc_oci_field" and contains(text(), "Inventors")]')
        authors_value = authors_field.find_element(
            By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
        authors_text = authors_value.text
        formatted_authors = re.sub(r'(\w+)\s+(\w+)(?:\s+(\w+))?',
                                   process_authors, authors_text)

    # Find Journal/Book/Source
    journal_field = None
    fields = ["Journal", "Book", "Source"]
    for field in fields:
        try:
            journal_field = driver.find_element(
                By.XPATH,
                f'//div[@class="gsc_oci_field" and contains(text(), "{field}")]'
            )
            break
        except:
            pass
    try:
        journal_value = journal_field.find_element(
            By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
        journal_text = journal_value.text
    except:
        journal_text = 'N/A'

    # Find Volume

    # Find Pages
    try:
        pages_field = driver.find_element(
            By.XPATH,
            '//div[@class="gsc_oci_field" and contains(text(), "Pages")]')
        pages_value = pages_field.find_element(
            By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
        pages_text = pages_value.text
    except:
        pages_text = 'N/A'

    try:
        volume_field = driver.find_element(
            By.XPATH,
            '//div[@class="gsc_oci_field" and contains(text(), "Volume")]')
        volume_value = volume_field.find_element(
            By.XPATH, './following-sibling::div[@class="gsc_oci_value"]')
        volume_text = volume_value.text
    except:
        volume_text = 'N/A'

    return formatted_authors, journal_text, volume_text, pages_text


with open('Research_paper_details.csv', mode='w', newline='',
          encoding='utf-8') as csv_file:
    fieldnames = [
        'NAME', 'AUTHORS', 'YEAR', 'TITLE', 'JOURNAL', 'VOLUME', 'PAGES'
    ]
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()

    with open('research_papers.csv', mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        with open('research_papers_output.txt', mode='w',
                  encoding='utf-8') as output_file:
            for row in reader:
                link = row['LINK']
                name = row['NAME']
                title = row['TITLE']
                year = row['YEAR']
                print(name)
                print(link)
                authors, journal, volume, pages = scrape_paper_details(link)

                # Construct citation format dynamically, removing 'N/A' variables
                citation_parts = []
                if authors != 'N/A':
                    citation_parts.append(authors)
                if year != 'N/A':
                    citation_parts.append(f"({year}).")
                if title != 'N/A':
                    citation_parts.append(f'{title}.')
                if journal != 'N/A':
                    citation_parts.append(f'{journal},')
                if volume != 'N/A':
                    citation_parts.append(f'{volume},')
                if pages != 'N/A':
                    citation_parts.append(f'{pages}.')

                # Adjust punctuation marks
                if 'N/A' not in (
                        journal,
                        volume):  # If either journal or volume is not 'N/A'
                    if f'{volume},' == citation_parts[
                            -1]:  # If volume is the last to be appended
                        citation_parts[-1] = ''.join(citation_parts[-1].rsplit(
                            ',', 1)) + '.'
                    elif f'{journal},' == citation_parts[
                            -1]:  # If journal is the last to be appended
                        citation_parts[-1] = ''.join(citation_parts[-1].rsplit(
                            ',', 1)) + '.'

                # Join citation parts with proper punctuation
                formatted_citation = ' '.join(citation_parts)
                output_line = f"{formatted_citation}\n"
                output_file.write(output_line)

                # Write paper details to CSV
                writer.writerow({
                    'NAME': name,
                    'AUTHORS': authors,
                    'YEAR': year,
                    'TITLE': title,
                    'JOURNAL': journal,
                    'VOLUME': volume,
                    'PAGES': pages
                })

            print("SAVED TO Research_paper_details.csv")
# Close the WebDriver
driver.quit()
