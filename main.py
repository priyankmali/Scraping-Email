    from dotenv import load_dotenv
import os
import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import re
import pandas as pd
import openpyxl

options = webdriver.ChromeOptions()

options.add_argument("--headless")  # Do not open a browser window
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# https://www.instanceit.com

# https://www.narolainfotech.com
# https://elaunchinfotech.com
# https://avataratechnobiz.com
# https://www.3i-infotech.com
# https://www.bigscal.com
# https://weetechsolution.com
# https://kamaldhari.com
# https://netsolitsolution.com
# https://vpninfotech.com
# https://www.enacton.com
# https://www.xcellence-it.com
# https://uniqualitech.com
main_url = "https://uniqualitech.com" 
driver.get(main_url)

time.sleep(5)

soup = BeautifulSoup(driver.page_source, "html.parser")

header_sections = (
    soup.find("header") or
    soup.find("nav")
)

headers_links = []
contact_us_link = ""

if header_sections:
    for link in header_sections.find_all('a' , href=True):
        url = link['href']
        if url.startswith("/"):
            url = main_url + url

        if url.lower().endswith(("/contact-us/", "/contact/", "/us-contact/", "/contact-us", "/contact",
        "/us-contact" , "/enquiry/" , '/enquiry','/reach-us/' , '/reach-us')):
            contact_us_link = url

headers_links = list(set(headers_links))

email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
phone_pattern = r"\+?\d{1,3}?[-.\s]?\(?\d{2,4}?\)?[-.\s]?\d{2,5}[-.\s]?\d{4,9}"

emails = []
phones = []
if contact_us_link:
    driver.get(contact_us_link)
    time.sleep(3)

    page_soup = BeautifulSoup(driver.page_source, "html.parser")
    page_text = page_soup.get_text(separator="\n", strip=True)

    emails = re.findall(email_pattern, page_text)
    phones = re.findall(phone_pattern, page_text)


filter_phones = [num for num in phones if len(re.sub(r"\D", "", num)) >= 10]
emails = list(set(emails))
filter_phones = list(set(filter_phones))

print(emails , filter_phones)

max_length = max(len(emails) , len(filter_phones))
emails = emails + [None]*(max_length - len(emails))
filter_phones = filter_phones + [None]*(max_length - len(filter_phones))



df = pd.DataFrame({
    "Website": [main_url] * max_length ,  # Assuming all contacts are from the same website
    "Email": emails,
    "Contact": filter_phones
})

my_file = "companies_email.xlsx"

if os.path.exists(my_file):
    with pd.ExcelWriter(my_file , engine='openpyxl' ,mode='a' ,if_sheet_exists='overlay') as writer:
        sheet = writer.sheets['Sheet1']
        start_row = sheet.max_row 

        blank_df = pd.DataFrame([["", "", ""]])
        blank_df.to_excel(writer, index=False, header=False, startrow=start_row)
        df.to_excel(writer, index=False, header=False, startrow=start_row + 1)
else:
    df.to_excel(my_file , index=False , engine='openpyxl')

driver.quit()



# if url.startswith("/"):  
#             url = main_url + url
#         if main_url in url and url not in headers_links:
#             headers_links.append(url)
