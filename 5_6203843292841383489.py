from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import re
import time

# Set up Selenium WebDriver
options = Options()
options.add_argument("--headless")  # Run in headless mode
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Target company URLs
company_urls = {
    "Koli Infotech": "https://koliinfotech.com/",
    "Narola Infotech": "https://www.narolainfotech.com/",
    "Avataratechnobiz": "https://avataratechnobiz.com/",
    "Bigscal Technologies": "https://www.bigscal.com/",
    "WeeTech Solution Pvt Ltd": "https://weetechsolution.com/",
    "KamalDhari Infotech": "https://kamaldhari.com/",
    "Netsol IT Solutions Pvt. Ltd.": "https://netsolitsolution.com/",
    "VPN INFOTECH": "https://vpninfotech.com/",
    "EnactOn Technologies": "https://www.enacton.com/",
    "Instance IT Solutions": "https://www.instanceit.com/",
    "Xcellence IT": "https://www.xcellence-it.com/",
    "ELaunch Solution Pvt. Ltd.": "https://www.elaunchinfotech.com/",
    "UniQual iTech": "https://uniqualitech.com/",
    "TechVizor": "https://techvizor.com/",
    "Triveni Global Software Services LLP": "https://www.triveniglobalsoft.com/",
    "Tenacious Techies": "https://www.tenacioustechies.com/",
    "Crest Infosystems Pvt. Ltd.": "https://www.crestinfosystems.com/",
    "Differenz System India": "https://www.differenzsystem.com/",
    "August Infotech": "https://www.augustinfotech.com/",
    "Sassy Infotech Pvt. Ltd.": "https://www.sassyinfotech.com/",
    "Deep Technologies": "https://www.deepit.com/",
    "Generation Next": "https://www.generationnext.in/"
}

# Function to navigate to the Contact Us page
def navigate_to_contact_page():
    try:
        contact_link = driver.find_element(By.PARTIAL_LINK_TEXT, "Contact")
        contact_url = contact_link.get_attribute("href")
        if contact_url:
            driver.get(contact_url)
            time.sleep(3)  # Allow page to load
            return True
    except Exception:
        print("Contact Us page not found, extracting from homepage.")
    return False

# Function to click the Contact button if it exists
def click_contact_button():
    try:
        contact_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Contact')]"))
        )
        contact_button.click()
        time.sleep(3)  # Allow page to update
        return True
    except Exception:
        print("No clickable 'Contact' button found.")
    return False

# Function to extract contact details dynamically
def extract_contact_info(name, url):
    driver.get(url)
    time.sleep(5)  # Allow page to load

    # Try clicking the "Contact" button if present
    button_clicked = click_contact_button()

    # If homepage has no contact details, try navigating to "Contact Us" page
    contact_page_loaded = navigate_to_contact_page() if not button_clicked else False

    # Scroll to load content
    for _ in range(3):  
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    # Parse HTML using BeautifulSoup
    soup = BeautifulSoup(driver.page_source, "html.parser")

    # Extract Emails
    emails = set(re.findall(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b", soup.text))
    valid_emails = [email for email in emails if "@" in email]

    # Extract Phone Numbers
    phones = set(re.findall(r"\+?\d{1,3}[-.\s]?\(?\d{2,5}\)?[-.\s]?\d{2,5}[-.\s]?\d{2,5}", soup.text))

    # Extract Address (Dynamically find location-related text)
    address_text = soup.get_text()

    # Extracting all possible addresses dynamically
    address_matches = re.findall(r"(?:Address|Location|Office|Head Office|Branch Office)[:\s-]+(.*)", address_text, re.IGNORECASE)
    full_address = " | ".join([addr.strip() for addr in address_matches if addr.strip()])

    # Cleaning placeholders
    if not full_address or full_address == "HEAD OFFICE - SURAT: BRANCH OFFICE - AHMEDABAD:":
        full_address = "Not Available"

    return {
        "Company Name": name,
        "Website": url,
        "Emails": ", ".join(valid_emails) if valid_emails else "Not Found",
        "Phone Numbers": ", ".join(phones) if phones else "Not Found",
        "Address": full_address
    }

# Loop through each company and extract details
data = []
for name, url in company_urls.items():
    print(f"Extracting data for: {name}")
    info = extract_contact_info(name, url)
    data.append(info)
    print(info)  # Print output for verification

# Close the Selenium WebDriver
driver.quit()

# Save results to an Excel file
df = pd.DataFrame(data)
output_file = "company_contact_details.xlsx"
df.to_excel(output_file, index=False)

print(f"\nScraping completed. Data saved to {output_file}")
