from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import openpyxl
import time
import random

url = 'very secret url'

# Set up Chrome options to simulate a mobile device
mobile_emulation = {
    "deviceMetrics": {"width": 360, "height": 640, "pixelRatio": 3.0},
    "userAgent": "Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1"
}

chrome_options = Options()
chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)
#chrome_options.add_argument('--headless')
chrome_options.add_argument('--ignore-certificate-errors')

#chrome_options.add_argument('user-agent=Mozilla/5.0 (Linux; Android 8.0.0; Pixel 2 Build/OPD3.170816.012)')

driver = webdriver.Chrome(options=chrome_options)

try:
    driver.get(url)

    # Scroll to load dynamic content
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        WebDriverWait(driver, 15).until(
            lambda driver: driver.execute_script("return document.body.scrollHeight") > last_height
        )
        last_height = driver.execute_script("return document.body.scrollHeight")

        # Add a condition to break the loop if you've scrolled to the end of the content
        # For example, if you know there are a certain number of items to load

        # You can adjust the condition based on your specific case

except Exception as e:
    print(f"Error during scrolling: {e}")

    # Get the updated page source after scrolling
    html = driver.page_source

    # Parse the updated page source with BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')

    # Find and extract the desired information based on itemprop attributes
    items = soup.find_all('div', {'itemprop': 'offers'})

    # Create a new Excel workbook and add a worksheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Add headers to the worksheet
    ws.append(['Title', 'Price', 'Description'])
    processed_urls = set()  # Создаем множество для отслеживания обработанных URL

    for item in items:
        title_element = item.find_previous('p', {'itemprop': 'name'})
        price_element = item.find('div', {'itemprop': 'price'})
        url_element = item.find_next('a', {'itemprop': 'url'})

        # Check if elements are found before accessing their text attribute
        title = title_element.text.strip() if title_element else "N/A"
        price_text = price_element.text.strip() if price_element else "N/A"

        # Strip ₽ from the price and remove spaces
        price = price_text.replace('₽', '').replace('\xa0', '')

        # Check if the title contains "iphone 15"
        if 'iphone 15' in title.lower():
            url = 'https://m.avito.ru' + url_element['href'] if url_element else None

            # Check if the URL has already been processed
            if url and url not in processed_urls:
                print(f"Processing URL: {url}")
                driver.get(url)
                time.sleep(random.uniform(3, 15))
                # Add the URL to the set of processed URLs
                processed_urls.add(url)

                # Wait for the page to load and find the color information
                try:
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, '[data-marker^="item-properties-item"]'))
                    )


                    # Find all elements with data-marker attribute
                    data_marker_elements = driver.find_elements(By.CSS_SELECTOR, '[data-marker^="item-properties-item"]')
                    # Initialize color_title and color_value to None
                    color_title = None
                    color_value = None

                    # Iterate through elements and find Color information
                    for element in data_marker_elements:
                        data_marker = element.get_attribute('data-marker')
                        if 'title' in data_marker:
                            color_title = element.text.strip()
                        elif 'description' in data_marker:
                            color_value = element.text.strip()

                    # Add data to the worksheet only if the "Цвет" (Color) property is present
                    if color_title == 'Цвет':
                        ws.append([title, color_value, price])
                    time.sleep(3)
                except Exception as e:
                    print(f"Error processing URL: {e}")
finally:
    # Save the Excel workbook
    wb.save('MISHKA.xlsx')
    driver.quit()
