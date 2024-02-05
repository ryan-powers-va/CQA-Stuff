from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
import pandas as pd

# Setup Selenium WebDriver
driver_path = r'C:\Users\oitcopowerr\Downloads\edgedriver_win64\msedgedriver.exe'
edge_options = Options()
edge_options.add_argument('--log-level=3')
service = Service(executable_path=driver_path)
driver = webdriver.Edge(service=service, options=edge_options)

# Read the Excel file with URLs
input_excel_path = r'C:\Users\oitcopowerr\CQA Stuff\Content MASTER.xlsx'  # Correct path to your input Excel file
sheet_df = pd.read_excel(input_excel_path, sheet_name='All Urls')  # Replace 'YourSheetName' with the actual sheet name

# Ensure the DataFrame has URLs in the expected column before trying to access it
if not sheet_df.empty and sheet_df.shape[1] > 0:
    urls = sheet_df.iloc[:, 0].dropna().tolist()  # Assuming URLs are in the 1st column (index 0)

    # Initialize a list to store all the questions and titles
    all_data_list = []

    # Loop through the URLs
    for url in urls:
        # Validate and clean the URL before trying to visit
        if isinstance(url, str) and url.startswith(('http://', 'https://')):
            driver.get(url)

            # Extract the page title
            page_title_element = driver.find_element(By.CSS_SELECTOR, 'h1')
            page_title_text = page_title_element.text

            # Scrape the questions from the shadow DOM
            accordion_items = driver.find_elements(By.CSS_SELECTOR, 'va-accordion-item')

            for item in accordion_items:
                shadow_root = driver.execute_script('return arguments[0].shadowRoot', item)
                question_spans = shadow_root.find_elements(By.CSS_SELECTOR, '.header-text')

                for question_span in question_spans:
                    all_data_list.append({
                        'Question': question_span.text,
                        'Page Title': page_title_text
                    })
        else:
            print(f"Skipping invalid URL: {url}")

    # Convert the list to a pandas DataFrame
    all_data_df = pd.DataFrame(all_data_list)

    # Define the path for the output Excel file
    output_excel_path = r'C:\Users\oitcopowerr\CQA Stuff\questions2.xlsx'  # Correct path to your output Excel file

    # Save the DataFrame to an Excel file
    all_data_df.to_excel(output_excel_path, index=False)

    # Close the browser
    driver.quit()

    print(f'Scraped data has been saved to {output_excel_path}')
else:
    print(f"The sheet is empty or does not exist.")