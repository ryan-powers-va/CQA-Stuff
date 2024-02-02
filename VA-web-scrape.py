from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
import pandas as pd

# Specify the path to ChromeDriver executable
driver_path = r'C:\Users\oitcopowerr\Downloads\edgedriver_win64/msedgedriver.exe'  # Use the actual path
edge_options = Options()
edge_options.add_argument('--log-level=3')  # This sets the logging level to ERRORS only

# Initialize the WebDriver
service = Service(executable_path=driver_path)
driver = webdriver.Edge(service=service, options=edge_options)

# Navigate to the page
driver.get('https://www.va.gov/health-care/about-va-health-benefits/where-you-go-for-care/')

# Extract the page title
page_title_element = driver.find_element(By.CSS_SELECTOR, 'h1')
page_title_text = page_title_element.text

# Initialize a list to store questions and titles
questions_and_titles = []

# Find all custom elements that might contain shadow-rooted questions
accordion_items = driver.find_elements(By.CSS_SELECTOR, 'va-accordion-item')

# Extract questions from the shadow root of each custom element
for item in accordion_items:
    shadow_root = driver.execute_script('return arguments[0].shadowRoot', item)
    question_span = shadow_root.find_element(By.CSS_SELECTOR, '.header-text')
    questions_and_titles.append({'Question': question_span.text, 'Page Title': page_title_text})

# Convert the list to a pandas DataFrame
df = pd.DataFrame(questions_and_titles)

# Define the Excel file path
excel_file_path = 'C:\\Users\\oitcopowerr\\CQA STuff\\questions.xlsx'

# Save the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False)

# Close the browser
driver.quit()

print(f'Scraped data has been saved to {excel_file_path}')