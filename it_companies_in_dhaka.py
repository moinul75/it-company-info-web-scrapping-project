from selenium import webdriver 
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys 
import pandas as pd 
from selenium.common.exceptions import NoSuchElementException




# Company Name | Projects Done | Number Of Employees | Location
compnies_names = []
projects_dones = []
number_of_employees = []
locations = []
compnies_website = []
rank_of_company = []


#url for this 

page_number = 0


#go to the website 
driver = webdriver.Chrome()
driver.maximize_window()



#get the data main xpath 
while True:
    url = f'https://themanifest.com/bd/it-services/companies/dhaka?page={page_number}'
    driver.get(url)
    try:
        compaines_info = driver.find_elements(By.XPATH, '//li[@class="provider-card"]')

        #for one element details 
        for company in compaines_info: 
            comany_name = company.find_element(By.XPATH, './div/div/h3/a').text
            projects_done = company.find_element(By.XPATH, './div[@class="provider-card__body"]/ul/li/span').text 
            num_of_employees = company.find_element(By.XPATH, './div[@class="provider-card__body"]/ul/li[@aria-label="Employees"]/span').text 
            location = company.find_element(By.XPATH, './div[@class="provider-card__body"]/ul/li[@aria-label="Location"]/span/span').text
            website = company.find_element(By.XPATH, './div/div/h3/a').get_attribute('href').split('/?')[0]
            rank = company.find_element(By.XPATH, './span').text

            #append all this 
            compnies_names.append(comany_name)
            projects_dones.append(projects_done)
            number_of_employees.append(num_of_employees)
            locations.append(location)
            compnies_website.append(website)
            rank_of_company.append(rank)  
    except NoSuchElementException: 
        print(f"No more pages. Collected data from {page_number - 1} pages.")
        break
    page_number += 1

    if page_number >= 16:
        break


driver.quit()

#get data one by one 
data = {'Company_Name':compnies_names,'Company_Rank':rank_of_company, 'Projects_Done':projects_dones, 'Number_Of_Employee':number_of_employees,'Location':locations,'Company_Website':compnies_website}

#make data frame 
df = pd.DataFrame(data)

selected_columns = ['Company_Name', 'Company_Rank', 'Projects_Done', 'Number_Of_Employee', 'Location', 'Company_Website']
df_final = df[selected_columns]


#save to excel format
file_name = 'it_company_in_dhaka.xlsx'

with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    df_final.to_excel(writer, index=False, header=True, sheet_name='Sheet1')

    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Add a cell format with centered text alignment
    centered_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

    # Apply the cell format to all cells in the worksheet
    for col_num, value in enumerate(df_final.columns.values):
        worksheet.write(0, col_num, value, centered_format)

    for row_num in range(1, len(df_final) + 1):
        for col_num in range(len(df_final.columns)):
            worksheet.write(row_num, col_num, df_final.iloc[row_num - 1, col_num], centered_format)

print(f"Data exported to {file_name}")



