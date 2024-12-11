'''import tagui as t
import numpy as np
import pandas as pd
import rpa as r

# Path to your Excel file
df = pd.read_excel(r"C:/NIC/RPA/XL to website/OSOCAsheet.xlsx")

# Ensure TagUI is initialized
t.init()
r.init(visual_automation = True)
r.click(900, 300)

# Function to create and run TagUI script
def run_tagui_script(row):
    # Open the target URL
    t.url('http://localhost:3000/inputVerify')
    t.wait(20)  # Wait for the page to load

    # Fill in the form fields with data from the row
    t.type('//*[@id="organization_name"]', row['Organization'])
    t.wait(5)
    
    t.type('//*[@id="person_name"]', row['Name of Person'])
    t.wait(5)
    
    t.type('//*[@id="address"]', row['Address Organization'])
    t.wait(5)

    # Convert row['Phone Number'] to string explicitly
    phone_number_str = str(row['Phone Number'])

    # Call the automation function with the string phone_number_str
    t.type('//*[@id="phone"]', phone_number_str)
    t.wait(5)

    t.type('//*[@id="email"]', row['Email Id (Optional)'])
    t.wait(5)
    
    # Fax_number_str = str(row['Fax Number'])

    # t.type('//*[@id="FaxNumber"]', Fax_number_str)
    # t.wait(5)

    company_button_xpath = '//*[@id="companyProfile"]'
    t.click(company_button_xpath)
    t.wait(5)
    

    file_path = row["C:/Users/KIIT/Downloads/MA-21002(DM)-CS_MID_FEB_2024.pdf"]
    t.upload('//*[@id="companyProfile"]', file_path)
    t.wait(5)

    t.type('//*[@id="numProducts"]', row['Number of products to approve:'])
    t.wait(5)

    # t.select('//*[@id="checkboxes"]', row['ProductionUnits for Input Manufacturing'])
    # t.wait(5)
     
    # t.type('//*[@id="approval"]', row['Inputs For Approval Required'])
    # t.wait(5)

    # plus_button_xpath = '//*[@id="add-btn"]'
    # t.click(plus_button_xpath)
    # t.wait(5)

    # t.type('//*[@id="capacity_unit"]', row['Capacity Of the Unit'])
    # t.wait(5)
    
    t.click('//*[@id="myMultiselect"]', row['Type Of Activities'])
    t.wait(5)

    select_button_xpath = '//*[@id="Manufacturing"]'
    t.click(select_button_xpath)
    t.wait(5)
    
    t.type('//*[@id="Product_List"]', row['Product List'])
    t.wait(5)
   
    add_button_xpath = '//*[@id="addlist"]'
    t.click(add_button_xpath)
    t.wait(5)

    t.type('//*[@id="technicalName"]', row['Product Technical Name:'])
    t.wait(5)
    
    t.type('//*[@id="ratio"]', row['Ratio:'])
    t.wait(5)
    
    plus_button_xpath = '//*[@id="plus_btn1"]'
    t.click(plus_button_xpath)
    t.wait(5)

    t.type('//*[@id="slno"]', row['Slno.:'])
    t.wait(5)

    t.type('//*[@id="inputs"]', row['Inputs:'])
    t.wait(5)

    t.type('//*[@id="plantSource"]', row['Plant Source:'])
    t.wait(5)

    t.type('//*[@id="animalSource"]', row['Animal Source:'])
    t.wait(5)

    t.type('//*[@id="microbial"]', row['Microbial:'])
    t.wait(5)

    t.type('//*[@id="mineral"]', row['Mineral:'])
    t.wait(5)

    add_button_xpath = '//*[@id="add_btn1"]'
    t.click(add_button_xpath)
    t.wait(5)

    t.type('//*[@id="locationName"]', row['Location Name:'])
    t.wait(5)

    t.type('//*[@id="location"]', row['Address:'])
    t.wait(5)

    t.type('//*[@id="storageType"]', row['Type Of Storage:'])
    t.wait(5)

    t.type('//*[@id="materialsStored"]', row['Materials Stored:'])
    t.wait(5)

    t.type('//*[@id="responsiblePerson"]', row['Responsible Person:'])
    t.wait(5)

    add2_button_xpath = '//*[@id="add_btn2"]'
    t.click(add2_button_xpath)
    t.wait(5)

    file_path = row[r"C:/Users/KIIT/Downloads/MA-21002(DM)-CS_MID_FEB_2024.pdf"]
    t.upload('//*[@id="descriptionProcessingActivity"]', file_path)
    t.wait(5)

    t.click('//*[@id="nonOrganicInputs"]', row['Any Non-Organic Inputs being produced?'])
    t.wait(5)
    
    yes1_button_xpath = '//*[@id="yes1"]'
    t.click(yes1_button_xpath)
    t.wait(5)
    
    t.click('//*[@id="additionalInput"]', row['Do you need additional inputs?'])
    t.wait(5)
    
    yes2_button_xpath = '//*[@id="yes2"]'
    t.click(yes2_button_xpath)
    t.wait(5)

    tick_button_xpath = '//*[@id="declarationCheckbox1"]'
    t.click(tick_button_xpath)
    t.wait(5)


    # Click the submit button
    submit_button_xpath = '//*[@id="sub_mit"]'
    t.click(submit_button_xpath)
    t.wait(5)  # Wait for the submission to complete

# Iterate through each row in the Excel sheet and execute the script
for index, row in df.iterrows():
    run_tagui_script(row)
    t.wait(5)  # Add a delay to ensure each form submission completes

# Close TagUI after the loop
t.close()
'''
import tagui as t
import rpa as r
import pandas as pd


# Path to your Excel file
df = pd.read_excel(r"C:/NIC/RPA/XL to website/OSOCAsheet.xlsx")

# Initialize TagUI

r.init(visual_automation=True)

# Function to create and run TagUI script
def run_tagui_script(row):
    try:
        # Open the target URL
        t.url('http://localhost:3000/inputVerify')
        t.wait(5)  # Wait for the page to load

        # Fill in the form fields with data from the row
        t.type('//*[@id="organization_name"]', row['Organization'])
        t.wait(1)
        
        t.type('//*[@id="person_name"]', row['Name of Person'])
        t.wait(1)
        
        t.type('//*[@id="address"]', row['Address Organization'])
        t.wait(1)
        # Convert row['Phone Number'] to string explicitly
        phone_number_str = str(row['Phone Number'])

        # Call the automation function with the string phone_number_str
        t.type('//*[@id="phone"]', phone_number_str)
        t.wait(1)

        t.type('//*[@id="email"]', row['Email Id (Optional)'])
        t.wait(1)
        
        # Click the button to open the file dialog
        t.click('//*[@id="companyProfile"]')
        # t.wait(5)

        # Use rpa to handle the file upload dialog
        # file_path = row[r"C:/Users/KIIT/Downloads/MA-21002(DM)-CS_MID_FEB_2024.pdf"]
        # r.click("file.png")
        # r.type(file_path)
        # r.keyboard('[ctrl][l]')
        # r.keyboard(file_path)
        # r.keyboard('[enter]')
        # t.wait(5)
        r.keyboard(r"C:\Users\KIIT\Downloads\EAadhaar_2724110521976020220829111956_14062024132536.pdf")
        r.keyboard('[enter]')
        t.click('//*[@id="numProducts"]')
        t.wait(2)
        products_number_str = str(row['Number of products to approve:'])
        t.type('//*[@id="numProducts"]', products_number_str)
        t.wait(1)
        
        # r.keyboard(r"C:\Users\KIIT\Downloads\EAadhaar_2724110521976020220829111956_14062024132536.pdf")
        # r.keyboard('[enter]')
        t.click('//*[@id="mySelectLabel"]')
        t.wait(1)
        
        t.click('//*[@id="checkboxes"]')
        t.wait(1)
        
        t.click('//*[@id="Product_List"]')
        t.type('//*[@id="Product_List"]', row['Product List'])
        t.wait(1)
        
        t.click('//*[@id="addlist"]')
        t.wait(1)
        
        t.type('//*[@id="technicalName"]', row['Product Technical Name:'])
        t.wait(1)
        
        ratio_number_str = str(row['Ratio:'])
        t.type('//*[@id="ratio"]', ratio_number_str)
        t.wait(1)
        
        t.click('//*[@id="plus-btn1"]')
        t.wait(1)
        
        t.click('//*[@id="slno"]')
        t.wait(1)

        sl_number_str = str(row['Slno.:'])
        t.type('//*[@id="slno"]', sl_number_str)
        t.wait(1)
        
        t.type('//*[@id="inputs"]', row['Inputs:'])
        t.wait(1)
        
        t.type('//*[@id="plantSource"]', row['Plant Source:'])
        t.wait(1)
        
        t.type('//*[@id="animalSource"]', row['Animal Source:'])
        t.wait(1)
        
        t.type('//*[@id="microbial"]', row['Microbial:'])
        t.wait(1)
        
        t.type('//*[@id="mineral"]', row['Mineral:'])
        t.wait(1)
        
        t.click('//*[@id="add-btn1"]')
        t.wait(1)
        
        t.type('//*[@id="locationName"]', row['Location Name:'])
        t.wait(1)
        
        t.type('//*[@id="location"]', row['Address:'])
        t.wait(1)
        
        t.type('//*[@id="storageType"]', row['Type Of Storage:'])
        t.wait(1)
        
        t.type('//*[@id="materialsStored"]', row['Materials Stored:'])
        t.wait(1)
        
        t.type('//*[@id="responsiblePerson"]', row['Responsible Person:'])
        t.wait(1)
        
        t.click('//*[@id="add-btn2"]')
        t.wait(1)
        
        # Handle the second file upload
        t.click('//*[@id="description"]')
        t.click('//*[@id="description"]')
        t.wait(5)
        # r.click('//*[@id="description"]')

        r.keyboard(r"C:\Users\KIIT\Downloads\EAadhaar_2724110521976020220829111956_14062024132536.pdf")
        r.keyboard('[enter]')

        # r.type(file_path)
        # r.keyboard('[enter]')
        # t.wait(5)
        
        t.click('//*[@id="nonOrganicInputs"]')
        t.wait(1)

        t.type('//*[@id="nonOrganicInputs"]', row['Any Non-Organic Inputs being produced?'])
        t.wait(1)
        
        t.click('//*[@id="yes1"]')
        t.wait(1)
        
        t.click('//*[@id="additionalInput"]')
        t.wait(1)

        t.type('//*[@id="additionalInput"]', row['Do you need additional inputs?'])
        t.wait(1)
        
        t.click('//*[@id="yes2"]')
        t.wait(1)
        
        t.click('//*[@id="declarationCheckbox1"]')
        t.wait(1)
         
        t.click('//*[@id="date"]')
        t.wait(1)

        date_number_str = str(row['Date'])
        t.type('//*[@id="date"]', date_number_str)
        t.wait(1)

        # Click the submit button
        t.click('//*[@id="sub_mit"]')
        t.wait(1)  # Wait for the submission to complete

    except Exception as e:
        print(f"An error occurred: {e}")

# Iterate through each row in the Excel sheet and execute the script
for index, row in df.iterrows():
    run_tagui_script(row)
    t.wait(5)  # Add a delay to ensure each form submission completes

# Close TagUI after the loop
t.close()
r.close()
