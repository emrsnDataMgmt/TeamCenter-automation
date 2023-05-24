__author__ = """Nripesh Niketan"""
__email__ = """nripesh.niketan@emerson.com"""

from datetime import datetime
import os
import shutil
import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import win32com.client
import PySimpleGUI as sg
from pydomo import Domo

domo = Domo('c7b587d1-08a8-4931-b71a-bb793fcd82de',
            '1fd4cda307f6df15bb52e3183fab9e387e1d95c044dfa0b3c4fe9f80c1423e37')

def read_excel_file(file_name, sheet_name):
    df = pd.read_excel(file_name, sheet_name=sheet_name)
    parts = df['Part#'].unique()
    parts = parts[~pd.isnull(parts)]
    print(parts)
    return parts


def login_to_website(driver, user, password):
    driver.get('http://prod-crtcr.emerson.com/awc/#/showHome')
    driver.maximize_window()

    input_field = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '//input[@type="text"]')))
    input_field.send_keys(user)

    input_field = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '//input[@type="password"]')))
    input_field.send_keys(password)

    sign_in_button = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//button[@class="aw-base-button ng-binding" and @ng-click="login()"]')))
    sign_in_button.click()
    # time.sleep(1000)

def search_and_download_parts(driver, parts):
    out = pd.DataFrame(columns=['Part#', 'Status'])
    search_input = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '//input[@type="text" and @aria-label="Search Box" and @placeholder="Search"]')))
    advanced_search = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-view"]/aw-include/div/div/div/div/div/ui-view/aw-page/div/div/div/div[1]/aw-header/div/header/div[2]/aw-global-search/div/aw-include/div/aw-search-global/div/div[1]/div[1]/div[2]/div[2]/div[2]/aw-link/div/a')))
    advanced_search.click()
    # //*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/aw-tab-container/div[1]/div[1]/ul/li[2]/a
    advance = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/aw-tab-container/div[1]/div[1]/ul/li[2]/a')))
    advance.click()
    # //*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/div/div[2]/form/aw-advsearch-lov-val/div/aw-property-error/div/div/input
    drop_down = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/div/div[2]/form/aw-advsearch-lov-val/div/aw-property-error/div/div/input')))
    drop_down.click()
    # //*[@id="ui-id-4"]/div/div/div/div/ul/li[13]/aw-property-lov-child/div/div[2]/div[1]
    option = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "aw-aria-border") and .//div[text()="General..."]]')))
    option.click()
    # //*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/div/form[2]/div/form/div[4]/aw-checkbox-multiselect/div/div/div/aw-checkbox-list/div/div/input
    owner = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/div/form[2]/div/form/div[4]/aw-checkbox-multiselect/div/div/div/aw-checkbox-list/div/div/input')))
    owner.click()
    checkbox_span = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[contains(@class, "afx-checkbox-md-style")]')))
    checkbox_span.click()
    owner.click()
    time.sleep(1)
    # //*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/div/form[2]/div/form/div[3]/aw-checkbox-multiselect/div/div/div/aw-checkbox-list/div/div/input
    input_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/div/form[2]/div/form/div[3]/aw-checkbox-multiselect/div/div/div/aw-checkbox-list/div/div/input')))
    input_element.click()
    # //*[@id="ui-id-14"]/div/div/div/div/div/input
    input_element1 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@type="text" and contains(@class, "ng-pristine") and @autocomplete="off" and @ng-model="listFilterText"]')))
    input_element1.click()
    input_element1.send_keys("PDF")
    pdf_label = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//label[contains(@class, "aw-jswidgets-checkboxLabel") and @title="PDF"]')))
    pdf_label.click()
    input_element.click()
    # time.sleep(100)
    
    out = pd.DataFrame(columns=['Part#', 'Status'])
    
    for part in parts:
        
        print(out)
        
        try:
        
            search_part = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/form/aw-tab-set/div/form[2]/div/form/div[1]/aw-widget/div/div/div/aw-property-val/div/div/aw-property-string-val/div/div/aw-property-text-area-val/div/aw-property-error/div/textarea')))
            search_part.click()
            search_part.clear()
            search_part.send_keys(part)
            search_button= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="aw_navigation"]/div/aw-include/div/div/div[2]/div/div/button[1]')))
            search_button.click()
            time.sleep(1)
            download_pdf = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-view"]/aw-include/div/div/div/div/div/ui-view/aw-page/div/div/div/div[2]/div/ng-transclude/ui-view/aw-sublocation/div/div/div/main/aw-sublocation-body/div[1]/aw-primary-workarea/div/aw-include/div/aw-scrollpanel/aw-list/div/div/div/ul/li[1]/div/div[2]/div/aw-cell-command-bar/aw-command[2]/button')))
            # clear input field
            # Find the element using the class name
            # element = driver.find_element(By.CSS_SELECTOR, ".aw-widgets-propertyValue")
            
            # time.sleep(1000)
            # <label class="aw-widgets-propertyValue aw-base-small ng-isolate-scope" title="001" aw-highlight-property-html="" display-val="cellProp.value">001</label>
            element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/aw-include/div/div/div/div/div/ui-view/aw-page/div/div/div/div[2]/div/ng-transclude/ui-view/aw-sublocation/div/div/div/main/aw-sublocation-body/aw-secondary-workarea/div/aw-selection-summary/div/aw-include/div/aw-xrt-summary/div[2]/aw-xrt-2/aw-walker-view/div[1]/aw-walker-element/div[3]/form/aw-walker-element/div/div/aw-walker-objectset/div[2]/aw-list/div/div/div/ul/li/div/div[1]/aw-default-cell/div/div[2]/aw-default-cell-content/div[2]/div/label[2]')))

            # Extract the value from the 'title' attribute
            value = element.get_attribute("title")

            # Print the extracted value
            print(value)
            # time.sleep(1000)
            download_pdf.click()
            out= out.append({'Part#': part, 'Status': 'Success', 'Revision#': value}, ignore_index=True)
            
            time.sleep(3)
            
            
        
        except TimeoutException:
            print(f"No result found for {part}")
            out = out.append({'Part#': part, 'Status': f"No result found for {part}"}, ignore_index=True)
            continue
    return out

def send_mail(ol, subject, body, to, cc, attachment=None):
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = subject
    for i in to:
        newmail.Recipients.Add(i)
    # if cc is not None:
    if cc != '':
        for i in cc:
            newmail.Recipients.Add(i)
    newmail.Body= body
    attach = attachment
    newmail.Attachments.Add(attach)
    newmail.Send()
    
# move from downloads folder to another folder
def move_file(parts, folder):
    path = os.path.expanduser('~') + '\\OneDrive - Emerson\\Subash.Bose\\Downloads'
    for part in parts:
        for file in os.listdir(path):
            if part in file:
                # if part already exists in folder, delete it
                if os.path.exists(os.path.join(folder, file)):
                    os.remove(os.path.join(folder, file))
                shutil.move(os.path.join(path, file), folder)
                break
            

if __name__ == '__main__':
    
    
    sg.theme('DarkAmber')
    # ask user either to download all pdfs or blast email with 2 different buttons
    layout = [[sg.Text('Do you want to download all pdfs or send email to supplier?')],
                [sg.Button('Download PDFs'), sg.Button('Send Email')]]
    window = sg.Window('Teamcenter Automation', layout)
    while True:
        event, values = window.read()
        
        if event == 'Download PDFs':
            # Define variables
            user = 'subash.bose'
            password = 'Guts2787'
            

            supplier = pd.read_excel('Supplier Contact.xlsx')
            
            # Perform the query
            query = """
            SELECT ITEM_NUMBER, VENDOR_NAME
            FROM table
            WHERE (ASL_DISABLE_FLAG <> 'Y' OR ASL_DISABLE_FLAG IS NULL OR TRIM(ASL_DISABLE_FLAG) = '')
            AND ASL_ORG_CODE = 'P56'
            """

            df = domo.ds_query('3d555ed9-a5dd-4002-9036-48523149b9d9', query)
            
            # ask user for input file
            file_name = sg.popup_get_file('Please select the input file', title='Teamcenter Automation', file_types=(("Excel Files", "*.xlsx"),))
            # file_name = 'Part details.xlsx'
            sheet_name = 'DATA'
            # Read the excel file and get the parts list
            parts = read_excel_file(file_name, sheet_name)
            # Initialize the webdriver
            chrome_driver = os.path.expanduser('~') + '\\OneDrive - Emerson\\Subash.Bose\\Documents\\Team Center Drawings\\Teamcenter automation\\chromedriver.exe'
            driver = webdriver.Chrome(chrome_driver)
            # Login to the website
            login_to_website(driver, user, password)

            # Search and download each part
            out = search_and_download_parts(driver, parts)
            
            out = pd.merge(out, df, how='left', left_on='Part#', right_on='ITEM_NUMBER')
            
            # left join with supplier contact
            out = pd.merge(out, supplier, how='left', left_on='VENDOR_NAME', right_on='Supplier Name')
            
            if 'Email Sent' not in out.columns:
                        out['Email Sent'] = ''
                        out['Email Sent Date'] = ''
            # convert revision to integer
            out['Revision#'] = out['Revision#'].astype(int)
            # append the output to a csv file instead of overwriting, create header if not exists
            if os.path.exists('output.csv'):
                out.to_csv('output.csv', mode='a', index=False, header=False)
            else:
                out.to_csv('output.csv', index=False)
            
            move_file(parts, os.path.expanduser('~') + '\\OneDrive - Emerson\\Subash.Bose\\Documents\\Team Center Drawings\\Teamcenter Downloads')
            
            # Close the webdriver
            driver.quit()
            sg.popup('Download completed', title='Teamcenter Automation')
        
        elif event == 'Send Email':
            ol = win32com.client.Dispatch("Outlook.Application")
            # read the output file
            out = pd.read_csv('output.csv')
            for i in range(len(out)):
                # if status is success, send email
                if out['Status'][i] == 'Success' and out['Email Sent'][i] != 'Yes' and out['Contact email'][i] != 'None' and out['Contact email'][i] != 'nan' and out['Contact email'][i] != '':
                    subject = f"Emerson Dubai Part number {out['Part#'][i]} Drawing Revision {out['Revision#'][i]}"
                    body = f"Hi {out['Supplier Name'][i]},\n\nFind the attached latest revision for the drawings. Please update in your system and acknowledge the reciept.\n\nRegards,\nSubash Bose"
                    # to = out['Contact email'][i]
                    
                    # to is a list of all the email ids
                    if 'Contact email' in out and isinstance(out['Contact email'][i], str):
                        to = out['Contact email'][i].split(';')
                    else:
                        break
                    print(to)
                    
                    if 'Contact email B' in out and isinstance(out['Contact email B'][i], str):
                        cc = out['Contact email B'][i].split(';')
                    else:
                        cc = ''
                    attachment = os.path.expanduser('~') + '\\OneDrive - Emerson\\Subash.Bose\\Documents\\Team Center Drawings\\Teamcenter Downloads\\' + out['Part#'][i] + '.pdf'
                    send_mail(ol, subject, body, to, cc, attachment)
                    # add a column to the dataframe with email sent status as yes and date
                    stat = 'Yes'
                    out.loc[i, 'Email Sent'] = stat
                    out.loc[i, 'Email Sent Date'] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            out.to_csv('output.csv', index=False)
            # popup message
            sg.popup('Email sent successfully')
            
        elif event == sg.WIN_CLOSED:
            break
            