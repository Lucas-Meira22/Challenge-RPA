from time import sleep
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager  
import pandas as pd
from selenium.webdriver.common.by import By


def wdriver():
    options = webdriver.ChromeOptions()
    options.add_argument('-headless')
    prefs = {"download.default_directory" : "Planilhas"}  #TODO
    options.add_experimental_option("prefs",prefs);
    driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options) 
    return driver

def collect_data(driver): 
    '''
        Download Exel with the data
        Transform in Pandas
    '''
    button_download = driver.find_element(By.XPATH,"//a[text() = ' Download Excel ']")
    button_download.click()
    print("Downloading the file")
    sleep(5)

    data = pd.read_excel('Planilhas/challenge.xlsx')
    return data


def etl(data,driver):

    #Starting the Challenge
    driver.find_element(By.XPATH,"/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button").click()
    print("Challange started...")

    number_users = len(data.index)

    i= 0
    while(i < number_users):

        print(f"Inserting data [{i+1}/{number_users}]")

        driver.find_element(By.XPATH,"//input[@ng-reflect-name= 'labelFirstName']").send_keys(data["First Name"][i]) #FirstName

        driver.find_element(By.XPATH,"//input[@ng-reflect-name= 'labelLastName' ]").send_keys(data["Last Name "][i])  #LastName

        driver.find_element(By.XPATH,"//input[@ng-reflect-name= 'labelCompanyName' ]").send_keys(data["Company Name"][i]) #CompanyName

        driver.find_element(By.XPATH,"//input[@ng-reflect-name= 'labelRole' ]").send_keys(data["Role in Company"][i]) #RoleCompany

        driver.find_element(By.XPATH,"//input[@ng-reflect-name= 'labelAddress' ]").send_keys(data["Address"][i]) #Address

        driver.find_element(By.XPATH,"//input[@ng-reflect-name= 'labelEmail' ]").send_keys(data["Email"][i]) #Email

        phone = str(data["Phone Number"][i])
        driver.find_element(By.XPATH,"//input[@ng-reflect-name= 'labelPhone' ]").send_keys(phone) 

        submit = driver.find_element(By.XPATH,'/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input')
        submit.click()

        i+= 1
    print("Data insert completed")

#-------------------------------------------------#
#                   Execução                      #
#-------------------------------------------------#

driver = wdriver()
driver.get('https://www.rpachallenge.com/?lang=EN')

data = collect_data(driver)

etl(data,driver)

print(driver.find_element(By.XPATH,"//div[@class = 'message2']").text) #TODO 
print("Script completed")
driver.close()

