import time
import  selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Thoothukudi_Collectors_list"
sheet.append(['Sl. No', 'Collector Name', 'From', 'To'])

service = Service(executable_path = "chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.get("https://thoothukudi.nic.in/about-district/list-of-collectors/")
time.sleep(20)

cur_sl_no = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/thead/tr[2]/td[1]").text
cur_collector_name = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/thead/tr[2]/td[2]").text
cur_from_date = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/thead/tr[2]/td[3]").text
cur_to_data = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/thead/tr[2]/td[4]").text

print(cur_sl_no)
print(cur_collector_name)
print(cur_from_date)
print(cur_to_data)

for i in range(1,26):
    sl_no = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/tbody/tr["+str(i)+"]/td[1]").text
    collector_name = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/tbody/tr["+str(i)+"]/td[2]").text
    from_date = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/tbody/tr["+str(i)+"]/td[3]").text
    to_date = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div[3]/div/div/table/tbody/tr["+str(i)+"]/td[4]").text


    print(sl_no)
    print(collector_name)
    print(from_date)
    print(to_date)
    sheet.append([sl_no, collector_name, from_date, to_date])

excel.save("Thoothukudi_Collectors_list.xlsx")
