# Following Script scrapes first three search results from Google
#Author : Mithilesh Shah
# Input : Clean Bank Names
# Output : Matched first three search results
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
import time

#Path to the chrome driver
path = "chromedriver_win32\chromedriver.exe"

# Function returns a dict with the results
def get_results(search_term):
    url = "https://www.google.co.in/search?q="
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    browser = webdriver.Chrome(path,options=option)
    browser.get(url+search_term)
    browser.find_element(By.ID,"L2AGLb").click()
    #search_box = browser.find_element(By.XPATH, "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/div")
    #search_box.send_keys(search_term)
    #search_box.send_keys(Keys.ENTER)
    #search_box.submit()

    try:
        links = browser.find_element(By.CLASS_NAME,"SPZz6b")
        #print(links.text)
    except :
        links = ""
    try:
        results = browser.find_elements(By.TAG_NAME,'h3')
    except:
        results = []
    if(links == ""):
        return (links,results[0].text,results[1].text, results[2].text)
    else:
        return (links.text,results[0].text,results[1].text, results[2].text)
    browser.close()

    #print(links[0])
    #results = []
    #for link in links:
      #  title = link.get_attribute("text")
      #  print(title)
       # results.append(title)


    #return results
# Load the input data-set
input = pd.read_excel('input.xlsx')
input_values = input['clean_name'].unique()
#Temp dict to store the results
final_table = []
for names in input_values :
    return_values = get_results(names)

    dict={}
    dict.update({"clean_name": names})
    dict.update({"search_term_1" : return_values[0]})
    dict.update({"search_term_2": return_values[1]})
    dict.update({"search_term_3": return_values[2]})
    dict.update({"search_term_4": return_values[3]})
    final_table.append(dict)
    print(len(final_table))

#Save the ouput to an excel file
matched = pd.DataFrame(final_table)
matched.to_excel("matched.xlsx")





