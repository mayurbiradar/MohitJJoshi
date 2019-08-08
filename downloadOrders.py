from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import pandas as pd
import os, sys
def getResult(fdate,tdate,download_directory,url,driver_path,data_file_name):    
    chrome_options = webdriver.ChromeOptions()
    prefs = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}],
             "download.default_directory" : download_directory,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True}
    chrome_options.add_experimental_option("prefs",prefs)
    chrome_options.add_argument("--mute-audio")
    #chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(executable_path=driver_path,chrome_options=chrome_options)
    driver.set_page_load_timeout(200)
    driver.get(url)  
    element = driver.find_element_by_id("from_date")
    element.send_keys(fdate)
    element.send_keys(Keys.RETURN)
    element = driver.find_element_by_id("to_date")
    element.send_keys(tdate)
    element.send_keys(Keys.RETURN)

    ad = EC.element_to_be_clickable((By.XPATH,'//a[img/@src="../images/speaker7.png"]'))
    adb = WebDriverWait(driver, 30).until(ad)
    adb.click()
     
    def readCaptcha():
        capt="";
        for var in list(range(4)):
            try:
                element=driver.find_element_by_id("audio"+str(var+1))
                url=str(element.get_attribute("src"))
                arr=url.split("/")
                if capt=="":
                    capt=str(arr[len(arr)-1])[0]
                else:
                    capt=capt+str(arr[len(arr)-1])[0]
            except Exception as e:
                time.sleep(4)
        return capt
                                        
    while(len(readCaptcha())!=4):
        time.sleep(1)
        
    capt=readCaptcha()
    element = driver.find_element_by_id("captcha")
    element.send_keys(capt)
    element.send_keys(Keys.RETURN)
    
    button = driver.find_element_by_name("submit1")
    button.click()
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.TAG_NAME,"tr")))
    trs = driver.find_elements(By.TAG_NAME, "tr") 

    length = len(trs)
    try:
        os.chdir(download_directory)
    except FileNotFoundError as e:
        os.mkdir(download_directory)
        
        
    os.chdir(download_directory) 
    srlist=[]
    yearList=[]
    orderDateList=[]
    oNumberList=[]
    for i in range(length): 
        tds = trs[i].find_elements(By.TAG_NAME, "td")
        if(len(tds)>0):
            srlist.append(tds[0].text)
            yearList.append(tds[1].text)
            orderDateList.append(tds[2].text)
            if(os.path.exists("display_pdf.pdf")):
                os.remove("display_pdf.pdf")
            (tds[3].find_element_by_css_selector('a')).click()
            counter=0
            while(os.path.exists("display_pdf.pdf") is not True and counter is not 10):
                time.sleep(1)
                counter=counter+1        
            time.sleep(2)
            try:
                new_name=(str(tds[1].text).replace("/","_")).replace(".","_")+str(tds[2].text)+".pdf"
                os.rename("display_pdf.pdf",new_name)
                oNumberList.append(download_directory+new_name)
            except Exception as e:
                oNumberList.append("File Not Found")
                continue
                          
    df = pd.DataFrame({ "Sr No": srlist,
                        "Case Type/Case Number/Case Year": yearList,
                        "Order Date	": orderDateList,
                        "Order Number": oNumberList})
    df.to_excel(data_file_name)
    driver.close()
    return df

if __name__ == "__main__":
    fdate="01-08-2019"
    tdate="01-08-2019"
    download_directory="C:\\pythonpractice3\\docs\\sikkim\\"
    driver_path="C:\\Users\\mbiradar\\Desktop\\RapidDeployer docs\\BI scripts\\chromedriver.exe"
    url="https://services.ecourts.gov.in/ecourtindiaHC/cases/s_orderdate.php?state_cd=24&dist_cd=1&court_code=1&stateNm=Sikkim#"
    #url="https://services.ecourts.gov.in/ecourtindiaHC/cases/s_orderdate.php?state_cd=25&dist_cd=1&court_code=1&stateNm=Manipur"
    data_file_name=fdate+tdate+".xlsx"    
    df=getResult(fdate,tdate,download_directory,url,driver_path,data_file_name)
    




