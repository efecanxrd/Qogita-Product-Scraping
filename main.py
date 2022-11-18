from openpyxl import Workbook,load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, WebDriverException
import time
import yaml
import os
with open(f'{os.getcwd()}\\config.yml') as configfile:
    config = yaml.load(configfile, Loader=yaml.FullLoader)

opt = Options()
if config['headless_mode'] == True:
    opt.add_argument("--headless")
    print('[Log] Headless mode on!')
else:
    print('[Log] Headless mode off!')
opt.add_argument("user-agent=Macbook Pro")
opt.add_experimental_option("excludeSwitches", ["enable-logging"])
browser = webdriver.Chrome(options=opt)
wb = load_workbook("veri.xlsx")
ws = wb.active

browser.get("https://www.qogita.com/products/?hitsPerPage=72&page=1")
time.sleep(config['sleep'])

while True:
    try:
        for e in range(1, 140):
            for i in range(1,73):
                veri = browser.find_element("xpath", '//*[@id="__next"]/div/div[2]/div/main/div[1]/div[1]/span')
                if veri.text.startswith("0 results found"):
                    print(f'[Log] All products has scraped. Cant find information on {e}th page. ')
                    break
                else:
                    urun = browser.find_element("xpath", f'//*[@id="__next"]/div/div[2]/div/main/div[2]/div[1]/div[{i}]') 
                    time.sleep(1)
                    urun.click()
                    time.sleep(config['sleep'])
                    raw_informations = browser.find_element("xpath", '//*[@id="__next"]/div/div[2]/div/div/div[2]/div[2]')
                    raw_informations = raw_informations.text.replace('\n', '').split(' ')[-1:]
                    raw_informations = ' '.join(raw_informations)
                    barcode = raw_informations[-13:]
                    product_name = browser.find_element("xpath", '//*[@id="__next"]/div/div[2]/div/div/div[2]/div[1]/h1') 
                    brand = browser.find_element("xpath", '//*[@id="__next"]/div/div[2]/div/div/div[2]/div[2]/div[1]/a') 
                    price = browser.find_element("xpath", '//*[@id="__next"]/div/div[2]/div/div/div[2]/div[2]/div[3]/span[2]') 
                    stock = browser.find_element("xpath", '/html/body/div[1]/div/div[2]/div/div/div[2]/form/div[1]/div[2]') 
                    stock = stock.text.replace('in stock', '')
                    
                    print("Product Name: "+product_name.text+"  |  Brand: "+brand.text+"  |  Price: "+price.text+"  |  Barcode: "+barcode+"  |  Stock: "+stock)
                    ws.append([product_name.text,brand.text,price.text,barcode,stock])
                    wb.save("veri.xlsx")
                    time.sleep(1) 
                    browser.get(f"https://www.qogita.com/products/?hitsPerPage=72&page={e}")
                    time.sleep(config['sleep'])
    except (NoSuchElementException) as e:
        print("[ERROR] Element Not Found!\nThis may be due to your internet speed or slow loading of the page.\nTry increasing the sleep time from the config file!")
        break
    except (ConnectionRefusedError,ConnectionError):
        print('[ERROR] There was a problem connecting to the page! \nCheck your internet connection.')
        break
    except (PermissionError):
        print("[ERROR] Could not access data file! \nThis error may occur if you left the file open while the program was running.")
        break
    except (WebDriverException):
        print("[ERROR] WebDriver error!\n1. Make sure there is a Chrome WebDriver file in the project directory.\n2. Make sure WebDriver is up-to-date and working.\n3. This error may also have occurred due to an Internet connection problem!\n4. And also, if it is not running in headless mode, never interfere with the program manually")
        break
