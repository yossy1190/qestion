from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import openpyxl

def set_driver(hidden_chrome: bool=False):
    USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36"
    options = ChromeOptions()
    options.add_argument('--headless')
    options.add_argument(f'--user-agent={USER_AGENT}') # ブラウザの種類を特定するための文字列
    options.add_argument('log-level=3') # 不要なログを非表示にする
    options.add_argument('--ignore-certificate-errors') # 不要なログを非表示にする
    options.add_argument('--ignore-ssl-errors') # 不要なログを非表示にする
    options.add_experimental_option('excludeSwitches', ['enable-logging']) # 不要なログを非表示にする
    options.add_argument('--incognito') # シークレットモードの設定を付与
    
    service=Service(ChromeDriverManager().install())
    return Chrome(service=service, options=options)

wb=openpyxl.load_workbook("kimuland_20220505.xlsx")
wsinfo=wb['検索情報']
wsresult=wb["検索結果"]
amount=wsinfo.cell(4,11).value
from_finprc=wsinfo.cell(4,12).value
to_finprc=wsinfo.cell(4,13).value
from_yes=wsinfo.cell(4,14).value
to_yes=wsinfo.cell(4,15).value
    # print(amount,from_finprc,to_finprc,from_yes,to_yes)

def main():
    driver=set_driver()
    url="https://www.traders.co.jp/market_jp/screening"
    driver.get(url)
    driver.find_element(by=By.CSS_SELECTOR,value=("#flg_sel03")).click()
    driver.find_element(by=By.CSS_SELECTOR,value=("#flg_sel04")).click()
    
    driver.find_element(by=By.CSS_SELECTOR,value=("#close_price_min")).send_keys(from_finprc)
    driver.find_element(by=By.CSS_SELECTOR,value=("#close_price_max")).send_keys(to_finprc)
    driver.find_element(by=By.CSS_SELECTOR,value=("#change_price_min")).send_keys(from_yes)
    driver.find_element(by=By.CSS_SELECTOR,value=("#change_price_max")).send_keys(to_yes)
    driver.find_element(by=By.CSS_SELECTOR,value=("#submit_btn")).click()
    time.sleep(2)
 
    tbody=driver.find_elements(by=By.TAG_NAME,value="tbody")
    elems=tbody[2].find_elements(by=By.TAG_NAME,value=("td"))

    for i,elem in enumerate(elems,1):
        if i<9:
            wsresult.cell(row=rows,column=1+i).value=elem.text
        elif i>=9:
            i=
             
    wb.save("kimuland_20220505.xlsx")

main()
