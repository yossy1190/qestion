import enum
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import openpyxl
import xlwings

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

# 検索情報を読み込み
amount=wsinfo.cell(4,11).value
from_finprc=wsinfo.cell(4,12).value
to_finprc=wsinfo.cell(4,13).value
from_yes=wsinfo.cell(4,14).value
to_yes=wsinfo.cell(4,15).value

# 検索結果シートの最終行を取得
Sheet_Max=wsresult.max_row+1

def main():
    driver=set_driver()
    print("実行中です。10秒程度お待ちください")
    url="https://www.traders.co.jp/market_jp/screening"
    driver.get(url)
    driver.find_element(by=By.CSS_SELECTOR,value=("#flg_sel03")).click()
    driver.find_element(by=By.CSS_SELECTOR,value=("#flg_sel04")).click()
    
    driver.find_element(by=By.CSS_SELECTOR,value=("#close_price_min")).send_keys(from_finprc)
    driver.find_element(by=By.CSS_SELECTOR,value=("#close_price_max")).send_keys(to_finprc)
    driver.find_element(by=By.CSS_SELECTOR,value=("#change_price_min")).send_keys(from_yes)
    driver.find_element(by=By.CSS_SELECTOR,value=("#change_price_max")).send_keys(to_yes)
    driver.find_element(by=By.CSS_SELECTOR,value=("#submit_btn")).click()
    time.sleep(1)
    # 検索結果が０件だったら、なしと記入したい。
    trs = driver.find_elements(by=By.CSS_SELECTOR, value="table.data_table > tbody > tr") 
    for r,tr in enumerate(trs):
        tds=tr.find_elements(by=By.CSS_SELECTOR,value="td")
        for c,td in enumerate(tds):
            wsresult.cell(row=r+4,column=c+2).value=td.text                
            wsresult.cell(row=r+4,column=c+2).value=td.text            
        wsresult.cell(row=r+4,column=1).value=amount            
    
    # 空白行を削除
    for bi in reversed(range(4,Sheet_Max)):
        GetValue1=wsresult.cell(row=bi,column=4).value
        if GetValue1==None:
            wsresult.delete_rows(bi)
        else:
            Getvalue2=GetValue1.strip()
            if len(Getvalue2)==0:
                wsresult.delete_rows(bi)
    
    wb.save("kimuland_20220505.xlsx")

main()
