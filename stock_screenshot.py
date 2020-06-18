from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import time, os, datetime, shutil
from Screenshot import Screenshot_Clipping

today = datetime.date.today().strftime("%Y-%m-%d")
#刪除舊資料
shutil.rmtree("/Users/kkday/Documents/Python_Codes/stock_screenshot")
print('刪除舊資料！')

#將chrome設定於背景執行
options = Options()
options.add_argument('--headless')
#options.add_argument('--disable-gpu') # 允許在無GPU的環境下運行，可選
options.add_argument('--window-size=1280,1024') # 建議設置

#chromedriver的路徑
ob=Screenshot_Clipping.Screenshot()
driver = webdriver.Chrome(executable_path = '/Users/kkday/Downloads/chromedriver',chrome_options = options)
driver.get('https://be2.kkday.com/auth/login')
driver.find_element_by_xpath('//*[@id="username"]').send_keys('alice.lee@kkday.com')
driver.find_element_by_xpath('//*[@id="login"]/div[3]/input').send_keys('charlie1111')
driver.find_element_by_xpath('//*[@id="login"]/div[4]/div[2]/button').click()
driver.find_element_by_xpath('/html/body/div/aside[1]/section/ul/li[9]/a').click()

print('登入成功！')

time.sleep(1)
driver.find_element_by_xpath('/html/body/div/aside[1]/section/ul/li[9]/ul/li[3]/a').click()
time.sleep(1)

s1 = Select(driver.find_element_by_xpath('//*[@id="wh_upload_type"]'))  #ddl 0
s1.select_by_index(0) 
time.sleep(1)

driver.find_element_by_xpath('//*[@id="selectBtn"]').click()
time.sleep(5)

driver.maximize_window()
img_url=ob.full_Screenshot(driver, save_path=r'.', image_name='Myimage.png')
os.remove(img_url)

TotalPage = driver.find_element_by_xpath('//*[@id="ngApp"]/div/div[2]/div[2]/div/span').text.replace('Total: ','')
print("總共有 " + str(TotalPage) + " 筆資料")
print("資料共 " + str(int(TotalPage)//10 +1) + " 頁")
#生成資料夾放檔案
os.makedirs("/Users/kkday/Documents/Python_Codes/stock_screenshot")

scroll_width = driver.execute_script('return document.body.parentNode.scrollWidth')
scroll_height = driver.execute_script('return document.body.parentNode.scrollHeight')
driver.set_window_size(scroll_width, scroll_height)
driver.save_screenshot('/Users/kkday/Documents/Python_Codes/stock_screenshot/1.png')
print('已擷取第 1 頁資料')

for i in range(2,int(TotalPage)//10 +2):
    driver.find_element_by_xpath('//*[@id="ngApp"]/div/div[2]/div[2]/ul/li[13]/a').click()
    time.sleep(10)
    driver.save_screenshot('/Users/kkday/Documents/Python_Codes/stock_screenshot/' + str(i) + '.png')
    print('已擷取第 ' + str(i) + ' 頁資料')

print('處理完畢！ 檔案已放置於 stock_screenshot 資料夾')
driver.quit()