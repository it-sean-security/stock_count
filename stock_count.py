# coding=utf-8
import pygsheets as pyg
import requests, json, math, datetime
gc = pyg.authorize(service_file= '/Users/kkday/Documents/Python_Codes/My Project 87532-60e5410c48f8.json')
sh = gc.open('be2 test')
wks = sh.worksheet_by_title('工作表1')
url = 'https://be2.kkday.com/auth/login'
url2 = 'https://be2.kkday.com/warehouse/ajax_warehouse_item_location_list'

headers = {
'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
'accept-encoding': 'gzip, deflate, br',
'accept-language': 'zh-TW,zh;q=0.9,ko;q=0.8,en-US;q=0.7,en;q=0.6',
'cache-control': 'max-age=0',
'content-length': '62',
'content-type': 'application/x-www-form-urlencoded',
'origin': 'https://be2.kkday.com',
'referer': 'https://be2.kkday.com/auth/login',
'sec-fetch-dest': 'document',
'sec-fetch-mode': 'navigate',
'sec-fetch-site': 'same-origin',
'sec-fetch-user': '?1',
'upgrade-insecure-requests': '1',
'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36}'
}
login_data = {
'lang': 'zh-tw',
'username': 'alice.lee@kkday.com',
'password': 'charlie1111'
}
session = requests.session()
res = session.post(url, headers = headers, data = login_data)
if res == 200:
        print('登入成功！')
headers2 = {
'accept': 'application/json, text/plain, */*',
'accept-encoding': 'gzip, deflate, br',
'accept-language': 'zh-TW,zh;q=0.9,ko;q=0.8,en-US;q=0.7,en;q=0.6',
'content-length': '132',
'content-type': 'application/x-www-form-urlencoded',
'cookie': 'KKUD=dc35ae12c9116e60708e6dab180a4419; _gcl_au=1.1.974011066.1572830590; cto_lwid=1be195c0-31ed-4163-bbbe-ccffd4dd1341; _ga_P6E2T71LW0=GS1.1.1579676270.1.1.1579676297.0; _ga=GA1.2.1934077265.1572830591; mp_b8150a8ddf736c19fdc0f146b9ffac24_mixpanel=%7B%22distinct_id%22%3A%20%2216e3402372b889-0b4a5b94e039e1-1d3e6a5a-13c680-16e3402372cad6%22%2C%22%24device_id%22%3A%20%2216e3402372b889-0b4a5b94e039e1-1d3e6a5a-13c680-16e3402372cad6%22%2C%22Platform%22%3A%20%22www.kkday.com%22%2C%22LoginChannel%22%3A%20%22NO%22%2C%22DisplayCurrency%22%3A%20%22TWD%22%2C%22DisplayLang%22%3A%20%22zh-tw%22%2C%22DisplayCountry%22%3A%20%22TW%22%2C%22IsInternal%22%3A%20false%2C%22Cid%22%3A%20null%2C%22Ud1%22%3A%20null%2C%22Ud2%22%3A%20null%2C%22%24initial_referrer%22%3A%20%22%24direct%22%2C%22%24initial_referring_domain%22%3A%20%22%24direct%22%2C%22utm_source%22%3A%20%22facebook%22%2C%22utm_medium%22%3A%20%22cpc%22%2C%22utm_campaign%22%3A%20%22TW%22%7D; _hjid=5a34c676-6d22-49c9-9cc7-e1d525d744eb; lang_ui=zh-tw; '+ 
'be_ci_sessions=' + session.cookies.get_dict()['be_ci_sessions'] + ";"+ 
'kkday_be2_web_session=' + session.cookies.get_dict()['kkday_be2_web_session'] + ";"+ 
'authToken=' + session.cookies.get_dict()['authToken'],
'origin': 'https://be2.kkday.com',
'referer': 'https://be2.kkday.com/warehouse/warehouse_item_location_list',
'sec-fetch-dest': 'empty',
'sec-fetch-mode': 'cors',
'sec-fetch-site': 'same-origin',
'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
'x-xsrf-token': session.cookies.get_dict()['XSRF-TOKEN']
}
def datapage(page):
    data = {
    'currentPage': page,  #查詢頁
    'wh_item_oid':'',
    'wh_item_desc':'',
    'wh_upload_type':'',   #2-電子票
    'wh_location_oid':'',
    'is_delete': 'N',
    'sort_by_modify_date':'',
    'sort_by_wh_item_oid': 'ASC'
    }
    return data
#品項名稱 wh_item_desc
#倉位名稱 wh_location_desc
#出/入庫數量	wh_in_oid_count/wh_out_oid_count
#可使用票券數量/高風險票   remained_ticket_count/high_risk_ticket_count
#過期票/無檔案票  outdated_ticket_count/halted_ticket_count
#安全庫存   wh_safety_qty   
#異動人/異動時間 GMT+0  modify_user/modify_date
#是否作廢   is_delete
#allpage = math.ceil(count_total/10)
#res2 = session.post(url2,headers = headers2, data = datapage(1))

res2 = requests.post(url2,headers = headers2, data = datapage(1))
ticket_data = json.loads(res2.text)['data']
count_total = ticket_data[0]['count_total']
allpage = math.ceil(count_total/10)+1

#洗掉舊資料
wks.clear('A2','Z1000')
#取得第一頁資料
for item in range(10):
    wks.append_table(values=[ticket_data[item]['wh_item_desc'],ticket_data[item]['wh_item_oid'],ticket_data[item]['wh_location_desc'],ticket_data[item]['wh_in_oid_count'],ticket_data[item]['wh_out_oid_count'],ticket_data[item]['remained_ticket_count'],ticket_data[item]['high_risk_ticket_count'],ticket_data[item]['outdated_ticket_count'],ticket_data[item]['halted_ticket_count'],ticket_data[item]['wh_safety_qty'],ticket_data[item]['modify_user'],ticket_data[item]['modify_date'],ticket_data[item]['is_delete']])
    #取得後續幾頁的資料
for page in range(2,allpage):
    res2 = requests.post(url2,headers = headers2, data = datapage(page))
    ticket_data = json.loads(res2.text)['data']
    for item in range(len(ticket_data)):
        wks.append_table(values=[ticket_data[item]['wh_item_desc'],ticket_data[item]['wh_item_oid'],ticket_data[item]['wh_location_desc'],ticket_data[item]['wh_in_oid_count'],ticket_data[item]['wh_out_oid_count'],ticket_data[item]['remained_ticket_count'],ticket_data[item]['high_risk_ticket_count'],ticket_data[item]['outdated_ticket_count'],ticket_data[item]['halted_ticket_count'],ticket_data[item]['wh_safety_qty'],ticket_data[item]['modify_user'],ticket_data[item]['modify_date'],ticket_data[item]['is_delete']])

wks.update_value('N2', datetime.datetime.today().strftime("%Y/%m/%d %H:%M"))

print('資料彙整完成')