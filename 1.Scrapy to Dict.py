from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
from pymongo import MongoClient


browser = webdriver.Chrome()
wait = WebDriverWait(browser, 5)


tech_list = []

# 1.获取全部科技类别
def search_techs():
    browser.get('https://www.mor-research.com/technologies/?typesList=&invstatusList=&medfieldsList=')
    num = 0
    while True:
        try:
            num = num + 1
            tech_dict = {}
            title = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#techResults > ul > li:nth-child(%d) > h3 > a' % num)))
            summary = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#techResults > ul > li:nth-child(%d) > p:nth-child(2)' % num)))
            tech_item = title.text
            tech_url = title.get_attribute('href')
            tech_summ = summary.text

            # 获取当前主页面句柄
            main_handle = browser.current_window_handle
            js = 'window.open("");'
            browser.execute_script(js)
            # 获取当前窗口句柄集合（列表类型）
            handles = browser.window_handles
            item_handle = None
            for handle in handles:
                if handle != main_handle:
                    item_handle = handle
            browser.switch_to.window(item_handle)

            item_dict = search_tech_items(tech_url)
            tech_dict['Title'] = tech_item
            tech_dict['URL'] = tech_url
            tech_dict['Summary'] = tech_summ

            tech_dict.update(item_dict)
            tech_list.append(tech_dict)

            #print('tech_dict', tech_dict)

            browser.close()
            browser.switch_to.window(main_handle)
            print(num,"\n",tech_dict)
        except:
            break

    return tech_list




#2.获取各个科技类别的详细信息
def search_tech_items(tech_url):
    browser.get(tech_url)
    i = 0
    n = 2
    t = 0
    dict = {}
    content = ""
    while True:
        try:
            i = i + 1
            try:
                key = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#main-section > div.mainDiv.pageContent > div.leftColumn > ul > li:nth-child(%d) > b' % i)))
            except:
                key = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#main-section > div.mainDiv.pageContent > div.leftColumn > ul > li:nth-child(%d) > p > b' % i)))
            value = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#main-section > div.mainDiv.pageContent > div.leftColumn > ul > li:nth-child(%d)' % i)))
            key_text = key.text.replace(':','').strip()
            value_text = value.text.replace(key.text,'').strip()
            dict[key_text] = value_text
        except:
            break


    while True:
        try:
            n = n + 1
            contarea = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#main-section > div.mainDiv.pageContent > div.leftColumn > p:nth-child(%d)' % n)))
            content = content + contarea.text + '\n'
        except:
            break

    while True:
        try:
            t = t + 1
            contarea = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#main-section > div.mainDiv.pageContent > div.leftColumn > ul:nth-child(6) > li:nth-child(%d)' % t)))
            content = content + contarea.text
        except:
            dict['Content'] = content
            break

    return dict


def save_to_MongoDB(list,dbname,colname):
    conn = MongoClient('127.0.0.1', 27017)
    db = conn[dbname]
    collist = db.list_collection_names()
    #print(collist)
    tech_mgdb = db[colname]
    if colname in collist: tech_mgdb.drop()
    tech_mgdb.insert_many(list)


if __name__ == '__main__':
    tech_list = search_techs()

    dbname = 'Technologies爬取库'
    colname = '科技类别信息_201909'
    save_to_MongoDB(tech_list,dbname,colname)



