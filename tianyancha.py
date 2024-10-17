import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import user_set as user
from time import sleep
import random
from pymongo import MongoClient




def connect_mongodb():
    """连接到MongoDB数据库"""
    client = MongoClient('mongodb://localhost:27017/')  # 修改为你的MongoDB连接字符串
    db = client['company_data']  # 数据库名称
    collection = db['registered_capital']  # 集合名称
    return collection

def save_to_mongodb(collection, data):
    """将查询数据保存到MongoDB"""
    collection.update_one(
        {'公司名称': data['公司名称']},  # 以公司名称为唯一键，避免重复插入
        {'$set': data},  # 更新数据
        upsert=True  # 如果不存在则插入
    )
def start_browser():
    """启动浏览器并返回驱动对象"""

    # headers = random.choice(user.UserAgent1)
    account = random.choice(user.Account)
    chrome_options = Options()
    chrome_options.add_argument('--disable-gpu')
    # chrome_options.add_argument('headless')
    chrome_options.add_argument('--no-sandbox')
    # chrome_options.add_argument('--user-agent=%s' % headers)
    chrome_options.add_argument('--ignore-certificate-errors')
    driver = webdriver.Chrome(options=chrome_options)
    # 登录天眼查
    driver.get('https://www.tianyancha.com/login?')
    # driver.get('https://www.baidu.com')
    print("正在登录...")
    print(account)
    WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.XPATH,
                                        '//*[@id="web-content"]/div/div/div/div/div[2]'))
    )
    driver.find_element(By.XPATH,'//*[@id="web-content"]/div/div/div/div/div[2]').click()
    driver.find_element(By.XPATH, '//*[@id="web-content"]/div/div/div/div/div[6]/div/div[1]/div[2]').click()
    driver.find_element(By.XPATH, '//*[@id="agreement-checkbox-account"]').click()
    driver.find_element(By.XPATH, '//*[@id="mobile"]').send_keys(account[0])
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(account[1], Keys.ENTER)

    WebDriverWait(driver, 60).until(EC.url_changes('https://www.tianyancha.com/login?'))  # 等待登录完成
    # sleep(10)

    return driver


def fetch_registered_capital(excel_file):
    db = connect_mongodb()  # 连接MongoDB

    driver = start_browser()  # 初始化selenium

    # 读取Excel文件
    xls = pd.ExcelFile(excel_file)
    total_count = 0  # 全局计数器

    districts = ['黄浦区', '徐汇区', '长宁区', '静安区', '普陀区', '虹口区', '杨浦区', '闵行区', '宝山区', '嘉定区','浦东新区', '金山区', '松江区', '青浦区', '奉贤区', '崇明区']

    for sheet_name in districts:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        企业名称列 = '公司名称'  # 使用新的列名
        所属领域列 = '所属领域'
        # 获取该sheet对应的MongoDB集合（根据sheet_name生成集合名称）
        collection = db[sheet_name]


        # 搜索每个企业
        for index, row in df.iterrows():
            if total_count >= 3146:  # 可以更改该条件以跳过指定数量的企业
                企业名称 = row[企业名称列]
                所属领域 = row[所属领域列]
                search_url = f'https://www.tianyancha.com/search?key={企业名称}'
                driver.get(search_url)

                # 等待注册资本加载
                try:
                    reg_cap = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '//*[@id="page-container"]/div/div[2]/div/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div[2]/span'))
                    ).text
                    df.at[index, '注册资本'] = reg_cap
                    total_count += 1  # 查询到一个企业，计数+1
                    print(f"已查询到{sheet_name} 共{total_count} 个企业：{企业名称}, 注册资本: {reg_cap}")

                    # 保存结果到MongoDB的该集合中
                    data = {'所属领域':所属领域,'公司名称': 企业名称, '注册资本': reg_cap}
                    save_to_mongodb(collection, data)

                except Exception as e:

                    try:
                        search_url = f'https://www.tianyancha.com/search?key={企业名称}'
                        driver.get(search_url)
                        reg_cap = WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            '//*[@id="page-container"]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div[2]/span'))
                        ).text
                        df.at[index, '注册资本'] = reg_cap
                        total_count += 1  # 查询到一个企业，计数+1
                        print(f"已查询到 {total_count} 个企业：{企业名称}, 注册资本: {reg_cap}")

                        # 保存结果到MongoDB的该集合中
                        data = {'所属领域':所属领域,'公司名称': 企业名称, '注册资本': reg_cap}
                        save_to_mongodb(collection, data)
                    except Exception as e:
                        print(f"未找到 {企业名称} 的注册资本，重新来")
                        driver.quit()

                        driver = start_browser()
                        search_url = f'https://www.tianyancha.com/search?key={企业名称}'
                        driver.get(search_url)
                        sleep(5)
                        reg_cap = WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            '//*[@id="page-container"]/div/div[2]/div/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div[2]/span'))
                        ).text
                        df.at[index, '注册资本'] = reg_cap
                        total_count += 1  # 查询到一个企业，计数+1
                        print(f"已查询到 {total_count} 个企业：{企业名称}, 注册资本: {reg_cap}")

                        # 保存结果到MongoDB的该集合中
                        data = {'所属领域': 所属领域, '公司名称': 企业名称, '注册资本': reg_cap}
                        save_to_mongodb(collection, data)



            else:
                total_count += 1  # 已经查询过的企业，计数+1
                print(total_count)

    # 关闭浏览器
    driver.quit()
    print(f"共查询到 {total_count} 个企业的注册资本。")

if __name__ == '__main__':
    # start_browser()
    # 调用函数
    fetch_registered_capital('company.xlsx')  # 修改为你的文件名