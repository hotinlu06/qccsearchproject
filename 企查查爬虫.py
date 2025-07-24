import csv
import logging
import random
import time
import re

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from bs4 import BeautifulSoup as bs, BeautifulSoup # type: ignore
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

url = 'https://www.qcc.com/'

# 结果文件路径(需更改)
#result_file_path = r'data_result/data0717_01.csv'

desktop_path = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
result_file_path = os.path.join(desktop_path, 'data0717_01.csv')

# 参数为搜索关键词，调用时传入即可
def crawl_company_info(keyword):
    # 确保目录和文件存在
    directory = os.path.dirname(result_file_path)
    if not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)
    if not os.path.exists(result_file_path):
        with open(result_file_path, 'w', newline='', encoding='utf-8') as f:
            pass

    fields = [
         '企业名称', '注册资本', '国标行业', '统一社会信用代码', '负责人', '组织机构代码',
        '所属地区', '营业场所', '经营范围', '登记状态', '企业类型', '人员规模',
        '工商注册号', '营业期限', '参保人数', '登记机关', '英文名', '成立日期',
        '纳税人识别号', '纳税人资质', '核准日期',  '实缴资本',
        '法定代表人', '分支机构参保人数', '地址变更前', '进出口企业代码', '地址变更后'
    ]
    df = pd.DataFrame(columns=fields)
    pd.DataFrame(columns=fields).to_csv(result_file_path, index=False, encoding='utf-8')

    options = webdriver.ChromeOptions()
    # options.add_argument("--headless=new")
    # options.add_argument("--start-minimized")
    # service = Service(r'C:\Users\HTJ\Desktop\py_file\chrome_driver\chromedriver.exe')
    # driver = webdriver.Chrome(options=options, service=service)
    service = Service(executable_path="chromedriver")
    driver = webdriver.Chrome(options=options)
    driver.get(url)
    time.sleep(2)

    # 登录账号
    login_button = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.qccd-btn.qccd-btn-primary.qcc-header-login-btn"))
    )
    login_button.click()

    print("请30s内扫码登录")
    time.sleep(30)
    driver.refresh()


    # 关键词搜索
    try:
        # 输入关键词
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "searchKey"))
        )
        search_box.send_keys(keyword)
        time.sleep(2)

        # 点击搜索
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span.input-group-btn button"))
        )
        search_button.click()
        time.sleep(1)

        # 翻页遍历
        try:
            # 获取总页数
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//ul[@class='pagination']/li[last()-1]/a"))
            )
            text = element.text

            total_pages = int(re.search(r'\d+', text).group())
            print(f"一共 {total_pages} 页！")

            for page in range(1, total_pages + 1):
                print(f"正在遍历第 {page} 页")

                # 遍历公司信息
                try:
                    # 点击公司
                    company_elements = WebDriverWait(driver, 10).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.title.copy-value"))
                    )
                    for company in company_elements:
                        company_name = company.text
                        print(f"正在爬取公司: {company_name}")

                        # 当前窗口句柄
                        main_window = driver.current_window_handle

                        company.click()
                        time.sleep(random.uniform(2, 4))

                        # 等待新窗口加载
                        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

                        # 切换到新窗口
                        for window_handle in driver.window_handles:
                            if window_handle != main_window:
                                driver.switch_to.window(window_handle)
                                break

                        # 获取公司基础信息及地址变更信息
                        try:
                            # 基础信息
                            comp_info_parent = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "div.cominfo-normal")
                                )
                            )
                            comp_info_element = comp_info_parent.find_element(By.CSS_SELECTOR, "table.ntable")
                            comp_info = comp_info_element.get_attribute('outerHTML')

                            soup1 = BeautifulSoup(comp_info, 'html.parser')
                            table1 = soup1.find('table', class_='ntable')

                            # 将所有基础信息存进列表
                            comp_info_data = []
                            for row in table1.find_all('tr'):
                                row_data = []
                                for cell in row.find_all(['td', 'th']):
                                    row_data.append(cell.get_text(strip=True))
                                comp_info_data.append(row_data)

                            # 地址变更(可能不存在)
                            address_change_data = []
                            try:
                                address_change_element = WebDriverWait(driver, 3).until(
                                    EC.presence_of_element_located(
                                        (By.CSS_SELECTOR, 'table.ntable.app-ntable-expand-all.hide-info'))
                                )
                                address_change = address_change_element.get_attribute('outerHTML')

                                soup2 = BeautifulSoup(address_change, 'html.parser')
                                table2 = soup2.find('table', class_='ntable app-ntable-expand-all hide-info')

                                # 将所有地址变更信息存进列表
                                for row in table2.find_all('tr'):
                                    row_data = []
                                    for cell in row.find_all(['td']):
                                        row_data.append(cell.get_text(strip=True))
                                    address_change_data.append(row_data)

                            except Exception as e:
                                print(f"公司无变更信息！")

                            # 存放最终result
                            row_data = {field: "" for field in fields}

                            # 将基础信息结构化
                            for row in comp_info_data:
                                # 确保行数据至少包含2个元素（字段名和值）
                                if len(row) < 2:
                                    continue
                                # 偶数为字段名，奇数为值
                                for i in range(0, len(row), 2):
                                    # 确保值的索引存在
                                    if i + 1 < len(row):
                                        field_name = row[i].replace(' ', '')
                                        field_value = row[i + 1]

                                        # 仅处理目标字段列表中的字段，且值为空时才覆盖
                                        if field_name in fields and not row_data[field_name]:
                                            row_data[field_name] = field_value

                            # 将地址变更信息结构化
                            for row in address_change_data:
                                if len(row) >= 4 and '地址' in row[2]:
                                    address_old = row[3]
                                    address_new = row[4]

                                    # 最新的数据排在越前面，只需遍历第一个即可
                                    row_data['地址变更前'] = address_old
                                    row_data['地址变更后'] = address_new
                                    break

                            # 写入文件
                            pd.DataFrame([row_data]).to_csv(
                                result_file_path,
                                mode='a',
                                header=False,
                                index=False,
                                encoding='utf-8'
                            )

                            time.sleep(random.uniform(2, 4))

                        except Exception as e:
                            print(f"公司信息爬取错误：{str(e)}")
                            continue

                        finally:
                            # 关闭新窗口并切回主窗口
                            if driver.current_window_handle != main_window:
                                driver.close()
                            driver.switch_to.window(main_window)

                            time.sleep(random.uniform(2, 3))

                except Exception as e:
                    print(f"公司遍历错误：{str(e)}")
                    break

                # 如果不是最后一页，点击翻页
                if page < total_pages:
                    next_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), ">")]'))
                    )
                    next_button.click()
                    time.sleep(2)

        except Exception as e:
            print(f"翻页遍历错误：{str(e)}")

    except Exception as e:
        print(f"关键词搜索错误：{str(e)}")

    finally:
        driver.quit()



def csv_to_excel_with_highlight(csv_file_path, excel_file_path):
    try:
        # 读取 CSV 文件（可按需添加编码参数，如 encoding='utf-8'）
        df = pd.read_csv(csv_file_path)

        # 保存为 Excel 文件
        df.to_excel(excel_file_path, index=False, engine='openpyxl')

        # 加载 Excel 文件进行格式设置
        wb = openpyxl.load_workbook(excel_file_path)
        ws = wb.active

        # 定义颜色填充
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        # 获取列名所在的行号（通常是第一行）
        header_row = 1

        # 获取各列的列号，增加默认值避免后续报错
        capital_col = None
        employee_col = None
        address_change_col = None

        # 查找列名对应的列号（支持模糊匹配，如“注册资本”“参保人数”“地址变更后”）
        for col in range(1, ws.max_column + 1):
            col_name = ws.cell(row=header_row, column=col).value
            if col_name and re.search("注册资本", col_name):
                capital_col = col
            elif col_name and re.search("参保人数", col_name):
                employee_col = col
            elif col_name and re.search("地址变更后", col_name):
                address_change_col = col

        # 遍历每一行数据，从 header_row + 1 开始（跳过表头）
        for row in range(header_row + 1, ws.max_row + 1):
            # 标红逻辑：仅当“地址变更后”列存在时执行
            if address_change_col is not None:
                address_val = ws.cell(row=row, column=address_change_col).value
                # 清洗值：去除空白、判断非空
                clean_addr = str(address_val).strip() if pd.notna(address_val) else ""
                if clean_addr != "":
                    # 标红整行
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=row, column=col).fill = red_fill
                    # 标红后跳过标黄逻辑
                    continue

            # 标黄逻辑：需“注册资本”和“参保人数”列存在
            if capital_col is not None and employee_col is not None:
                # 处理注册资本
                capital_val = ws.cell(row=row, column=capital_col).value
                capital_over = False
                if pd.notna(capital_val):
                    # 清洗逻辑：提取“万元”数值（假设仅处理人民币万元，可扩展）
                    cap_match = re.search(r"(\d+(\.\d+)?)万元", str(capital_val))
                    if cap_match:
                        capital_num = float(cap_match.group(1)) * 10000  # 转换为元
                        capital_over = capital_num > 5000000  # 判断是否超500万（元）

                # 处理参保人数
                employee_val = ws.cell(row=row, column=employee_col).value
                employee_over = False
                if pd.notna(employee_val):
                    # 清洗逻辑：提取数字（如处理“5(2024年报)”“<null>”等）
                    emp_match = re.search(r"(\d+)", str(employee_val))
                    if emp_match:
                        employee_num = int(emp_match.group(1))
                        employee_over = employee_num > 30

                # 标黄条件：注册资本超500万 或 人数超30
                if capital_over or employee_over:
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=row, column=col).fill = yellow_fill

        # 保存修改后的 Excel 文件
        wb.save(excel_file_path)
        print(f"成功转换并标记颜色，文件已保存至: {excel_file_path}")

    except Exception as e:
        print(f"转换过程中出错: {str(e)}")



if __name__ == '__main__':

    crawl_company_info("汉唐大厦")

    csv_to_excel_with_highlight(
        csv_file_path=result_file_path,
        excel_file_path=r'data_result/data0717_02.xlsx'
    )