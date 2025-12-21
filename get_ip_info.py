# -*- coding: utf-8 -*-
"""
IP信息查询脚本
从Excel/CSV文件读取IP地址，查询iplark.com获取详细信息

使用说明：
1. 支持的文件格式：.csv, .xlsx, .xls
2. 修改下方配置区域的参数来适配你的文件

配置示例：
- IP_COLUMN = 'A'  表示读取A列
- IP_COLUMN = 'H'  表示读取H列（第8列）
- IP_COLUMN = 'ip' 表示自动查找列名包含'ip'的列
- IP_COLUMN = None 表示自动检测（查找包含'ip'的列名，否则使用最后一列）

输出文件：
1. ip_info_result_时间戳_UTC+X.xlsx - 纯查询结果
2. 原文件名_ip_info_result_时间戳_UTC+X.xlsx - 原表+查询结果
"""

import pandas as pd
import time
import re
import os
from datetime import datetime, timezone
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException


def setup_driver():
    """配置并启动Chrome浏览器"""
    chrome_driver_path = r'D:\Program\chromedriver-win64\chromedriver.exe'
    chrome_binary_path = r'D:\Program\chrome-win64\chrome.exe'

    options = Options()
    options.binary_location = chrome_binary_path
    # options.add_argument('--headless')  # 无头模式，取消注释可隐藏浏览器窗口
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(30)
    return driver


def extract_ip_from_hostname(hostname):
    """从主机名中提取IP地址，如果是域名则返回None"""
    hostname = hostname.strip()
    ip_pattern = r'^(\d{1,3}\.){3}\d{1,3}$'
    if re.match(ip_pattern, hostname):
        return hostname
    return None


def safe_find_text(driver, by, selector, default=''):
    """安全地查找元素并返回文本"""
    try:
        element = driver.find_element(by, selector)
        return element.text.strip()
    except NoSuchElementException:
        return default


def safe_find_texts(driver, by, selector):
    """安全地查找多个元素并返回文本列表"""
    try:
        elements = driver.find_elements(by, selector)
        return [e.text.strip() for e in elements]
    except:
        return []


def get_ip_info(driver, ip, retry_count=2):
    """查询单个IP的信息，带重试机制"""
    url = f"https://iplark.com/{ip}"
    result = {
        'IP': ip,
        '类型': '',
        'IP属性': '',
        '数字地址': '',
        '国家/地区': '',
        'ASN': '',
        '企业': '',
        'IP评分': '',
        '备注': '',
        '地理位置': '',
        '使用类型': '',
        '威胁': '',
        'IP类型': '',
        '提供商': '',
        '公共代理': '',
        '代理类型': '',
        '标签': '',
        '查询状态': ''
    }

    for attempt in range(retry_count + 1):
        try:
            driver.get(url)

            # 智能等待页面加载完成
            wait = WebDriverWait(driver, 30)

            # 1. 等待基础结构加载
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ip-card')))

            # 2. 等待关键数据元素出现（任意一个即可）
            try:
                wait.until(lambda d: (
                    d.find_elements(By.CSS_SELECTOR, '.ip-tags .tag') or
                    d.find_elements(By.CSS_SELECTOR, '.info-item .value') or
                    d.find_elements(By.ID, 'score-value')
                ))
            except:
                pass

            # 3. 等待页面完全加载（document.readyState）
            wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')

            # 4. 短暂等待动态内容渲染
            time.sleep(1.5)

            # 获取类型标签（家宽/机房等）
            tags = safe_find_texts(driver, By.CSS_SELECTOR, '.ip-tags .tag')
            if len(tags) >= 1:
                result['类型'] = tags[0]
            if len(tags) >= 2:
                result['IP属性'] = tags[1]

            # 获取基本信息项
            info_items = driver.find_elements(By.CSS_SELECTOR, '.info-item')
            for item in info_items:
                try:
                    label = item.find_element(By.TAG_NAME, 'label').text.strip()
                    value_elem = item.find_element(By.CSS_SELECTOR, '.value')
                    value = value_elem.text.strip()

                    if '数字地址' in label:
                        match = re.match(r'(\d+)', value)
                        if match:
                            result['数字地址'] = match.group(1)
                    elif '国家' in label or '地区' in label:
                        result['国家/地区'] = value
                    elif 'ASN' in label:
                        result['ASN'] = value
                    elif '企业' in label:
                        result['企业'] = value
                    elif '备注' in label:
                        result['备注'] = value
                except:
                    continue

            # 获取IP评分 - 多种方式尝试
            score = safe_find_text(driver, By.ID, 'score-value')
            if not score:
                score = safe_find_text(driver, By.CSS_SELECTOR, '.score-value')
            if not score:
                # 尝试从score-ratio获取
                ratio = safe_find_text(driver, By.ID, 'score-ratio')
                if ratio and '/' in ratio:
                    score = ratio.split('/')[0]
            result['IP评分'] = score

            # 获取地理位置（取第一个来源）
            location = safe_find_text(driver, By.ID, 'location-info1')
            if not location:
                geo_sources = driver.find_elements(By.CSS_SELECTOR, '.geo-source span:not(.source-tag)')
                if geo_sources:
                    location = geo_sources[0].text.strip()
            result['地理位置'] = location

            # 获取IP情报
            intel_section = safe_find_text(driver, By.ID, 'ip-intelligence')
            if intel_section:
                lines = intel_section.split('\n')
                current_key = ''
                for line in lines:
                    line = line.strip()
                    if '使用类型:' in line:
                        current_key = '使用类型'
                    elif '威胁:' in line:
                        current_key = '威胁'
                    elif 'IP类型:' in line:
                        current_key = 'IP类型'
                    elif '提供商:' in line:
                        current_key = '提供商'
                    elif '公共代理:' in line:
                        current_key = '公共代理'
                    elif '代理类型:' in line:
                        current_key = '代理类型'
                    elif '标签:' in line:
                        current_key = '标签'
                    elif current_key and line and line != '-':
                        if not result[current_key]:
                            result[current_key] = line

            result['查询状态'] = '成功'
            return result

        except TimeoutException:
            if attempt < retry_count:
                print(f"    超时，重试 {attempt + 1}/{retry_count}...")
                time.sleep(2)
                continue
            result['查询状态'] = '超时'
        except WebDriverException as e:
            if attempt < retry_count:
                print(f"    浏览器异常，重试 {attempt + 1}/{retry_count}...")
                time.sleep(2)
                continue
            result['查询状态'] = f'浏览器错误'
        except Exception as e:
            if attempt < retry_count:
                print(f"    错误，重试 {attempt + 1}/{retry_count}...")
                time.sleep(2)
                continue
            result['查询状态'] = f'错误: {str(e)[:30]}'

    return result


def get_timestamp_with_timezone():
    """
    获取当前时间戳和时区信息
    返回格式: (时间戳字符串, 时区字符串)
    例如: ('2025-12-21-231530', 'UTC+8')
    """
    now = datetime.now()

    # 获取时区偏移（秒）
    utc_offset_seconds = time.timezone if time.daylight == 0 else time.altzone
    utc_offset_hours = -utc_offset_seconds // 3600  # 注意符号相反

    # 格式化时间戳: yyyy-mm-dd-hhmmss
    timestamp = now.strftime('%Y-%m-%d-%H%M%S')

    # 格式化时区
    if utc_offset_hours >= 0:
        tz_str = f'UTC+{utc_offset_hours}'
    else:
        tz_str = f'UTC{utc_offset_hours}'

    return timestamp, tz_str


def column_letter_to_index(letter):
    """
    将Excel列字母转换为索引（从0开始）
    例如: 'A' -> 0, 'B' -> 1, 'Z' -> 25, 'AA' -> 26
    """
    letter = letter.upper()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1


def read_file_to_dataframe(file_path):
    """
    读取文件为DataFrame

    参数:
        file_path: 文件路径，支持 .csv, .xlsx, .xls 格式

    返回:
        (DataFrame, 成功标志)
    """
    try:
        file_ext = os.path.splitext(file_path)[1].lower()
        df = None

        if file_ext == '.csv':
            for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
        elif file_ext in ['.xlsx', '.xls']:
            try:
                df = pd.read_excel(file_path, engine='openpyxl' if file_ext == '.xlsx' else 'xlrd')
            except Exception:
                df = pd.read_excel(file_path)
        else:
            print(f"不支持的文件格式: {file_ext}")
            print("支持的格式: .csv, .xlsx, .xls")
            return None, False

        if df is None or df.empty:
            print("无法读取文件或文件为空")
            return None, False

        return df, True
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None, False


def get_ip_column_name(df, ip_column=None):
    """
    确定IP所在列的列名

    参数:
        df: DataFrame
        ip_column: 用户指定的列

    返回:
        列名 或 None
    """
    if ip_column is None:
        for col in df.columns:
            if 'ip' in str(col).lower():
                return col
        return df.columns[-2] if len(df.columns) >= 2 else df.columns[-1]
    elif len(ip_column) <= 2 and ip_column.isalpha():
        col_index = column_letter_to_index(ip_column)
        if col_index < len(df.columns):
            return df.columns[col_index]
        else:
            print(f"列 {ip_column} 超出范围，文件只有 {len(df.columns)} 列")
            return None
    else:
        if ip_column in df.columns:
            return ip_column
        for col in df.columns:
            if ip_column.lower() in str(col).lower():
                return col
        print(f"未找到列: {ip_column}")
        print(f"可用的列: {list(df.columns)}")
        return None


def extract_ips_from_column(df, column_name):
    """
    从指定列提取IP地址

    返回:
        (IP列表, IP到行索引的映射字典)
    """
    ips = []
    ip_to_rows = {}  # IP -> 行索引列表

    for idx, value in df[column_name].items():
        if pd.isna(value):
            continue
        value = str(value).strip().replace('\t', '')
        ip = extract_ip_from_hostname(value)
        if ip:
            if ip not in ip_to_rows:
                ip_to_rows[ip] = []
                ips.append(ip)
            ip_to_rows[ip].append(idx)

    return ips, ip_to_rows


def main():
    # ==================== 用户配置区域 ====================
    # 输入文件路径（支持 .csv, .xlsx, .xls 格式）
    INPUT_FILE = r'E:\AAAAAcodedata\getiplarkusage\testIP.xlsx'

    # 输出目录（留空则与输入文件同目录）
    OUTPUT_DIR = r''

    # IP地址所在列配置
    # 可选值:
    #   - None     : 自动检测（查找列名包含'ip'的列）
    #   - 'A'      : 使用A列（第1列）
    #   - 'B'      : 使用B列（第2列）
    #   - 'H'      : 使用H列（第8列）
    #   - 'AA'     : 使用AA列（第27列）
    #   - '登录ip' : 使用名为"登录ip"的列
    IP_COLUMN = None  # <-- 修改这里来指定IP所在列
    # ======================================================

    # 获取查询开始时间戳
    timestamp, tz_str = get_timestamp_with_timezone()
    time_suffix = f"_{timestamp}_{tz_str}"

    print("=" * 60)
    print("IP信息查询工具 - iplark.com")
    print(f"查询时间: {timestamp} ({tz_str})")
    print("=" * 60)

    # 确定输出目录和文件名
    input_dir = os.path.dirname(INPUT_FILE)
    input_basename = os.path.splitext(os.path.basename(INPUT_FILE))[0]
    output_dir = OUTPUT_DIR if OUTPUT_DIR else input_dir

    # 两个输出文件路径
    output_file_1 = os.path.join(output_dir, f"ip_info_result{time_suffix}.xlsx")
    output_file_2 = os.path.join(output_dir, f"{input_basename}_ip_info_result{time_suffix}.xlsx")

    print(f"\n[1/4] 正在读取文件: {INPUT_FILE}")

    # 读取原始DataFrame
    df_original, success = read_file_to_dataframe(INPUT_FILE)
    if not success:
        return

    # 确定IP列
    ip_column_name = get_ip_column_name(df_original, IP_COLUMN)
    if ip_column_name is None:
        return

    print(f"使用列: {ip_column_name} (索引: {list(df_original.columns).index(ip_column_name)})")

    # 提取IP地址
    ips, ip_to_rows = extract_ips_from_column(df_original, ip_column_name)

    if not ips:
        print("未找到有效的IP地址！")
        return

    print(f"找到 {len(ips)} 个唯一IP地址:")
    for ip in ips:
        print(f"  - {ip}")

    print("\n[2/4] 正在启动浏览器...")
    driver = setup_driver()

    print("\n[3/4] 正在查询IP信息...")
    results = []
    ip_to_result = {}  # IP -> 查询结果

    try:
        for i, ip in enumerate(ips, 1):
            print(f"  查询 ({i}/{len(ips)}): {ip}")
            info = get_ip_info(driver, ip)
            results.append(info)
            ip_to_result[ip] = info

            status_icon = "OK" if info['查询状态'] == '成功' else "X"
            print(f"    [{status_icon}] {info['类型']} | {info['国家/地区']} | 评分:{info['IP评分']}")

            if i < len(ips):
                time.sleep(2)
    finally:
        driver.quit()

    print("\n[4/4] 正在保存结果...")

    # ===== 文件1: 纯查询结果 =====
    df_result = pd.DataFrame(results)
    columns_order = ['IP', '类型', 'IP属性', '国家/地区', '地理位置', 'ASN', '企业',
                     'IP评分', '使用类型', 'IP类型', '公共代理', '威胁', '代理类型',
                     '标签', '数字地址', '备注', '查询状态']
    columns_order = [c for c in columns_order if c in df_result.columns]
    df_result = df_result[columns_order]
    df_result.to_excel(output_file_1, index=False, engine='openpyxl')
    print(f"  文件1: {output_file_1}")

    # ===== 文件2: 原表 + 查询结果 =====
    # 要追加的列（不包含IP列，因为原表已有）
    append_columns = ['查询_类型', '查询_IP属性', '查询_国家地区', '查询_地理位置',
                      '查询_ASN', '查询_企业', '查询_IP评分', '查询_使用类型',
                      '查询_IP类型', '查询_公共代理', '查询_状态']
    result_keys = ['类型', 'IP属性', '国家/地区', '地理位置', 'ASN', '企业',
                   'IP评分', '使用类型', 'IP类型', '公共代理', '查询状态']

    # 初始化新列
    for col in append_columns:
        df_original[col] = ''

    # 填充查询结果到对应行
    for ip, row_indices in ip_to_rows.items():
        if ip in ip_to_result:
            result = ip_to_result[ip]
            for row_idx in row_indices:
                for col_name, key in zip(append_columns, result_keys):
                    df_original.at[row_idx, col_name] = result.get(key, '')

    df_original.to_excel(output_file_2, index=False, engine='openpyxl')
    print(f"  文件2: {output_file_2}")

    success_count = sum(1 for r in results if r['查询状态'] == '成功')
    print(f"\n查询完成: 成功 {success_count}/{len(ips)}")


if __name__ == '__main__':
    main()
