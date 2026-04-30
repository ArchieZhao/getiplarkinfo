# -*- coding: utf-8 -*-
"""
IP信息查询脚本
从Excel/CSV文件读取IP地址，查询iplark.com获取详细信息

使用说明：
1. 支持的文件格式：.csv, .xlsx, .xls
2. Excel文件会自动读取全部工作表，不要求工作表名为Sheet1/Sheet2
3. 修改下方配置区域的参数来适配你的文件

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

GEO_SOURCES = [
    'Ip-api', 'Moe', 'Moe+', 'Ease', 'Internet', 'Maxmind', 'Ipstack',
    'IPinfo', 'IP2Location', 'Digital Element', 'DB-IP',
    'Aliyun', 'TencentCloud', 'Cloudflare',
]


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
        '使用场景': '',
        'IP评分': '',
        '备注': '',
        '使用类型': '',
        '威胁': '',
        'IP类型': '',
        '提供商': '',
        '公共代理': '',
        '代理类型': '',
        '标签': '',
        'IP情报-使用类型': '',
        'IP情报-威胁': '',
        'IP情报-IP类型': '',
        'IP情报-提供商': '',
        'IP情报-公共代理': '',
        'IP情报-代理类型': '',
        'IP情报-标签': '',
        '查询状态': ''
    }
    for src in GEO_SOURCES:
        result[f'地理位置-{src}'] = ''

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
                        # 先尝试点击“小眼睛”显示完整数字地址，再提取纯数字
                        numeric_text = value
                        try:
                            eye_icons = value_elem.find_elements(
                                By.CSS_SELECTOR,
                                'span.js-tool-remove[title*="显示"], span.js-tool-remove[title*="点击显示"], span.js-tool-remove[title*="IP"]'
                            )
                            if eye_icons and '*' in numeric_text:
                                try:
                                    eye_icons[0].click()
                                except Exception:
                                    driver.execute_script("arguments[0].click();", eye_icons[0])

                                try:
                                    WebDriverWait(driver, 5).until(
                                        lambda d: '*' not in value_elem.text
                                    )
                                except Exception:
                                    pass

                                numeric_text = value_elem.text.strip()
                        except Exception:
                            numeric_text = value

                        digits_only = ''.join(re.findall(r'\d+', numeric_text))
                        if digits_only:
                            result['数字地址'] = digits_only
                        else:
                            first_token = numeric_text.split()[0] if numeric_text else ''
                            result['数字地址'] = first_token
                    elif '国家' in label or '地区' in label:
                        # 尝试把国旗 alt（如 China）与文本（如 中国）拼接为 China中国
                        country_text = value
                        try:
                            flag_imgs = value_elem.find_elements(By.TAG_NAME, 'img')
                            if flag_imgs:
                                alt = (flag_imgs[0].get_attribute('alt') or '').strip()
                                text_only = (value_elem.text or '').strip()
                                if alt and text_only:
                                    country_text = text_only if alt in text_only else f"{alt}{text_only}"
                                elif alt:
                                    country_text = alt
                        except Exception:
                            pass
                        result['国家/地区'] = country_text
                    elif 'ASN' in label:
                        result['ASN'] = value
                    elif '企业' in label:
                        result['企业'] = value
                    elif '使用场景' in label:
                        result['使用场景'] = value
                        if not result['使用类型']:
                            result['使用类型'] = value
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

            # 获取地理位置（多源对比）
            geo_source_divs = driver.find_elements(By.CSS_SELECTOR, '.geo-source')
            for geo_div in geo_source_divs:
                try:
                    source_tag = geo_div.find_element(By.CSS_SELECTOR, '.source-tag').text.strip()
                    value_spans = geo_div.find_elements(By.CSS_SELECTOR, 'span:not(.source-tag)')
                    geo_text = ' '.join(s.text.strip() for s in value_spans if s.text.strip())
                    key = f'地理位置-{source_tag}'
                    if key in result:
                        result[key] = geo_text
                except:
                    continue

            # 获取IP情报
            try:
                intel_elem = driver.find_element(By.ID, 'ip-intelligence')
                span_elems = intel_elem.find_elements(By.CSS_SELECTOR, 'span')

                label_to_key = {
                    '使用类型': 'IP情报-使用类型',
                    '威胁': 'IP情报-威胁',
                    'IP类型': 'IP情报-IP类型',
                    '提供商': 'IP情报-提供商',
                    '公共代理': 'IP情报-公共代理',
                    '代理类型': 'IP情报-代理类型',
                    '标签': 'IP情报-标签',
                }
                legacy_keys = {'使用类型', '威胁', 'IP类型', '提供商', '公共代理', '代理类型', '标签'}

                i = 0
                while i < len(span_elems):
                    strongs = span_elems[i].find_elements(By.TAG_NAME, 'strong')
                    if not strongs:
                        i += 1
                        continue

                    raw_label = strongs[0].text.strip().rstrip(':：')
                    key = label_to_key.get(raw_label)

                    value = ''
                    if i + 1 < len(span_elems):
                        next_has_strong = span_elems[i + 1].find_elements(By.TAG_NAME, 'strong')
                        if not next_has_strong:
                            value = span_elems[i + 1].text.strip()
                            i += 1

                    # IP情报字段需要保留“-”等占位值，因此不跳过 value == '-'
                    if key and value and not result.get(key):
                        result[key] = value

                    # 兼容原有字段：仅在有实际值（非 '-'）时回填
                    if raw_label in legacy_keys and value and value != '-' and not result.get(raw_label):
                        result[raw_label] = value
                        if raw_label == '使用类型' and not result.get('使用场景'):
                            result['使用场景'] = value

                    i += 1
            except NoSuchElementException:
                pass

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


def read_file_to_dataframes(file_path):
    """
    读取文件为工作表DataFrame字典

    参数:
        file_path: 文件路径，支持 .csv, .xlsx, .xls 格式

    返回:
        ({工作表名: DataFrame}, 成功标志)
    """
    try:
        file_ext = os.path.splitext(file_path)[1].lower()
        sheets = None

        if file_ext == '.csv':
            df = None
            for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            if df is not None:
                sheets = {'原始数据': df}
        elif file_ext in ['.xlsx', '.xls']:
            try:
                sheets = pd.read_excel(
                    file_path,
                    sheet_name=None,
                    engine='openpyxl' if file_ext == '.xlsx' else 'xlrd'
                )
            except Exception:
                sheets = pd.read_excel(file_path, sheet_name=None)
        else:
            print(f"不支持的文件格式: {file_ext}")
            print("支持的格式: .csv, .xlsx, .xls")
            return None, False

        if not sheets:
            print("无法读取文件或文件为空")
            return None, False

        has_data_sheet = any(df is not None and not df.empty for df in sheets.values())
        if not has_data_sheet:
            print("无法读取文件或文件为空")
            return None, False

        return sheets, True
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None, False


def read_file_to_dataframe(file_path):
    """
    兼容旧调用：读取第一个工作表为DataFrame

    返回:
        (DataFrame, 成功标志)
    """
    sheets, success = read_file_to_dataframes(file_path)
    if not success:
        return None, False
    first_df = next(iter(sheets.values()))
    return first_df, True


def get_ip_column_name(df, ip_column=None):
    """
    确定IP所在列的列名

    参数:
        df: DataFrame
        ip_column: 用户指定的列

    返回:
        列名 或 None
    """
    if df is None or len(df.columns) == 0:
        print("当前工作表没有可用的列")
        return None

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


def collect_ips_from_sheets(original_sheets, ip_column=None):
    """
    从全部工作表提取唯一IP和回填位置

    返回:
        (唯一IP列表, IP到工作表行索引的映射)
    """
    ips = []
    ip_to_rows = {}  # IP -> [(工作表名, 行索引)]

    for sheet_name, df_original in original_sheets.items():
        if df_original is None or df_original.empty:
            print(f"工作表 [{sheet_name}] 为空，跳过IP提取")
            continue

        ip_column_name = get_ip_column_name(df_original, ip_column)
        if ip_column_name is None:
            print(f"工作表 [{sheet_name}] 未找到IP列，跳过")
            continue

        print(f"工作表 [{sheet_name}] 使用列: {ip_column_name} (索引: {list(df_original.columns).index(ip_column_name)})")

        sheet_ips, sheet_ip_to_rows = extract_ips_from_column(df_original, ip_column_name)
        if not sheet_ips:
            print(f"工作表 [{sheet_name}] 未找到有效IP")
            continue

        print(f"工作表 [{sheet_name}] 找到 {len(sheet_ips)} 个唯一IP")
        for ip in sheet_ips:
            if ip not in ip_to_rows:
                ip_to_rows[ip] = []
                ips.append(ip)
            ip_to_rows[ip].extend((sheet_name, row_idx) for row_idx in sheet_ip_to_rows[ip])

    return ips, ip_to_rows


def main():
    # ==================== 用户配置区域 ====================
    # 输入文件路径（支持 .csv, .xlsx, .xls 格式）
    INPUT_FILE = r'E:\AAAAAcodedata\getiplarkinfo\testip.xlsx'

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

    # 读取原始工作表。Excel会读取全部sheet；CSV按单个工作表处理。
    original_sheets, success = read_file_to_dataframes(INPUT_FILE)
    if not success:
        return

    print(f"读取到 {len(original_sheets)} 个工作表:")
    for sheet_name in original_sheets:
        print(f"  - {sheet_name}")

    # 确定每个工作表的IP列，并汇总全部工作表中的唯一IP
    ips, ip_to_rows = collect_ips_from_sheets(original_sheets, IP_COLUMN)

    if not ips:
        print("未找到有效的IP地址！")
        return

    print(f"全部工作表共找到 {len(ips)} 个唯一IP地址:")
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
    geo_columns = [f'地理位置-{src}' for src in GEO_SOURCES]
    columns_order = ['IP', '类型', 'IP属性', '国家/地区'] + geo_columns + [
                     'ASN', '企业', '使用场景', 'IP评分', 'IP情报-使用类型',
                     'IP情报-威胁', 'IP情报-IP类型', 'IP情报-提供商',
                     'IP情报-公共代理', 'IP情报-代理类型', 'IP情报-标签',
                     '数字地址', '备注', '查询状态']
    columns_order = [c for c in columns_order if c in df_result.columns]
    df_result = df_result[columns_order]
    df_result.to_excel(output_file_1, index=False, engine='openpyxl')
    print(f"  文件1: {output_file_1}")

    # ===== 文件2: 原表 + 查询结果 =====
    # 要追加的列（不包含IP列，因为原表已有）
    geo_append_cols = [f'查询_地理位置-{src}' for src in GEO_SOURCES]
    geo_result_keys = [f'地理位置-{src}' for src in GEO_SOURCES]
    append_columns = ['查询_类型', '查询_使用场景', '查询_IP属性', '查询_国家地区'
                      ] + geo_append_cols + [
                      '查询_ASN', '查询_企业', '查询_IP评分',
                      '查询_数字地址', '查询_备注', 'IP情报-使用类型', 'IP情报-威胁',
                      'IP情报-IP类型', 'IP情报-提供商', 'IP情报-公共代理',
                      'IP情报-代理类型', 'IP情报-标签', '查询_状态']
    result_keys = ['类型', '使用场景', 'IP属性', '国家/地区'
                   ] + geo_result_keys + [
                   'ASN', '企业', 'IP评分', '数字地址', '备注', 'IP情报-使用类型',
                   'IP情报-威胁', 'IP情报-IP类型', 'IP情报-提供商', 'IP情报-公共代理',
                   'IP情报-代理类型', 'IP情报-标签', '查询状态']

    # 初始化每个工作表的新列
    for df_original in original_sheets.values():
        if df_original is None:
            continue
        for col in append_columns:
            df_original[col] = ''

    # 填充查询结果到对应工作表的对应行
    for ip, row_indices in ip_to_rows.items():
        if ip in ip_to_result:
            result = ip_to_result[ip]
            for sheet_name, row_idx in row_indices:
                df_original = original_sheets[sheet_name]
                for col_name, key in zip(append_columns, result_keys):
                    df_original.at[row_idx, col_name] = result.get(key, '')

    with pd.ExcelWriter(output_file_2, engine='openpyxl') as writer:
        for sheet_name, df_original in original_sheets.items():
            df_original.to_excel(writer, index=False, sheet_name=sheet_name)
    print(f"  文件2: {output_file_2}")

    success_count = sum(1 for r in results if r['查询状态'] == '成功')
    print(f"\n查询完成: 成功 {success_count}/{len(ips)}")


if __name__ == '__main__':
    main()
