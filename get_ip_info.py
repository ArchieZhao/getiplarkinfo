# -*- coding: utf-8 -*-
"""
IP信息查询脚本
从Excel/CSV文件读取IP地址，查询iplark.com获取详细信息

使用说明：
1. 支持的文件格式：.csv, .xlsx, .xls
2. Excel文件会自动读取全部工作表，不要求工作表名为Sheet1/Sheet2
3. 地理位置支持14个数据源多源对比采集
4. 可通过 -i/--input 传入文件路径，或通过 -ip 直接传入一个或多个 IP

配置示例：
- IP_COLUMN = 'A'  表示读取A列
- IP_COLUMN = 'H'  表示读取H列（第8列）
- IP_COLUMN = 'ip' 表示指定查找列名包含'ip'的列
- IP_COLUMN = None 表示自动检测（扫描各列内容，提取公网 IPv4 地址）

输出文件：
1. ip_info_result_时间戳_UTC+X.xlsx - 纯查询结果
2. 原文件名_ip_info_result_时间戳_UTC+X.xlsx - 原表全部工作表+查询结果
"""

import argparse
import ipaddress
import os
import re
import time
from datetime import datetime

import pandas as pd
from openpyxl.utils import get_column_letter
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
GEO_RESULT_KEYS = [f'地理位置-{source}' for source in GEO_SOURCES]
GEO_SOURCE_LOOKUP = {source.casefold(): source for source in GEO_SOURCES}

INTEL_FIELD_MAPPINGS = [
    ('使用类型', 'IP情报-使用类型'),
    ('威胁', 'IP情报-威胁'),
    ('IP类型', 'IP情报-IP类型'),
    ('提供商', 'IP情报-提供商'),
    ('公共代理', 'IP情报-公共代理'),
    ('代理类型', 'IP情报-代理类型'),
    ('标签', 'IP情报-标签'),
]
INTEL_LABEL_ALIASES = {
    '代理': '公共代理',
}
LEGACY_INTEL_KEYS = [label for label, _ in INTEL_FIELD_MAPPINGS]
INTEL_RESULT_KEYS = [key for _, key in INTEL_FIELD_MAPPINGS]
INTEL_RESULT_KEY_BY_LABEL = dict(INTEL_FIELD_MAPPINGS)
for alias, canonical_label in INTEL_LABEL_ALIASES.items():
    INTEL_RESULT_KEY_BY_LABEL[alias] = INTEL_RESULT_KEY_BY_LABEL[canonical_label]

BASE_RESULT_KEYS = [
    'IP', '页面顶部标签', '反查域名', '类型', 'IP属性', '数字地址', '国家/地区',
    'ASN', '企业', '使用场景', 'IP评分', '备注', '查询状态',
]

RESULT_FIELD_DESCRIPTIONS = {
    'IP': 'IP地址',
    '页面顶部标签': '页面 IP 标题下方的子标签，按页面顺序用分号连接',
    '反查域名': '页面顶部子标签中看起来像 DNS 反查主机名的值',
    '类型': '家宽、数据中心、商宽等',
    'IP属性': '原生IP、广播IP',
    '国家/地区': '所属国家或地区',
    'ASN': '自治系统编号',
    '企业': '所属企业/运营商',
    '使用场景': '网页“使用场景”字段（例如：普通宽带）',
    'IP评分': '0-100分',
    'IP情报-使用类型': 'IP情报区域中的“使用类型”（保留 -）',
    'IP情报-威胁': 'IP情报区域中的“威胁”（保留 -）',
    'IP情报-IP类型': 'IP情报区域中的“IP类型”（保留 -）',
    'IP情报-提供商': 'IP情报区域中的“提供商”（保留 -）',
    'IP情报-公共代理': 'IP情报区域中的“公共代理”（保留 -）',
    'IP情报-代理类型': 'IP情报区域中的“代理类型”（保留 -）',
    'IP情报-标签': 'IP情报区域中的“标签”（保留 -）',
    '数字地址': '会先点击小眼睛显示完整数字，再提取纯数字',
    '备注': 'ASN 规模等补充说明',
    '查询状态': '成功/超时/错误',
}

QUERY_IP_APPEND_COLUMN = '查询IP'
EXCEL_TEXT_NUMBER_FORMAT = '@'
LONG_NUMERIC_TEXT_MIN_DIGITS = 11
TEXT_PRESERVE_EXACT_COLUMNS = {
    'IP',
    QUERY_IP_APPEND_COLUMN,
    '数字地址',
}
TEXT_PRESERVE_COLUMN_KEYWORDS = [
    '证件号',
    '证件号码',
    '身份证',
    '身份证号',
    'QQ',
    '用户ID',
    '用户 ID',
    '支付平台用户ID',
    '支付平台用户 ID',
    'UID',
    'UserID',
    'User ID',
    'OpenID',
    'UnionID',
    '账号',
    '帐号',
    '账户',
    '手机号',
    '手机号码',
    '电话号码',
    '银行卡',
    '卡号',
    '订单号',
    '流水号',
    '交易号',
    '编号',
]


def normalize_label_text(text):
    """
    规范化页面标签文本，去除多余空白和中英文冒号。

    返回:
        规范化后的标签字符串
    """
    return re.sub(r'\s+', ' ', str(text or '')).strip().rstrip(':：')


def normalize_geo_source_name(source_name):
    """
    将页面中的地理位置数据源名称映射为标准名称。

    返回:
        GEO_SOURCES 中的标准名称；未知来源返回页面原始名称，避免丢数据。
    """
    normalized = normalize_label_text(source_name)
    if not normalized:
        return None
    return GEO_SOURCE_LOOKUP.get(normalized.casefold(), normalized)


def collect_geo_result_keys(results=None):
    """
    汇总地理位置结果列。

    返回:
        先包含 14 个标准数据源，再追加页面实际出现的新增数据源。
    """
    geo_keys = list(GEO_RESULT_KEYS)
    if results:
        for result in results:
            for key in result:
                if key.startswith('地理位置-') and key not in geo_keys:
                    geo_keys.append(key)
    return geo_keys


def build_result_columns(geo_result_keys=None):
    """
    构造纯查询结果文件的列顺序。

    返回:
        列名列表
    """
    if geo_result_keys is None:
        geo_result_keys = GEO_RESULT_KEYS
    return [
        'IP', '页面顶部标签', '反查域名', '类型', '使用场景', 'IP属性', '国家/地区',
    ] + list(geo_result_keys) + [
        'ASN', '企业', 'IP评分',
    ] + INTEL_RESULT_KEYS + [
        '数字地址', '备注', '查询状态',
    ]


def get_result_field_description(field_name):
    """
    获取查询结果字段说明。

    返回:
        字段说明文本；未知字段返回空字符串
    """
    if field_name.startswith('地理位置-'):
        source_name = field_name.replace('地理位置-', '', 1)
        if source_name in GEO_SOURCES:
            return f'{source_name} 来源的地理位置'
        return '页面实际出现但不在标准 14 源中的来源'
    return RESULT_FIELD_DESCRIPTIONS.get(field_name, '')


def build_result_field_description_rows(geo_result_keys=None):
    """
    构造纯查询结果文件的字段说明表。

    返回:
        包含序号、列号、字段、说明的 DataFrame
    """
    result_columns = build_result_columns(geo_result_keys)
    rows = []
    for index, field_name in enumerate(result_columns, 1):
        rows.append({
            '序号': index,
            '列号': column_index_to_letter(index - 1),
            '字段': field_name,
            '说明': get_result_field_description(field_name),
        })
    return pd.DataFrame(rows, columns=['序号', '列号', '字段', '说明'])


def build_append_column_mappings(geo_result_keys=None):
    """
    构造原表回填列与查询结果字段的映射。

    返回:
        [(回填列名, 结果字段名), ...]
    """
    if geo_result_keys is None:
        geo_result_keys = GEO_RESULT_KEYS
    return [
        (QUERY_IP_APPEND_COLUMN, 'IP'),
        ('查询_页面顶部标签', '页面顶部标签'),
        ('查询_反查域名', '反查域名'),
        ('查询_类型', '类型'),
        ('查询_使用场景', '使用场景'),
        ('查询_IP属性', 'IP属性'),
        ('查询_国家地区', '国家/地区'),
    ] + [
        (f'查询_{geo_key}', geo_key) for geo_key in geo_result_keys
    ] + [
        ('查询_ASN', 'ASN'),
        ('查询_企业', '企业'),
        ('查询_IP评分', 'IP评分'),
        ('查询_数字地址', '数字地址'),
        ('查询_备注', '备注'),
    ] + [
        (key, key) for key in INTEL_RESULT_KEYS
    ] + [
        ('查询_状态', '查询状态'),
    ]


def build_empty_result(ip):
    """
    创建单个 IP 的完整结果模板，确保所有输出字段稳定存在。

    返回:
        查询结果字典
    """
    result = {}
    for key in BASE_RESULT_KEYS + LEGACY_INTEL_KEYS + INTEL_RESULT_KEYS + GEO_RESULT_KEYS:
        result[key] = ''
    result['IP'] = ip
    return result


def is_excel_column_reference(value):
    """
    判断用户输入是否是 Excel 列字母引用，例如 A、H、AA。

    注意:
        `ip`/`IP` 应优先按列名关键词处理，不能误判为 Excel 的 IP 列。
    """
    value = str(value or '').strip()
    return bool(value) and value.lower() != 'ip' and len(value) <= 2 and value.isalpha()


def setup_driver():
    """配置并启动Chrome浏览器"""
    chrome_driver_path = r''  # 可选：留空时让 Selenium 自动查找 ChromeDriver
    chrome_binary_path = r''  # 可选：仅在使用独立 Chrome 时填写浏览器路径

    options = Options()
    if chrome_binary_path:
        options.binary_location = chrome_binary_path
    # options.add_argument('--headless')  # 无头模式，取消注释可隐藏浏览器窗口
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

    service = Service(chrome_driver_path) if chrome_driver_path else Service()
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(30)
    return driver


def is_public_ipv4(ip_text):
    """
    判断 IPv4 地址是否适合查询 iplark。

    参数:
        ip_text: 已通过四段数字格式校验的 IPv4 字符串

    返回:
        True 表示公网可路由 IPv4；私网、回环、链路本地、保留等地址返回 False
    """
    try:
        ip_obj = ipaddress.ip_address(ip_text)
    except ValueError:
        return False
    if ip_obj.version != 4:
        return False
    if (
        ip_obj.is_private or
        ip_obj.is_loopback or
        ip_obj.is_link_local or
        ip_obj.is_multicast or
        ip_obj.is_reserved or
        ip_obj.is_unspecified
    ):
        return False
    return ip_obj.is_global


def extract_ip_from_hostname(hostname):
    """从主机名中提取公网 IPv4 地址，如果是域名、内网或保留地址则返回 None。"""
    hostname = str(hostname or '').strip().replace('\t', '')
    ip_pattern = r'^(\d{1,3}\.){3}\d{1,3}$'
    if re.fullmatch(ip_pattern, hostname):
        octets = hostname.split('.')
        if all(octet.isdigit() and 0 <= int(octet) <= 255 for octet in octets) and is_public_ipv4(hostname):
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
    except Exception:
        return []


def dedupe_preserve_order(values):
    """
    按原顺序去重文本列表。

    返回:
        去重后的非空字符串列表
    """
    seen = set()
    result = []
    for value in values or []:
        text = re.sub(r'\s+', ' ', str(value or '')).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def looks_like_reverse_hostname(value):
    """
    判断页面顶部子标签是否像 DNS 反查主机名。

    返回:
        True 表示该值适合写入“反查域名”
    """
    hostname = str(value or '').strip().rstrip('.')
    if not hostname or ' ' in hostname or '.' not in hostname:
        return False
    if extract_ip_from_hostname(hostname):
        return False

    labels = hostname.split('.')
    for label in labels:
        if not label:
            return False
        if len(label) > 63:
            return False
        if label.startswith('-') or label.endswith('-'):
            return False
        if not re.fullmatch(r'[A-Za-z0-9-]+', label):
            return False

    top_level_label = labels[-1]
    return len(top_level_label) >= 2 and any(char.isalpha() for char in top_level_label)


def extract_top_sub_tags(driver, result):
    """
    提取页面标题下方的子标签，并拆出 DNS 反查主机名。

    参数:
        driver: Selenium WebDriver
        result: 待更新的结果字典
    """
    sub_tags = dedupe_preserve_order(
        safe_find_texts(driver, By.CSS_SELECTOR, '#hostname-container .sub-tag')
    )
    if not sub_tags:
        return

    result['页面顶部标签'] = '; '.join(sub_tags)

    reverse_hostnames = [
        tag for tag in sub_tags
        if looks_like_reverse_hostname(tag)
    ]
    if reverse_hostnames:
        result['反查域名'] = '; '.join(reverse_hostnames)


def extract_geo_locations(driver, result):
    """
    从 iplark 页面提取 14 个地理位置数据源的结果。

    参数:
        driver: Selenium WebDriver
        result: 待更新的结果字典
    """
    geo_source_divs = driver.find_elements(By.CSS_SELECTOR, '.geo-source')
    for geo_div in geo_source_divs:
        try:
            raw_source = geo_div.find_element(By.CSS_SELECTOR, '.source-tag').text
            source_name = normalize_geo_source_name(raw_source)
            if not source_name:
                continue

            value_spans = geo_div.find_elements(By.CSS_SELECTOR, 'span:not(.source-tag)')
            geo_text = ' '.join(s.text.strip() for s in value_spans if s.text.strip())
            if geo_text:
                result[f'地理位置-{source_name}'] = geo_text
        except Exception:
            continue


def extract_ip_intelligence(driver, result):
    """
    从 IP 情报区域提取字段，并同时兼容旧字段名。

    参数:
        driver: Selenium WebDriver
        result: 待更新的结果字典
    """
    try:
        intel_elem = driver.find_element(By.ID, 'ip-intelligence')
    except NoSuchElementException:
        return

    span_elems = intel_elem.find_elements(By.CSS_SELECTOR, 'span')
    legacy_keys = set(LEGACY_INTEL_KEYS)

    i = 0
    while i < len(span_elems):
        strongs = span_elems[i].find_elements(By.TAG_NAME, 'strong')
        if not strongs:
            i += 1
            continue

        raw_label = normalize_label_text(strongs[0].text)
        canonical_label = INTEL_LABEL_ALIASES.get(raw_label, raw_label)
        key = INTEL_RESULT_KEY_BY_LABEL.get(raw_label)

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
        if canonical_label in legacy_keys and value and value != '-' and not result.get(canonical_label):
            result[canonical_label] = value
            if canonical_label == '使用类型' and not result.get('使用场景'):
                result['使用场景'] = value

        i += 1


def get_ip_info(driver, ip, retry_count=2):
    """查询单个IP的信息，带重试机制"""
    url = f"https://iplark.com/{ip}"
    result = build_empty_result(ip)

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
            except TimeoutException:
                pass

            # 3. 等待页面完全加载（document.readyState）
            wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')

            # 4. 短暂等待动态内容渲染
            time.sleep(1.5)

            # 获取页面标题下方的子标签和反查域名
            extract_top_sub_tags(driver, result)

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
                except Exception:
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

            # 获取地理位置（14个数据源多源对比）和 IP 情报
            extract_geo_locations(driver, result)
            extract_ip_intelligence(driver, result)

            result['查询状态'] = '成功'
            return result

        except TimeoutException:
            if attempt < retry_count:
                print(f"    超时，重试 {attempt + 1}/{retry_count}...")
                time.sleep(2)
                continue
            result['查询状态'] = '超时'
        except WebDriverException:
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


def column_index_to_letter(index):
    """
    将Excel列索引转换为列字母（从0开始）。

    返回:
        Excel列字母，例如 0 -> A，25 -> Z，26 -> AA
    """
    if index < 0:
        raise ValueError('列索引不能小于0')

    result = ''
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result


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
                    df = pd.read_csv(
                        file_path,
                        encoding=encoding,
                        dtype=str,
                        keep_default_na=False
                    )
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
                    dtype=str,
                    keep_default_na=False,
                    engine='openpyxl' if file_ext == '.xlsx' else 'xlrd'
                )
            except Exception:
                sheets = pd.read_excel(
                    file_path,
                    sheet_name=None,
                    dtype=str,
                    keep_default_na=False
                )
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


def parse_args():
    """
    解析命令行参数。

    返回:
        argparse.Namespace，包含运行配置参数
    """
    parser = argparse.ArgumentParser(
        description='从 Excel/CSV 文件或命令行 IP 查询 iplark.com 获取详细信息。'
    )
    parser.add_argument(
        '-i', '--input',
        dest='input_file',
        help='输入文件路径，支持 .csv、.xlsx、.xls；不传则使用脚本内 INPUT_FILE 配置。'
    )
    parser.add_argument(
        '-ip', '--ip',
        dest='direct_ips',
        nargs='+',
        default=[],
        metavar='IP',
        help='直接指定一个或多个 IP 查询，不读取输入文件。示例: -ip 1.2.3.4 2.3.4.5'
    )
    parser.add_argument(
        '--retry-from',
        dest='retry_from',
        help='从历史查询结果 Excel 中筛选失败 IP 重试。'
    )
    parser.add_argument(
        '--retry-ip',
        dest='retry_ips',
        action='append',
        default=[],
        help='指定一个要重试的 IP，可重复传入。'
    )
    parser.add_argument(
        '--retry-ips',
        dest='retry_ips_csv',
        help='逗号、分号或空白分隔的多个 IP。'
    )
    parser.add_argument(
        '--force',
        action='store_true',
        help='即使历史结果中 IP 已成功，也重新查询。'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='只列出将要查询的 IP，不启动浏览器，不写入文件。'
    )
    parser.add_argument(
        '-o', '--output-dir',
        dest='output_dir',
        help='输出目录；不传则使用输入文件或 retry-from 文件所在目录。'
    )
    parser.add_argument(
        '--ip-column',
        dest='ip_column',
        help='命令行指定 IP 列，优先级高于脚本内 IP_COLUMN。'
    )
    return parser.parse_args()


def normalize_input_path(file_path):
    """
    规范化输入文件路径，支持环境变量和用户目录。

    参数:
        file_path: 用户配置或命令行传入的文件路径

    返回:
        规范化后的文件路径
    """
    return os.path.normpath(os.path.expandvars(os.path.expanduser(file_path)))


def get_ip_column_names(df, ip_column=None):
    """
    确定 IP 所在列的列名列表。

    参数:
        df: DataFrame
        ip_column: 用户指定的列

    返回:
        列名列表。自动模式会按列内容扫描，只有出现有效 IPv4 的列才返回。
    """
    if df is None or len(df.columns) == 0:
        print("当前工作表没有可用的列")
        return []

    if ip_column is None:
        return detect_ip_columns_by_content(df)

    ip_column_text = str(ip_column).strip()
    if ip_column_text in df.columns:
        return [ip_column_text]

    if is_excel_column_reference(ip_column_text):
        col_index = column_letter_to_index(ip_column_text)
        if col_index < len(df.columns):
            return [df.columns[col_index]]
        print(f"列 {ip_column_text} 超出范围，文件只有 {len(df.columns)} 列")
        return []

    for col in df.columns:
        if ip_column_text.lower() in str(col).lower():
            return [col]
    print(f"未找到列: {ip_column_text}")
    print(f"可用的列: {list(df.columns)}")
    return []


def get_ip_column_name(df, ip_column=None):
    """
    确定 IP 所在列的第一个列名。

    返回:
        列名 或 None
    """
    ip_column_names = get_ip_column_names(df, ip_column)
    return ip_column_names[0] if ip_column_names else None


def detect_ip_columns_by_content(df):
    """
    按单元格内容自动识别包含公网 IPv4 地址的列。

    参数:
        df: DataFrame

    返回:
        包含至少一个公网 IPv4 地址的列名列表
    """
    ip_column_names = []
    for column_name in df.columns:
        for value in df[column_name]:
            if pd.isna(value):
                continue
            if extract_ip_from_hostname(value):
                ip_column_names.append(column_name)
                break
    return ip_column_names


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

        ip_column_names = get_ip_column_names(df_original, ip_column)
        if not ip_column_names:
            print(f"工作表 [{sheet_name}] 未找到IP列，跳过")
            continue

        column_labels = [
            f"{column_name} (索引: {list(df_original.columns).index(column_name)})"
            for column_name in ip_column_names
        ]
        print(f"工作表 [{sheet_name}] 使用列: {', '.join(column_labels)}")

        sheet_ips = []
        sheet_ip_to_rows = {}
        for ip_column_name in ip_column_names:
            column_ips, column_ip_to_rows = extract_ips_from_column(df_original, ip_column_name)
            for ip in column_ips:
                if ip not in sheet_ip_to_rows:
                    sheet_ip_to_rows[ip] = []
                    sheet_ips.append(ip)
                for row_idx in column_ip_to_rows[ip]:
                    if row_idx not in sheet_ip_to_rows[ip]:
                        sheet_ip_to_rows[ip].append(row_idx)
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


def split_ip_argument_values(values):
    """
    拆分命令行 IP 参数值。

    参数:
        values: 原始参数字符串列表，支持逗号、分号和空白分隔

    返回:
        拆分后的字符串列表，不做 IP 有效性校验
    """
    raw_values = []
    for value in values or []:
        if value is None:
            continue
        parts = re.split(r'[,;\s]+', str(value))
        raw_values.extend(part.strip() for part in parts if part.strip())
    return raw_values


def normalize_ip_values(values):
    """
    将输入值规范化为有效 IP 列表。

    参数:
        values: 命令行或 Excel 中提取的原始值

    返回:
        有效 IP 字符串列表，保持原始顺序
    """
    ips = []
    for value in values or []:
        if value is None or pd.isna(value):
            continue
        value_text = str(value).strip().replace('\t', '')
        if not value_text:
            continue
        ip = extract_ip_from_hostname(value_text)
        if ip:
            ips.append(ip)
        else:
            print(f"跳过无效 IP: {value_text}")
    return ips


def dedupe_ips(ips):
    """
    去重 IP 并保留首次出现顺序。

    参数:
        ips: IP 字符串列表

    返回:
        去重后的 IP 字符串列表
    """
    unique_ips = []
    seen = set()
    for ip in ips or []:
        if ip in seen:
            continue
        seen.add(ip)
        unique_ips.append(ip)
    return unique_ips


def is_success_status(status):
    """
    判断查询状态是否成功。

    参数:
        status: 查询状态值

    返回:
        True 表示状态为“成功”
    """
    return str(status or '').strip() == '成功'


def clean_cell_value(value):
    """
    清理 Excel 单元格值。

    返回:
        空值转为空字符串，其他值保持原值
    """
    return '' if pd.isna(value) else value


def normalize_text_preserve_column_name(column_name):
    """
    规范化列名用于判断是否需要按文本保留。

    返回:
        小写并去除空白、下划线、连字符和常见括号后的列名
    """
    column_text = str(column_name or '').strip().casefold()
    return re.sub(r'[\s_\-()（）]+', '', column_text)


def is_text_preservation_column(column_name):
    """
    判断列是否应按 Excel 文本列写出。

    账号、证件号、QQ、用户 ID 等标识符即使全是数字，也不能当数值处理。
    """
    column_text = str(column_name or '').strip()
    if column_text in TEXT_PRESERVE_EXACT_COLUMNS:
        return True

    normalized = normalize_text_preserve_column_name(column_text)
    if normalized in {
        normalize_text_preserve_column_name(name)
        for name in TEXT_PRESERVE_EXACT_COLUMNS
    }:
        return True

    for keyword in TEXT_PRESERVE_COLUMN_KEYWORDS:
        keyword_normalized = normalize_text_preserve_column_name(keyword)
        if keyword_normalized and keyword_normalized in normalized:
            return True

    return False


def looks_like_long_numeric_identifier(value):
    """
    判断值是否像长数字标识符。

    返回:
        True 表示该值不适合作为 Excel 数值写出
    """
    if value is None or pd.isna(value):
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return len(str(abs(value))) >= LONG_NUMERIC_TEXT_MIN_DIGITS
    if isinstance(value, float):
        if not value.is_integer():
            return False
        return len(f'{abs(value):.0f}') >= LONG_NUMERIC_TEXT_MIN_DIGITS

    value_text = str(value).strip()
    if not value_text:
        return False
    value_text = value_text.lstrip("'").replace('\t', '').replace(' ', '')

    if re.fullmatch(r'[+-]?\d+', value_text):
        digits = value_text.lstrip('+-')
        return (
            len(digits) >= LONG_NUMERIC_TEXT_MIN_DIGITS or
            (digits.startswith('0') and len(digits) > 1)
        )
    return bool(re.fullmatch(r'[+-]?\d+(?:\.\d+)?[eE][+-]?\d+', value_text))


def column_has_long_numeric_identifier_values(series):
    """
    判断列中是否出现长数字标识符值。

    返回:
        True 表示整列写出时应使用文本格式
    """
    if series is None:
        return False
    return any(looks_like_long_numeric_identifier(value) for value in series)


def stringify_text_preserved_value(value):
    """
    将文本保真列的值转成适合写入 Excel 的字符串。

    返回:
        字符串；空值返回空字符串
    """
    if value is None or pd.isna(value):
        return ''
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return f'{value:.0f}'
    return str(value)


def prepare_dataframe_for_excel(df):
    """
    写入 Excel 前准备 DataFrame，并标记需要设置文本格式的列。

    返回:
        (处理后的 DataFrame, 需要按文本格式写出的列名集合)
    """
    if df is None:
        return df, set()

    df_output = df.copy()
    text_columns = set()
    for column in df_output.columns:
        if (
            is_text_preservation_column(column) or
            column_has_long_numeric_identifier_values(df_output[column])
        ):
            text_columns.add(column)

    for column in text_columns:
        df_output[column] = df_output[column].map(stringify_text_preserved_value)

    return df_output, text_columns


def apply_excel_text_formats(writer, sheet_name, df, text_columns, index=False):
    """
    对已写出的 Excel 工作表设置文本列格式。

    注意:
        必须配合 prepare_dataframe_for_excel 把值转成字符串；仅设置 number_format
        无法修复已经作为数值写出的长数字。
    """
    if not text_columns:
        return

    worksheet = writer.sheets.get(sheet_name)
    if worksheet is None:
        return

    index_offset = 1 if index else 0
    for column in text_columns:
        if column not in df.columns:
            continue
        excel_column_index = list(df.columns).index(column) + 1 + index_offset
        excel_column_letter = get_column_letter(excel_column_index)
        for cell in worksheet[excel_column_letter]:
            cell.number_format = EXCEL_TEXT_NUMBER_FORMAT


def write_dataframe_to_excel(writer, df, sheet_name, index=False):
    """
    写出 DataFrame，并对标识符类列做 Excel 文本保真处理。
    """
    df_output, text_columns = prepare_dataframe_for_excel(df)
    df_output.to_excel(writer, index=index, sheet_name=sheet_name)
    apply_excel_text_formats(
        writer,
        sheet_name,
        df_output,
        text_columns,
        index=index
    )


def is_query_append_column(column_name):
    """
    判断列是否属于原表回填的查询结果列。

    返回:
        True 表示该列不是原始输入列
    """
    column_text = str(column_name)
    return (
        column_text == QUERY_IP_APPEND_COLUMN or
        column_text.startswith('查询_') or
        column_text in INTEL_RESULT_KEYS
    )


def collect_geo_result_keys_from_augmented_columns(columns):
    """
    从原表回填列中提取地理位置结果字段。

    返回:
        标准 14 源加页面实际出现的新增来源字段
    """
    geo_result_keys = list(GEO_RESULT_KEYS)
    for column in columns:
        column_text = str(column)
        if not column_text.startswith('查询_地理位置-'):
            continue
        geo_key = column_text.replace('查询_', '', 1)
        if geo_key not in geo_result_keys:
            geo_result_keys.append(geo_key)
    return geo_result_keys


def build_augmented_result_column_mapping(columns):
    """
    构造原表回填列到纯查询结果字段的映射。

    返回:
        {原表回填列名: 查询结果字段名}
    """
    geo_result_keys = collect_geo_result_keys_from_augmented_columns(columns)
    return dict(build_append_column_mappings(geo_result_keys))


def get_history_ip_column_name(df, ip_column=None):
    """
    确定历史结果文件中用于 retry 的 IP 列。

    参数:
        df: 历史结果 DataFrame
        ip_column: 用户指定的原始 IP 列

    返回:
        列名 或 None
    """
    ip_column_names = get_history_ip_column_names(df, ip_column)
    return ip_column_names[0] if ip_column_names else None


def get_history_ip_column_names(df, ip_column=None):
    """
    确定历史结果文件中用于 retry 的 IP 列列表。

    参数:
        df: 历史结果 DataFrame
        ip_column: 用户指定的原始 IP 列

    返回:
        列名列表；未指定列时返回所有原始 IP 候选列
    """
    if df is None or len(df.columns) == 0:
        return []

    if 'IP' in df.columns:
        return ['IP']

    source_columns = [col for col in df.columns if not is_query_append_column(col)]

    if ip_column is None:
        ip_columns = [col for col in source_columns if 'ip' in str(col).lower()]
        if ip_columns:
            return ip_columns
        if QUERY_IP_APPEND_COLUMN in df.columns:
            return [QUERY_IP_APPEND_COLUMN]
        return []

    ip_column_text = str(ip_column).strip()
    if ip_column_text in df.columns:
        return [ip_column_text]

    if is_excel_column_reference(ip_column_text):
        col_index = column_letter_to_index(ip_column_text)
        if col_index < len(df.columns):
            return [df.columns[col_index]]
        print(f"列 {ip_column_text} 超出范围，文件只有 {len(df.columns)} 列")
        return []

    for col in source_columns:
        if ip_column_text.lower() in str(col).lower():
            return [col]
    return []


def extract_ip_from_row_values(row_values, ip_column_names):
    """
    从一行的多个候选 IP 列中提取第一个有效 IP。

    返回:
        (IP, 使用的列名, 原始值)；未找到有效 IP 时 IP 为空
    """
    invalid_values = []
    for column_name in ip_column_names:
        raw_value = row_values.get(column_name, '')
        if raw_value:
            invalid_values.append(raw_value)
        ip = extract_ip_from_hostname(str(raw_value).strip())
        if ip:
            return ip, column_name, raw_value

    raw_value = ' / '.join(str(value) for value in invalid_values if str(value).strip())
    return None, None, raw_value


def row_to_query_result(row_values, ip, ip_column_name, is_augmented, append_column_mapping):
    """
    将历史结果行转换为纯查询结果结构。

    参数:
        row_values: 行数据字典
        ip: 标准化后的 IP
        ip_column_name: 该行使用的 IP 列名
        is_augmented: 是否为原表回填结果
        append_column_mapping: 原表回填列到结果字段的映射

    返回:
        查询结果字典
    """
    if not is_augmented:
        result = dict(row_values)
        result['IP'] = ip
        if '查询状态' not in result:
            result['查询状态'] = result.get('查询_状态', '')
        return result

    result = build_empty_result(ip)
    for append_column, result_key in append_column_mapping.items():
        if append_column in row_values:
            result[result_key] = row_values.get(append_column, '')
    result['IP'] = ip
    result['查询状态'] = result.get('查询状态', '')
    if not result['查询状态'] and '查询_状态' in row_values:
        result['查询状态'] = row_values.get('查询_状态', '')
    return result


def read_history_result_sheet(df, sheet_name, ip_column=None):
    """
    读取单个历史结果工作表。

    返回:
        (查询结果字典列表, 是否识别到可用结构)
    """
    if df is None or df.empty:
        return [], True

    ip_column_names = get_history_ip_column_names(df, ip_column)
    if not ip_column_names:
        print(f"工作表 [{sheet_name}] 未找到可用于 retry 的 IP 列")
        print(f"可用的列: {list(df.columns)}")
        return [], False

    is_augmented = any(str(column).startswith('查询_') for column in df.columns)
    append_column_mapping = build_augmented_result_column_mapping(df.columns)

    results = []
    for _, row in df.iterrows():
        row_values = {}
        for column in df.columns:
            row_values[column] = clean_cell_value(row[column])

        ip, ip_column_name, raw_ip = extract_ip_from_row_values(row_values, ip_column_names)
        if not ip:
            if raw_ip:
                print(f"跳过历史结果中的无效 IP: {raw_ip}")
            continue

        results.append(row_to_query_result(
            row_values,
            ip,
            ip_column_name,
            is_augmented,
            append_column_mapping
        ))

    return results, True


def read_query_results_excel(file_path, ip_column=None):
    """
    读取历史查询结果 Excel，兼容纯查询结果和原表回填结果。

    参数:
        file_path: 历史结果文件路径
        ip_column: 原表回填结果中的 IP 列配置

    返回:
        (查询结果字典列表, 成功标志)
    """
    file_path = normalize_input_path(file_path)
    excel_file = None
    try:
        excel_file = pd.ExcelFile(file_path)
        if not excel_file.sheet_names:
            print("历史结果文件没有可读取的工作表")
            return [], False

        if '查询结果' in excel_file.sheet_names:
            sheet_names = ['查询结果']
        else:
            sheet_names = [
                sheet_name for sheet_name in excel_file.sheet_names
                if sheet_name != '输出字段说明'
            ]
            if not sheet_names:
                sheet_names = list(excel_file.sheet_names)

        results = []
        recognized_sheet_count = 0
        for sheet_name in sheet_names:
            df = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                dtype=str,
                keep_default_na=False
            )
            sheet_results, recognized = read_history_result_sheet(
                df,
                sheet_name,
                ip_column=ip_column
            )
            if recognized:
                recognized_sheet_count += 1
            results.extend(sheet_results)
    except Exception as e:
        print(f"读取历史结果失败: {e}")
        return [], False
    finally:
        if excel_file is not None:
            excel_file.close()

    if recognized_sheet_count == 0:
        print("历史结果缺少可识别的 IP 列")
        print("纯查询结果需要包含 IP 列；原表回填结果需要原始 IP 列，必要时可用 --ip-column 指定。")
        return [], False

    return results, True


def get_retry_output_prefix(retry_from):
    """
    从历史结果文件名提取 retry 输出前缀。

    返回:
        文件名前缀，例如 test_ip_info_result_... -> test_
    """
    if not retry_from:
        return ''

    stem = os.path.splitext(os.path.basename(retry_from))[0]
    for marker in ['_ip_info_result_merged', '_ip_info_result', '_ip_info_retry']:
        marker_index = stem.find(marker)
        if marker_index >= 0:
            return stem[:marker_index + 1]

    for marker in ['ip_info_result_merged', 'ip_info_result', 'ip_info_retry']:
        if stem.startswith(marker):
            return ''

    return f'{stem}_'


def build_retry_output_file(output_dir, retry_from, time_suffix):
    """
    构造 retry 完整结果文件路径。

    返回:
        输出文件路径
    """
    prefix = get_retry_output_prefix(retry_from)
    return os.path.join(output_dir, f"{prefix}ip_info_retry{time_suffix}.xlsx")


def read_augmented_history_sheets(file_path):
    """
    读取原表回填格式的历史结果工作簿。

    返回:
        (工作表字典, 是否为原表回填格式)
    """
    file_path = normalize_input_path(file_path)
    excel_file = None
    try:
        excel_file = pd.ExcelFile(file_path)
        if '查询结果' in excel_file.sheet_names:
            return {}, False

        data_sheet_names = [
            sheet_name for sheet_name in excel_file.sheet_names
            if sheet_name != '输出字段说明'
        ]
        sheets = {}
        has_augmented_columns = False
        for sheet_name in data_sheet_names:
            df = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                dtype=str,
                keep_default_na=False
            )
            sheets[sheet_name] = df
            if any(is_query_append_column(column) for column in df.columns):
                has_augmented_columns = True

        return sheets, has_augmented_columns
    except Exception as e:
        print(f"读取原表回填历史结果失败: {e}")
        return {}, False
    finally:
        if excel_file is not None:
            excel_file.close()


def collect_ips_from_history_sheets(history_sheets, ip_column=None):
    """
    从原表回填历史工作簿提取 IP 到行位置的映射。

    返回:
        (唯一 IP 列表, IP 到 (工作表名, 行索引) 的映射)
    """
    ips = []
    ip_to_rows = {}

    for sheet_name, df_original in history_sheets.items():
        if df_original is None or df_original.empty:
            continue

        ip_column_names = get_history_ip_column_names(df_original, ip_column)
        if not ip_column_names:
            print(f"工作表 [{sheet_name}] 未找到可回填的 IP 列，保留原内容")
            continue

        for row_idx, row in df_original.iterrows():
            row_values = {}
            for column in df_original.columns:
                row_values[column] = clean_cell_value(row[column])

            ip, _, _ = extract_ip_from_row_values(row_values, ip_column_names)
            if not ip:
                continue
            if ip not in ip_to_rows:
                ip_to_rows[ip] = []
                ips.append(ip)
            ip_to_rows[ip].append((sheet_name, row_idx))

    return ips, ip_to_rows


def collect_augmented_workbook_geo_result_keys(history_sheets, results):
    """
    汇总原表回填工作簿和查询结果中的地理位置字段。

    返回:
        地理位置结果字段列表
    """
    geo_result_keys = collect_geo_result_keys(results)
    for df in history_sheets.values():
        if df is None:
            continue
        for column in df.columns:
            column_text = str(column)
            if not column_text.startswith('查询_地理位置-'):
                continue
            geo_key = column_text.replace('查询_', '', 1)
            if geo_key not in geo_result_keys:
                geo_result_keys.append(geo_key)
    return geo_result_keys


def save_augmented_retry_workbook(history_sheets, ip_to_rows, ip_to_result, output_file):
    """
    保存原表回填格式的 retry 完整结果。

    参数:
        history_sheets: 历史原表回填工作表字典
        ip_to_rows: IP 到 (工作表名, 行索引) 的映射
        ip_to_result: IP 到合并后查询结果的映射
        output_file: 输出 Excel 文件路径
    """
    ensure_parent_dir(output_file)
    geo_result_keys = collect_augmented_workbook_geo_result_keys(
        history_sheets,
        list(ip_to_result.values())
    )
    append_column_mappings = build_append_column_mappings(geo_result_keys)

    output_sheets = {}
    for sheet_name, df_original in history_sheets.items():
        if df_original is None:
            continue
        df_output = df_original.copy()
        for column_name, _ in append_column_mappings:
            if column_name not in df_output.columns:
                df_output[column_name] = ''
            else:
                df_output[column_name] = df_output[column_name].astype(object)
        output_sheets[sheet_name] = df_output

    for ip, row_indices in ip_to_rows.items():
        result = ip_to_result.get(ip)
        if result is None:
            continue
        for sheet_name, row_idx in row_indices:
            df_output = output_sheets.get(sheet_name)
            if df_output is None:
                continue
            for column_name, key in append_column_mappings:
                df_output.at[row_idx, column_name] = result.get(key, '')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df_output in output_sheets.items():
            write_dataframe_to_excel(
                writer,
                df_output,
                sheet_name=sheet_name,
                index=False
            )


def build_ip_to_result(results):
    """
    建立 IP 到查询结果的索引。

    参数:
        results: 查询结果字典列表

    返回:
        {IP: 查询结果字典}；重复 IP 优先保留成功结果，失败结果保留靠后的记录
    """
    ip_to_result = {}
    for result in results or []:
        ip = extract_ip_from_hostname(str(result.get('IP', '')).strip())
        if not ip:
            continue

        existing = ip_to_result.get(ip)
        if existing is None:
            ip_to_result[ip] = result
            continue

        existing_success = is_success_status(existing.get('查询状态'))
        current_success = is_success_status(result.get('查询状态'))
        if current_success or not existing_success:
            ip_to_result[ip] = result

    return ip_to_result


def select_retry_targets(existing_results, requested_ips, force=False):
    """
    根据历史结果和命令行指定 IP 计算本次重试目标。

    参数:
        existing_results: 历史查询结果列表
        requested_ips: 命令行指定 IP 列表
        force: 是否强制重查历史成功 IP

    返回:
        本次需要查询的唯一 IP 列表
    """
    requested_ips = dedupe_ips(requested_ips)
    if not existing_results:
        return requested_ips

    ip_to_result = build_ip_to_result(existing_results)

    if requested_ips:
        targets = []
        for ip in requested_ips:
            result = ip_to_result.get(ip)
            if result is None:
                print(f"指定 IP 不在历史结果中，跳过: {ip}")
                continue
            if force or not is_success_status(result.get('查询状态')):
                targets.append(ip)
            else:
                print(f"历史结果已成功，跳过: {ip}（如需重查请加 --force）")
        return dedupe_ips(targets)

    targets = []
    seen = set()
    for result in existing_results:
        ip = extract_ip_from_hostname(str(result.get('IP', '')).strip())
        if not ip or ip in seen:
            continue
        seen.add(ip)
        current_result = ip_to_result.get(ip)
        if current_result is not None and not is_success_status(current_result.get('查询状态')):
            targets.append(ip)
    return targets


def query_ips(driver, ips):
    """
    查询一批已去重 IP。

    参数:
        driver: Selenium WebDriver
        ips: IP 字符串列表

    返回:
        (查询结果列表, IP 到查询结果的映射)
    """
    results = []
    ip_to_result = {}
    for index, ip in enumerate(ips, 1):
        print(f"  查询 ({index}/{len(ips)}): {ip}")
        info = get_ip_info(driver, ip)
        results.append(info)
        ip_to_result[ip] = info

        status_icon = "OK" if is_success_status(info.get('查询状态')) else "X"
        print(f"    [{status_icon}] {info.get('类型', '')} | {info.get('国家/地区', '')} | 评分:{info.get('IP评分', '')}")

        if index < len(ips):
            time.sleep(2)

    return results, ip_to_result


def ensure_parent_dir(file_path):
    """
    确保输出文件所在目录存在。

    参数:
        file_path: 输出文件路径
    """
    output_dir = os.path.dirname(file_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)


def save_query_results(results, output_file):
    """
    保存纯查询结果 Excel。

    参数:
        results: 查询结果字典列表
        output_file: 输出 Excel 文件路径
    """
    ensure_parent_dir(output_file)
    df_result = pd.DataFrame(results)
    geo_result_keys = collect_geo_result_keys(results)
    columns_order = build_result_columns(geo_result_keys)
    for column in columns_order:
        if column not in df_result.columns:
            df_result[column] = ''
    df_result = df_result[columns_order]
    df_field_descriptions = build_result_field_description_rows(geo_result_keys)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        write_dataframe_to_excel(
            writer,
            df_result,
            sheet_name='查询结果',
            index=False
        )
        write_dataframe_to_excel(
            writer,
            df_field_descriptions,
            sheet_name='输出字段说明',
            index=False
        )


def save_augmented_workbook(original_sheets, ip_to_rows, ip_to_result, output_file):
    """
    保存原表追加查询结果列的 Excel。

    参数:
        original_sheets: 原始工作表字典
        ip_to_rows: IP 到 (工作表名, 行索引) 的映射
        ip_to_result: IP 到查询结果的映射
        output_file: 输出 Excel 文件路径
    """
    ensure_parent_dir(output_file)
    geo_result_keys = collect_geo_result_keys(list(ip_to_result.values()))
    append_column_mappings = build_append_column_mappings(geo_result_keys)
    append_columns = [column for column, _ in append_column_mappings]
    sheets_with_results = set()
    for row_indices in ip_to_rows.values():
        for sheet_name, _ in row_indices:
            sheets_with_results.add(sheet_name)

    output_sheets = {}
    for sheet_name, df_original in original_sheets.items():
        if df_original is None:
            continue
        df_output = df_original.copy()
        if sheet_name in sheets_with_results:
            for column in append_columns:
                if column not in df_output.columns:
                    df_output[column] = ''
                else:
                    df_output[column] = df_output[column].astype(object)
        output_sheets[sheet_name] = df_output

    for ip, row_indices in ip_to_rows.items():
        result = ip_to_result.get(ip)
        if result is None:
            continue
        for sheet_name, row_idx in row_indices:
            df_output = output_sheets.get(sheet_name)
            if df_output is None:
                continue
            for column_name, key in append_column_mappings:
                df_output.at[row_idx, column_name] = result.get(key, '')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df_output in output_sheets.items():
            write_dataframe_to_excel(
                writer,
                df_output,
                sheet_name=sheet_name,
                index=False
            )


def merge_query_results(existing_results, retry_results):
    """
    合并历史结果和本次重试结果。

    参数:
        existing_results: 历史查询结果列表
        retry_results: 本次重试查询结果列表

    返回:
        合并后的查询结果列表
    """
    existing_ip_to_result = build_ip_to_result(existing_results)
    merged_results = []
    ip_to_index = {}

    for result in existing_results or []:
        ip = extract_ip_from_hostname(str(result.get('IP', '')).strip())
        if not ip or ip in ip_to_index:
            continue
        ip_to_index[ip] = len(merged_results)
        merged_results.append(dict(existing_ip_to_result.get(ip, result)))

    for result in retry_results or []:
        ip = extract_ip_from_hostname(str(result.get('IP', '')).strip())
        if not ip:
            continue
        result_copy = dict(result)
        if ip in ip_to_index:
            merged_results[ip_to_index[ip]] = result_copy
        else:
            ip_to_index[ip] = len(merged_results)
            merged_results.append(result_copy)

    return merged_results


def count_success_results(results):
    """
    统计成功查询结果数量。

    参数:
        results: 查询结果字典列表

    返回:
        查询状态为成功的数量
    """
    return sum(1 for result in results or [] if is_success_status(result.get('查询状态')))


def print_target_ips(title, ips):
    """
    打印目标 IP 列表。

    参数:
        title: 列表标题
        ips: IP 字符串列表
    """
    print(f"{title}: {len(ips)} 个")
    for ip in ips:
        print(f"  - {ip}")


def print_run_header(title, timestamp, tz_str):
    """
    打印运行头部信息。

    参数:
        title: 工具标题
        timestamp: 时间戳字符串
        tz_str: 时区字符串
    """
    print("=" * 60)
    print(title)
    print(f"查询时间: {timestamp} ({tz_str})")
    print(f"地理位置数据源: {len(GEO_SOURCES)} 个")
    print("=" * 60)


def build_runtime_config(args, default_input_file, default_output_dir, default_ip_column):
    """
    根据命令行参数和脚本默认值构建运行配置。

    参数:
        args: argparse.Namespace
        default_input_file: 脚本内默认输入文件
        default_output_dir: 脚本内默认输出目录
        default_ip_column: 脚本内默认 IP 列配置

    返回:
        运行配置字典
    """
    input_file_provided = bool(args.input_file)
    input_file = args.input_file.strip() if args.input_file else (default_input_file or '')
    retry_from = args.retry_from.strip() if args.retry_from else ''
    output_dir = args.output_dir.strip() if args.output_dir else (default_output_dir or '')
    ip_column = args.ip_column.strip() if args.ip_column else default_ip_column

    if input_file:
        input_file = normalize_input_path(input_file)
    if retry_from:
        retry_from = normalize_input_path(retry_from)
    if output_dir:
        output_dir = normalize_input_path(output_dir)

    direct_argument_values = list(args.direct_ips or [])
    direct_ips = dedupe_ips(normalize_ip_values(split_ip_argument_values(direct_argument_values)))

    retry_argument_values = list(args.retry_ips or [])
    if args.retry_ips_csv:
        retry_argument_values.append(args.retry_ips_csv)
    requested_ips = dedupe_ips(normalize_ip_values(split_ip_argument_values(retry_argument_values)))

    direct_mode = bool(direct_argument_values)
    retry_mode = not direct_mode and bool(retry_from or retry_argument_values)
    if not output_dir:
        if direct_mode:
            output_dir = ''
        elif retry_mode:
            if retry_from:
                output_dir = os.path.dirname(retry_from)
            elif input_file_provided and input_file:
                output_dir = os.path.dirname(input_file)
        elif input_file:
            output_dir = os.path.dirname(input_file)

    return {
        'input_file': input_file,
        'input_file_provided': input_file_provided,
        'retry_from': retry_from,
        'direct_ips': direct_ips,
        'direct_mode': direct_mode,
        'requested_ips': requested_ips,
        'retry_mode': retry_mode,
        'force': args.force,
        'dry_run': args.dry_run,
        'output_dir': output_dir,
        'ip_column': ip_column,
    }


def run_normal_mode(config):
    """
    执行普通输入文件查询流程。

    参数:
        config: 运行配置字典
    """
    input_file = config['input_file']
    if not input_file:
        print("未指定输入文件，请使用 -i/--input 传入文件路径，或配置 INPUT_FILE。")
        return

    timestamp, tz_str = get_timestamp_with_timezone()
    time_suffix = f"_{timestamp}_{tz_str}"
    print_run_header("IP信息查询工具 - iplark.com", timestamp, tz_str)

    input_basename = os.path.splitext(os.path.basename(input_file))[0]
    output_dir = config['output_dir']
    output_file_1 = os.path.join(output_dir, f"ip_info_result{time_suffix}.xlsx")
    output_file_2 = os.path.join(output_dir, f"{input_basename}_ip_info_result{time_suffix}.xlsx")

    print(f"\n[1/4] 正在读取文件: {input_file}")
    original_sheets, success = read_file_to_dataframes(input_file)
    if not success:
        return

    print(f"读取到 {len(original_sheets)} 个工作表:")
    for sheet_name in original_sheets:
        print(f"  - {sheet_name}")

    ips, ip_to_rows = collect_ips_from_sheets(original_sheets, config['ip_column'])
    if not ips:
        print("未找到有效的IP地址！")
        return

    print(f"全部工作表共找到 {len(ips)} 个唯一IP地址:")
    for ip in ips:
        print(f"  - {ip}")

    if config['dry_run']:
        print("dry-run 模式，不启动浏览器，不写入 Excel。")
        return

    print("\n[2/4] 正在启动浏览器...")
    driver = setup_driver()

    print("\n[3/4] 正在查询IP信息...")
    try:
        results, ip_to_result = query_ips(driver, ips)
    finally:
        driver.quit()

    print("\n[4/4] 正在保存结果...")
    save_query_results(results, output_file_1)
    print(f"  文件1: {output_file_1}")

    save_augmented_workbook(original_sheets, ip_to_rows, ip_to_result, output_file_2)
    print(f"  文件2: {output_file_2}")

    success_count = count_success_results(results)
    print(f"\n查询完成: 成功 {success_count}/{len(ips)}")


def run_direct_ip_mode(config):
    """
    执行命令行直接指定 IP 的查询流程。

    参数:
        config: 运行配置字典
    """
    ips = config['direct_ips']
    timestamp, tz_str = get_timestamp_with_timezone()
    time_suffix = f"_{timestamp}_{tz_str}"
    print_run_header("IP信息查询工具 - iplark.com（命令行 IP）", timestamp, tz_str)

    if not ips:
        print("未找到有效的命令行 IP。")
        return

    print()
    print_target_ips("命令行 IP", ips)

    if config['dry_run']:
        print("dry-run 模式，不启动浏览器，不写入 Excel。")
        return

    print("\n[1/3] 正在启动浏览器...")
    driver = setup_driver()

    print("\n[2/3] 正在查询IP信息...")
    try:
        results, _ = query_ips(driver, ips)
    finally:
        driver.quit()

    output_file = os.path.join(config['output_dir'], f"ip_info_result{time_suffix}.xlsx")

    print("\n[3/3] 正在保存结果...")
    save_query_results(results, output_file)
    print(f"  文件: {output_file}")

    success_count = count_success_results(results)
    print(f"\n查询完成: 成功 {success_count}/{len(ips)}")


def run_retry_mode(config):
    """
    执行失败 IP 重试流程。

    参数:
        config: 运行配置字典
    """
    timestamp, tz_str = get_timestamp_with_timezone()
    time_suffix = f"_{timestamp}_{tz_str}"
    print_run_header("IP信息查询工具 - iplark.com（失败 IP 重试）", timestamp, tz_str)

    existing_results = []
    if config['retry_from']:
        print(f"\n[1/4] 正在读取历史结果: {config['retry_from']}")
        existing_results, success = read_query_results_excel(
            config['retry_from'],
            ip_column=config['ip_column']
        )
        if not success:
            return
        print(f"历史结果共读取 {len(existing_results)} 条")
    else:
        print("\n[1/4] 未提供历史结果文件，使用命令行指定 IP。")

    target_ips = select_retry_targets(
        existing_results,
        config['requested_ips'],
        force=config['force']
    )

    if not target_ips:
        print("没有需要重试的 IP。")
        return

    print()
    print_target_ips("Retry 目标 IP", target_ips)

    if config['dry_run']:
        print("dry-run 模式，不启动浏览器，不写入 Excel。")
        return

    print("\n[2/4] 正在启动浏览器...")
    driver = setup_driver()

    print("\n[3/4] 正在查询IP信息...")
    try:
        retry_results, _ = query_ips(driver, target_ips)
    finally:
        driver.quit()

    output_dir = config['output_dir']
    retry_output_file = build_retry_output_file(
        output_dir,
        config['retry_from'],
        time_suffix
    )

    print("\n[4/4] 正在保存结果...")
    if config['retry_from']:
        merged_results = merge_query_results(existing_results, retry_results)
        history_sheets, is_augmented_history = read_augmented_history_sheets(config['retry_from'])
        if is_augmented_history:
            _, history_ip_to_rows = collect_ips_from_history_sheets(
                history_sheets,
                config['ip_column']
            )
            save_augmented_retry_workbook(
                history_sheets,
                history_ip_to_rows,
                build_ip_to_result(merged_results),
                retry_output_file
            )
        else:
            save_query_results(merged_results, retry_output_file)
        print(f"  retry 完整文件: {retry_output_file}")

        if config['input_file_provided'] and config['input_file'] and not is_augmented_history:
            print(f"\n正在读取原始输入以生成原表 retry 完整文件: {config['input_file']}")
            original_sheets, success = read_file_to_dataframes(config['input_file'])
            if success:
                _, ip_to_rows = collect_ips_from_sheets(original_sheets, config['ip_column'])
                merged_ip_to_result = build_ip_to_result(merged_results)
                input_basename = os.path.splitext(os.path.basename(config['input_file']))[0]
                augmented_output_file = os.path.join(
                    output_dir,
                    f"{input_basename}_ip_info_retry{time_suffix}.xlsx"
                )
                save_augmented_workbook(
                    original_sheets,
                    ip_to_rows,
                    merged_ip_to_result,
                    augmented_output_file
                )
                print(f"  原表 retry 完整文件: {augmented_output_file}")
            else:
                print("原始输入读取失败，已跳过原表 retry 完整文件。")
    elif config['input_file_provided']:
        save_query_results(retry_results, retry_output_file)
        print(f"  retry 文件: {retry_output_file}")
        print("未提供 --retry-from，无法合并历史成功结果；已只保存本次重试结果。")
    else:
        save_query_results(retry_results, retry_output_file)
        print(f"  retry 文件: {retry_output_file}")

    retry_success_count = count_success_results(retry_results)
    print(f"\n本次 retry 完成: 成功 {retry_success_count}/{len(target_ips)}")
    if config['retry_from']:
        merged_success_count = count_success_results(merged_results)
        print(f"retry 完整结果: 成功 {merged_success_count}/{len(merged_results)}")


def main():
    args = parse_args()

    # ==================== 用户配置区域 ====================
    # 输入文件路径（支持 .csv, .xlsx, .xls 格式）
    INPUT_FILE = r'examples/input.xlsx'

    # 输出目录（留空则与输入文件同目录）
    OUTPUT_DIR = r''

    # IP地址所在列配置
    # 可选值:
    #   - None     : 自动检测（扫描各列内容，提取公网 IPv4 地址）
    #   - 'A'      : 使用A列（第1列）
    #   - 'B'      : 使用B列（第2列）
    #   - 'H'      : 使用H列（第8列）
    #   - 'AA'     : 使用AA列（第27列）
    #   - '登录ip' : 使用名为"登录ip"的列
    IP_COLUMN = None  # <-- 修改这里来指定IP所在列
    # ======================================================

    config = build_runtime_config(args, INPUT_FILE, OUTPUT_DIR, IP_COLUMN)
    if config['direct_mode']:
        run_direct_ip_mode(config)
    elif config['retry_mode']:
        run_retry_mode(config)
    else:
        run_normal_mode(config)


if __name__ == '__main__':
    main()
