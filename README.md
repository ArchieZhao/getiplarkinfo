# IP信息查询工具

从 Excel/CSV 文件中读取 IP 地址，自动查询 [iplark.com](https://iplark.com/) 获取详细信息，包括：IP类型、归属地、ASN、企业、IP评分等。

## 功能特点

- 支持多种文件格式：`.csv`、`.xlsx`、`.xls`
- 灵活的列选择：支持列字母（A/B/C...）或列名
- 自动生成两个结果文件：
  - 纯查询结果表
  - 原表 + 查询结果追加列
- 文件名自动添加时间戳和时区

## 环境要求

- Python 3.7+
- Chrome 浏览器
- ChromeDriver（版本需与 Chrome 匹配）

---

## 一、安装依赖

```bash
pip install -r requirements.txt
```

或手动安装：

```bash
pip install pandas selenium openpyxl xlrd
```

---

## 二、配置 Chrome 和 ChromeDriver

### 1. 下载 Chrome 浏览器

如果需要独立版本（不影响系统已安装的 Chrome）：

- 下载地址：https://googlechromelabs.github.io/chrome-for-testing/

选择对应系统的 `chrome` 版本下载，解压到任意目录，例如：
```
D:\Program\chrome-win64\chrome.exe
```

### 2. 下载 ChromeDriver

- 下载地址：https://googlechromelabs.github.io/chrome-for-testing/

**重要**：ChromeDriver 版本必须与 Chrome 版本一致！

下载后解压到任意目录，例如：
```
D:\Program\chromedriver-win64\chromedriver.exe
```

### 3. 查看 Chrome 版本

打开 Chrome，访问 `chrome://version/`，查看版本号（如 `130.0.6723.58`）。

### 4. 修改脚本中的路径

打开 `get_ip_info.py`，找到 `setup_driver` 函数，修改以下两行：

```python
def setup_driver():
    """配置并启动Chrome浏览器"""
    chrome_driver_path = r'D:\Program\chromedriver-win64\chromedriver.exe'  # 修改为你的路径
    chrome_binary_path = r'D:\Program\chrome-win64\chrome.exe'              # 修改为你的路径
```

---

## 三、使用方法

### 1. 配置输入文件和列

打开 `get_ip_info.py`，找到 `main` 函数中的配置区域：

```python
def main():
    # ==================== 用户配置区域 ====================
    # 输入文件路径（支持 .csv, .xlsx, .xls 格式）
    INPUT_FILE = r'E:\AAAAAcodedata\getiplarkinfo\testIP.xlsx'

    # 输出目录（留空则与输入文件同目录）
    OUTPUT_DIR = r''

    # IP地址所在列配置
    IP_COLUMN = None  # <-- 修改这里来指定IP所在列
    # ======================================================
```

### 2. 运行脚本

```bash
python get_ip_info.py
```

### 3. 查看结果

脚本会在输出目录生成两个文件：

| 文件 | 说明 |
|------|------|
| `ip_info_result_2025-12-21-234310_UTC+8.xlsx` | 纯查询结果（每个唯一IP一行） |
| `原文件名_ip_info_result_2025-12-21-234310_UTC+8.xlsx` | 原表内容 + 追加的查询结果列 |

---

## 四、根据不同表格修改配置

### IP_COLUMN 配置说明

| 值 | 说明 | 示例 |
|----|------|------|
| `None` | 自动检测（查找列名包含 "ip" 的列） | `IP_COLUMN = None` |
| `'A'` | 使用 A 列（第1列） | `IP_COLUMN = 'A'` |
| `'B'` | 使用 B 列（第2列） | `IP_COLUMN = 'B'` |
| `'H'` | 使用 H 列（第8列） | `IP_COLUMN = 'H'` |
| `'AA'` | 使用 AA 列（第27列） | `IP_COLUMN = 'AA'` |
| `'登录ip'` | 使用名为 "登录ip" 的列 | `IP_COLUMN = '登录ip'` |
| `'IP地址'` | 使用名为 "IP地址" 的列 | `IP_COLUMN = 'IP地址'` |

### 示例

**示例1**：表格的 IP 在 A 列
```python
IP_COLUMN = 'A'
```

**示例2**：表格的 IP 在名为 "源IP" 的列
```python
IP_COLUMN = '源IP'
```

**示例3**：让脚本自动查找包含 "ip" 的列
```python
IP_COLUMN = None
```

---

## 五、输出字段说明

查询结果包含以下字段：

| 字段 | 说明 |
|------|------|
| IP | IP地址 |
| 类型 | 家宽、数据中心、商宽等 |
| IP属性 | 原生IP、广播IP |
| 国家/地区 | 所属国家或地区 |
| 地理位置 | 详细地理位置 |
| ASN | 自治系统编号 |
| 企业 | 所属企业/运营商 |
| IP评分 | 0-100分 |
| 使用类型 | ISP、数据中心等 |
| IP类型 | 分类标签 |
| 公共代理 | 是否为公共代理 |
| 查询状态 | 成功/超时/错误 |

---

## 六、常见问题

### Q: 提示 ChromeDriver 版本不匹配？

确保 ChromeDriver 版本与 Chrome 浏览器版本一致。访问 `chrome://version/` 查看 Chrome 版本，然后下载对应版本的 ChromeDriver。

### Q: 查询超时怎么办？

可能是网络问题。脚本已内置重试机制（默认重试2次）。如需调整，修改 `get_ip_info` 函数的 `retry_count` 参数。

### Q: 如何隐藏浏览器窗口？

在 `setup_driver` 函数中，取消以下行的注释：
```python
# options.add_argument('--headless')  # 取消这行注释即可隐藏浏览器
```

### Q: 支持哪些文件格式？

- `.csv` - CSV 文件（自动检测编码：UTF-8、GBK、GB2312、Latin1）
- `.xlsx` - Excel 2007+ 格式
- `.xls` - Excel 97-2003 格式

---

## 七、文件结构

```
getiplarkusage/
├── get_ip_info.py      # 主脚本
├── requirements.txt    # Python 依赖
├── README.md           # 说明文档
├── testIP.xlsx         # 示例输入文件
└── ip_info_result_*.xlsx  # 生成的结果文件
```

---

## 许可证

MIT License
