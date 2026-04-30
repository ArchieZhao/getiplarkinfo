# IP信息查询工具

从 Excel/CSV 文件中读取 IP 地址，自动查询 [iplark.com](https://iplark.com/) 获取详细信息，包括：类型、国家/地区、ASN、企业、使用场景、IP评分与 IP 情报等。

## 功能特点

- 支持多种文件格式：`.csv`、`.xlsx`、`.xls`
- Excel 文件会自动读取全部工作表，工作表不需要命名为 `Sheet1`、`Sheet2`
- 灵活的列选择：支持列字母（A/B/C...）或列名
- 自动生成两个结果文件：
  - 纯查询结果表
  - 原表 + 查询结果追加列（Excel 会保留原有多个工作表及其名称）
- 适配 iplark 新版页面结构：
  - “数字地址”会先点击小眼睛显示完整数字，再提取纯数字
  - “国家/地区”会拼接国旗 alt 与文本（例如 `China中国`）
  - 新增“使用场景”字段
  - 新增一组“IP情报-…”字段，并保留 `-` 占位值
  - 地理位置支持 14 个标准数据源多源对比采集，并会自动追加页面实际出现的新来源
- 文件名自动添加时间戳和时区

## 环境要求

- Python 3.7+
- 建议使用虚拟环境或 Conda 环境
- Chrome 浏览器
- ChromeDriver（版本需与 Chrome 匹配）

---

## 一、安装依赖

建议先创建并激活虚拟环境或 Conda 环境，再安装依赖：

```powershell
python -m pip install -r requirements.txt
```

或手动安装：

```powershell
python -m pip install pandas selenium openpyxl xlrd
```

如果系统中存在多个 Python 环境，Windows PowerShell 下也可以直接指定目标环境的 `python.exe`：

```powershell
& "<Python环境路径>\python.exe" -m pip install -r requirements.txt
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
    INPUT_FILE = r'examples/input.xlsx'

    # 输出目录（留空则与输入文件同目录）
    OUTPUT_DIR = r''

    # IP地址所在列配置
    IP_COLUMN = None  # <-- 修改这里来指定IP所在列
    # ======================================================
```

### 2. 运行脚本

进入项目目录后运行：

```powershell
python get_ip_info.py
```

如果系统中存在多个 Python 环境，可以指定目标解释器运行：

```powershell
& "<Python环境路径>\python.exe" get_ip_info.py
```

### 3. 查看结果

脚本会在输出目录生成两个文件：

| 文件 | 说明 |
|------|------|
| `ip_info_result_2025-12-21-234310_UTC+8.xlsx` | 纯查询结果（全部工作表中的每个唯一IP一行） |
| `原文件名_ip_info_result_2025-12-21-234310_UTC+8.xlsx` | 原表内容 + 追加的查询结果列；Excel 会按原工作表名分别写回 |

---

## 四、根据不同表格修改配置

### IP_COLUMN 配置说明

Excel 文件包含多个工作表时，脚本会逐个工作表识别 IP 列并提取 IP。相同 IP 即使出现在多个工作表，也只会查询一次，然后把结果回填到所有出现过的行。

| 值 | 说明 | 示例 |
|----|------|------|
| `None` | 每个工作表自动检测（查找列名包含 "ip" 的列） | `IP_COLUMN = None` |
| `'A'` | 每个工作表使用 A 列（第1列） | `IP_COLUMN = 'A'` |
| `'B'` | 使用 B 列（第2列） | `IP_COLUMN = 'B'` |
| `'H'` | 使用 H 列（第8列） | `IP_COLUMN = 'H'` |
| `'AA'` | 使用 AA 列（第27列） | `IP_COLUMN = 'AA'` |
| `'登录ip'` | 每个工作表查找名为或包含 "登录ip" 的列 | `IP_COLUMN = '登录ip'` |
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
| 地理位置-Ip-api | Ip-api 来源的地理位置 |
| 地理位置-Moe | Moe 来源的地理位置 |
| 地理位置-Moe+ | Moe+ 来源的地理位置 |
| 地理位置-Ease | Ease 来源的地理位置 |
| 地理位置-Internet | Internet 来源的地理位置 |
| 地理位置-Maxmind | Maxmind 来源的地理位置 |
| 地理位置-Ipstack | Ipstack 来源的地理位置 |
| 地理位置-IPinfo | IPinfo 来源的地理位置 |
| 地理位置-IP2Location | IP2Location 来源的地理位置 |
| 地理位置-Digital Element | Digital Element 来源的地理位置 |
| 地理位置-DB-IP | DB-IP 来源的地理位置 |
| 地理位置-Aliyun | Aliyun 来源的地理位置 |
| 地理位置-TencentCloud | TencentCloud 来源的地理位置 |
| 地理位置-Cloudflare | Cloudflare 来源的地理位置 |
| 地理位置-其他来源 | 页面实际出现但不在标准 14 源中的来源会自动追加，例如 `地理位置-IPLark`、`地理位置-CZ88`、`地理位置-Leak` |
| ASN | 自治系统编号 |
| 企业 | 所属企业/运营商 |
| 使用场景 | 网页“使用场景”字段（例如：普通宽带） |
| IP评分 | 0-100分 |
| 数字地址 | 会先点击小眼睛显示完整数字，再提取纯数字 |
| 备注 | ASN 规模等补充说明 |
| IP情报-使用类型 | IP情报区域中的“使用类型”（保留 `-`） |
| IP情报-威胁 | IP情报区域中的“威胁”（保留 `-`） |
| IP情报-IP类型 | IP情报区域中的“IP类型”（保留 `-`） |
| IP情报-提供商 | IP情报区域中的“提供商”（保留 `-`） |
| IP情报-公共代理 | IP情报区域中的“公共代理”（保留 `-`） |
| IP情报-代理类型 | IP情报区域中的“代理类型”（保留 `-`） |
| IP情报-标签 | IP情报区域中的“标签”（保留 `-`） |
| 查询状态 | 成功/超时/错误 |

原表追加列（`原文件名_ip_info_result_*.xlsx`）的顺序为：

1. `查询_类型`
2. `查询_使用场景`（紧跟在 `查询_类型` 后）
3. `查询_IP属性`
4. `查询_国家地区`
5. `查询_地理位置-Ip-api`
6. `查询_地理位置-Moe`
7. `查询_地理位置-Moe+`
8. `查询_地理位置-Ease`
9. `查询_地理位置-Internet`
10. `查询_地理位置-Maxmind`
11. `查询_地理位置-Ipstack`
12. `查询_地理位置-IPinfo`
13. `查询_地理位置-IP2Location`
14. `查询_地理位置-Digital Element`
15. `查询_地理位置-DB-IP`
16. `查询_地理位置-Aliyun`
17. `查询_地理位置-TencentCloud`
18. `查询_地理位置-Cloudflare`
19. `查询_地理位置-其他来源`（页面实际出现但不在标准 14 源中的来源会自动追加，例如 `查询_地理位置-IPLark`）
20. `查询_ASN`
21. `查询_企业`
22. `查询_IP评分`
23. `查询_数字地址`
24. `查询_备注`
25. `IP情报-使用类型`
26. `IP情报-威胁`
27. `IP情报-IP类型`
28. `IP情报-提供商`
29. `IP情报-公共代理`
30. `IP情报-代理类型`
31. `IP情报-标签`
32. `查询_状态`

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

```text
getiplarkinfo/
├── get_ip_info.py          # 主脚本
├── requirements.txt        # Python 依赖
├── README.md               # 说明文档
├── examples/input.xlsx     # 示例输入文件（需自行准备，不提交真实数据）
└── ip_info_result_*.xlsx   # 生成的结果文件
```

---

## 许可证

MIT License
