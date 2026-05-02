# IP信息查询工具

从 Excel/CSV 文件或命令行参数读取 IP 地址，自动查询 [iplark.com](https://iplark.com/) 获取详细信息，包括：类型、国家/地区、ASN、企业、使用场景、IP评分与 IP 情报等。

## 功能特点

- 支持多种文件格式：`.csv`、`.xlsx`、`.xls`
- Excel 文件会自动读取全部工作表，工作表不需要命名为 `Sheet1`、`Sheet2`
- 灵活的列选择：支持列字母（A/B/C...）或列名
- 自动生成两个结果文件：
  - 纯查询结果表
  - 原表 + 查询结果追加列（Excel 会保留原有多个工作表及其名称）
- 支持 `-ip` 直接传入一个或多个 IP，不需要准备输入表格
- 支持从历史查询结果中筛选失败 IP 重试，也支持手动指定一个或多个 IP 重试
- 适配 iplark 新版页面结构：
  - “数字地址”会先点击小眼睛显示完整数字，再提取纯数字
  - “国家/地区”会拼接国旗 alt 与文本（例如 `China中国`）
  - 新增“页面顶部标签”和“反查域名”字段，保留 IP 标题下方的子标签和 DNS 反查主机名
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
path\to\chrome-win64\chrome.exe
```

### 2. 下载 ChromeDriver

- 下载地址：https://googlechromelabs.github.io/chrome-for-testing/

**重要**：ChromeDriver 版本必须与 Chrome 版本一致！

下载后解压到任意目录，例如：
```
path\to\chromedriver-win64\chromedriver.exe
```

### 3. 查看 Chrome 版本

打开 Chrome，访问 `chrome://version/`，查看版本号（如 `130.0.6723.58`）。

### 4. 修改脚本中的路径

打开 `get_ip_info.py`，找到 `setup_driver` 函数，修改以下两行：

```python
def setup_driver():
    """配置并启动Chrome浏览器"""
    chrome_driver_path = r''  # 可选：留空时让 Selenium 自动查找 ChromeDriver
    chrome_binary_path = r''  # 可选：仅在使用独立 Chrome 时填写浏览器路径
```

---

## 三、使用方法

### 1. 准备输入文件和列

推荐运行时用 `-i` 或 `--input` 传入输入文件：

```powershell
python get_ip_info.py -i examples/input.xlsx
```

路径包含空格时请加引号：

```powershell
python get_ip_info.py -i "examples/input.xlsx"
```

如果不传 `-i/--input`，脚本会使用 `get_ip_info.py` 中 `main` 函数配置区域里的 `INPUT_FILE`。命令行传入的文件路径优先级高于脚本配置。

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

或直接指定输入文件：

```powershell
python get_ip_info.py -i examples/input.xlsx
```

如果系统中存在多个 Python 环境，可以指定目标解释器运行：

```powershell
& "<Python环境路径>\python.exe" get_ip_info.py -i examples/input.xlsx
```

### 3. 直接查询命令行 IP

不需要输入表格时，可以用 `-ip` 直接传入一个或多个 IP，脚本会生成纯查询结果 Excel：

```powershell
python get_ip_info.py -ip 1.2.3.4
```

多个 IP 可以放在同一个 `-ip` 后面：

```powershell
python get_ip_info.py -ip 1.2.3.4 2.3.4.5 3.4.5.6
```

也支持逗号、分号或带引号的空白分隔：

```powershell
python get_ip_info.py -ip "1.2.3.4,2.3.4.5;3.4.5.6"
```

指定输出目录：

```powershell
python get_ip_info.py -ip 1.2.3.4 2.3.4.5 -o results
```

### 4. 查看结果

使用输入文件时，脚本会在输出目录生成两个文件：

| 文件 | 说明 |
|------|------|
| `ip_info_result_2025-12-21-234310_UTC+8.xlsx` | 纯查询结果，包含 `查询结果` 和 `输出字段说明` 两个工作表 |
| `原文件名_ip_info_result_2025-12-21-234310_UTC+8.xlsx` | 原表内容 + 追加的查询结果列；Excel 会按原工作表名分别写回 |

使用 `-ip` 直接查询时，只会生成 `ip_info_result_2025-12-21-234310_UTC+8.xlsx` 纯查询结果文件。

---

## 四、失败 IP 重试

如果一次查询只有少量 IP 失败，可以基于历史结果文件重试失败项，避免重新查询已经成功的 IP。

`--retry-from` 支持两类历史文件：

- 纯查询结果文件，例如 `ip_info_result_2025-12-21-234310_UTC+8.xlsx`，读取 `IP` 和 `查询状态`。
- 原表回填结果文件，例如 `input_ip_info_result_2025-12-21-234310_UTC+8.xlsx`，从原始 IP 列读取 IP，并读取 `查询_状态` 等回填列。

历史结果会优先读取 `查询结果` 工作表；如果没有该工作表，则读取所有数据工作表。原表回填结果中的 IP 列会按 `--ip-column` / `IP_COLUMN` 配置识别；未指定时，每个工作表会自动扫描原始列内容，只提取单元格值为公网 IPv4 的行。

### 1. 查看会重试哪些 IP

`--dry-run` 只列出目标 IP，不启动浏览器，也不会写入 Excel：

```powershell
python get_ip_info.py --retry-from ip_info_result_2025-12-21-234310_UTC+8.xlsx --dry-run
```

默认只选择 `查询状态` 不是 `成功` 的 IP；空状态也会按失败处理。

### 2. 自动重试历史失败项

```powershell
python get_ip_info.py --retry-from ip_info_result_2025-12-21-234310_UTC+8.xlsx
```

重试模式会生成：

| 文件 | 说明 |
|------|------|
| `ip_info_retry_2025-12-21-234310_UTC+8.xlsx` | 历史结果与本次重试结果合并后的完整 retry 结果；本次结果会覆盖同 IP 的历史记录 |

如果传入的历史文件带有原始文件名前缀，retry 文件也会保留相同前缀。例如 `test_ip_info_result_2025-12-21-234310_UTC+8.xlsx` 会生成 `test_ip_info_retry_2025-12-21-235000_UTC+8.xlsx`。

如果 `--retry-from` 是原表回填结果文件，`ip_info_retry` 输出也会保持原表回填格式，包含之前成功的结果和本次 retry 后的结果；如果 `--retry-from` 是纯查询结果文件，则输出为纯查询结果格式。

### 3. 手动指定重试 IP

指定一个 IP：

```powershell
python get_ip_info.py --retry-ip 1.2.3.4
```

指定多个 IP：

```powershell
python get_ip_info.py --retry-ip 1.2.3.4 --retry-ip 5.6.7.8
```

也可以使用逗号、分号或空白分隔：

```powershell
python get_ip_info.py --retry-ips "1.2.3.4,5.6.7.8"
```

如果同时传入 `--retry-from` 和手动 IP，脚本只会从历史失败项中重试指定 IP。历史结果中已经成功的 IP 默认跳过；需要强制重查时加 `--force`：

```powershell
python get_ip_info.py --retry-from ip_info_result_2025-12-21-234310_UTC+8.xlsx --retry-ip 1.2.3.4 --force
```

### 4. 从纯查询结果生成原表 retry 回填文件

如果 `--retry-from` 使用的是纯查询结果文件，并且需要得到“原表 + retry 后查询结果追加列”的完整文件，需要同时传入原始输入文件：

```powershell
python get_ip_info.py --retry-from ip_info_result_2025-12-21-234310_UTC+8.xlsx -i examples/input.xlsx
```

此时会额外生成：

| 文件 | 说明 |
|------|------|
| `input_ip_info_retry_2025-12-21-234310_UTC+8.xlsx` | 原始输入全部工作表 + retry 后查询结果追加列 |

可选参数：

| 参数 | 说明 |
|------|------|
| `--dry-run` | 只列出目标 IP，不启动浏览器，不写入文件 |
| `--force` | 即使历史结果中 IP 已成功，也重新查询 |
| `-o/--output-dir` | 指定输出目录 |
| `--ip-column` | 命令行指定原始输入文件的 IP 列，优先级高于脚本内 `IP_COLUMN` |

---

## 五、根据不同表格修改配置

### IP_COLUMN 配置说明

Excel 文件包含多个工作表时，脚本会逐个工作表识别 IP 列并提取 IP。相同 IP 即使出现在多个工作表，也只会查询一次，然后把结果回填到所有出现过的行。自动检测模式下，不依赖表头是否叫 `IP`；只要某列中出现公网 IPv4 地址，该列就会被识别。局域网、回环、链路本地、组播、保留地址、运营商级 NAT 等非公网地址会跳过，不查询 iplark。没有有效公网 IP 的工作表会保持原样，不追加查询结果列。

| 值 | 说明 | 示例 |
|----|------|------|
| `None` | 每个工作表自动检测（扫描各列内容，提取公网 IPv4 地址） | `IP_COLUMN = None` |
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

**示例3**：让脚本按单元格内容自动识别 IP 列
```python
IP_COLUMN = None
```

---

## 六、输出字段说明

纯查询结果文件包含 `查询结果` 和 `输出字段说明` 两个工作表。

`输出字段说明` 工作表会由脚本自动生成 `序号`、`列号`、`字段`、`说明` 四列。`查询结果` 工作表默认包含以下字段；如果页面实际出现不在标准 14 源中的地理位置来源，会在 `地理位置-Cloudflare` 后自动追加实际来源列，例如 `地理位置-IPLark`、`地理位置-CZ88`、`地理位置-Leak`，后续字段列号会相应后移。

| 序号 | 列号 | 字段 | 说明 |
|------|------|------|------|
| 1 | A | IP | IP地址 |
| 2 | B | 页面顶部标签 | 页面 IP 标题下方的子标签，按页面顺序用分号连接 |
| 3 | C | 反查域名 | 页面顶部子标签中看起来像 DNS 反查主机名的值 |
| 4 | D | 类型 | 家宽、数据中心、商宽等 |
| 5 | E | 使用场景 | 网页“使用场景”字段（例如：普通宽带） |
| 6 | F | IP属性 | 原生IP、广播IP |
| 7 | G | 国家/地区 | 所属国家或地区 |
| 8 | H | 地理位置-Ip-api | Ip-api 来源的地理位置 |
| 9 | I | 地理位置-Moe | Moe 来源的地理位置 |
| 10 | J | 地理位置-Moe+ | Moe+ 来源的地理位置 |
| 11 | K | 地理位置-Ease | Ease 来源的地理位置 |
| 12 | L | 地理位置-Internet | Internet 来源的地理位置 |
| 13 | M | 地理位置-Maxmind | Maxmind 来源的地理位置 |
| 14 | N | 地理位置-Ipstack | Ipstack 来源的地理位置 |
| 15 | O | 地理位置-IPinfo | IPinfo 来源的地理位置 |
| 16 | P | 地理位置-IP2Location | IP2Location 来源的地理位置 |
| 17 | Q | 地理位置-Digital Element | Digital Element 来源的地理位置 |
| 18 | R | 地理位置-DB-IP | DB-IP 来源的地理位置 |
| 19 | S | 地理位置-Aliyun | Aliyun 来源的地理位置 |
| 20 | T | 地理位置-TencentCloud | TencentCloud 来源的地理位置 |
| 21 | U | 地理位置-Cloudflare | Cloudflare 来源的地理位置 |
| 22 | V | ASN | 自治系统编号 |
| 23 | W | 企业 | 所属企业/运营商 |
| 24 | X | IP评分 | 0-100分 |
| 25 | Y | IP情报-使用类型 | IP情报区域中的“使用类型”（保留 `-`） |
| 26 | Z | IP情报-威胁 | IP情报区域中的“威胁”（保留 `-`） |
| 27 | AA | IP情报-IP类型 | IP情报区域中的“IP类型”（保留 `-`） |
| 28 | AB | IP情报-提供商 | IP情报区域中的“提供商”（保留 `-`） |
| 29 | AC | IP情报-公共代理 | IP情报区域中的“公共代理”（保留 `-`） |
| 30 | AD | IP情报-代理类型 | IP情报区域中的“代理类型”（保留 `-`） |
| 31 | AE | IP情报-标签 | IP情报区域中的“标签”（保留 `-`） |
| 32 | AF | 数字地址 | 会先点击小眼睛显示完整数字，再提取纯数字 |
| 33 | AG | 备注 | ASN 规模等补充说明 |
| 34 | AH | 查询状态 | 成功/超时/错误 |

原表追加列（`原文件名_ip_info_result_*.xlsx`）的顺序为：

1. `查询IP`（本行回填结果对应的实际查询 IP）
2. `查询_页面顶部标签`
3. `查询_反查域名`
4. `查询_类型`
5. `查询_使用场景`（紧跟在 `查询_类型` 后）
6. `查询_IP属性`
7. `查询_国家地区`
8. `查询_地理位置-Ip-api`
9. `查询_地理位置-Moe`
10. `查询_地理位置-Moe+`
11. `查询_地理位置-Ease`
12. `查询_地理位置-Internet`
13. `查询_地理位置-Maxmind`
14. `查询_地理位置-Ipstack`
15. `查询_地理位置-IPinfo`
16. `查询_地理位置-IP2Location`
17. `查询_地理位置-Digital Element`
18. `查询_地理位置-DB-IP`
19. `查询_地理位置-Aliyun`
20. `查询_地理位置-TencentCloud`
21. `查询_地理位置-Cloudflare`
22. `查询_地理位置-其他来源`（页面实际出现但不在标准 14 源中的来源会自动追加，例如 `查询_地理位置-IPLark`）
23. `查询_ASN`
24. `查询_企业`
25. `查询_IP评分`
26. `查询_数字地址`
27. `查询_备注`
28. `IP情报-使用类型`
29. `IP情报-威胁`
30. `IP情报-IP类型`
31. `IP情报-提供商`
32. `IP情报-公共代理`
33. `IP情报-代理类型`
34. `IP情报-标签`
35. `查询_状态`

---

## 七、常见问题

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

## 八、文件结构

```text
getiplarkinfo/
├── get_ip_info.py          # 主脚本
├── requirements.txt        # Python 依赖
├── README.md               # 说明文档
├── examples/
│   └── input.xlsx           # 示例输入文件（自行创建或替换）
└── ip_info_result_*.xlsx    # 生成的结果文件（不要提交）
```

---

## 许可证

MIT License
