# 主动信息收集工具集

## 项目概述

这是一个用于网络安全测试的**主动信息收集工具集**，主要用于渗透测试前期侦察阶段。工具集整合了端口扫描、Web指纹识别、目录扫描等功能，并提供了自动化处理流水线。

**核心功能:**
- 端口扫描与服务识别
- Web应用指纹识别
- 目录与文件扫描
- 扫描结果自动化处理与报告生成

## 工具组件

### 可执行文件

| 文件 | 用途 |
|------|------|
| `ts.exe` | Tscan端口扫描工具，支持端口、URL、JS等多种扫描模式 |
| `ehole.exe` | Web指纹识别工具，从URL列表中识别应用类型和技术栈 |
| `spray.exe` | 目录扫描工具，对URL进行路径爆破 |

### Python脚本

| 文件 | 用途 |
|------|------|
| `1.py` | 主自动化脚本，协调整个扫描流程（spray + ehole） |
| `2.py` | 端口扫描脚本，执行ts扫描并解析结果生成Excel报告 |
| `ppp.py` | 端口扫描结果解析工具，生成美化的Excel报告 |
| `process_data.py` | 数据处理工具，处理JSON/Excel结果并美化表格 |

## 使用方法

### 快速启动（批处理脚本）

#### 完整扫描流程
```batch
:: 使用top100端口进行完整扫描（端口 + 目录 + 指纹）
轮子top100.bat

:: 使用top1000端口进行完整扫描
轮子top1000.bat
```

#### 仅端口扫描
```batch
:: 快速端口扫描（top100端口）
top100仅端口.bat

:: 深度端口扫描（top1000端口）
top1000仅端口.bat

:: 小字典快速扫描
小字典.bat
```

#### 结果处理
```batch
:: 处理端口扫描结果
端口处理.bat
```

### 手动执行命令

#### 端口扫描
```batch
:: 扫描IP列表的指定端口
ts -hf ip.txt -portf ports.txt -np -m port,url,js

:: 扫描IP列表（自动端口探测）
ts -hf ip.txt -np -m port,url,js
```

参数说明：
- `-hf` - 指定IP列表文件
- `-portf` - 指定端口列表文件
- `-np` - 不进行ping探测
- `-m` - 扫描模式：port(端口), url(URL探测), js(JS分析)

#### 目录扫描
```batch
spray.exe -l url.txt -d dirv2.txt -f res.json
```

#### 指纹识别
```batch
ehole finger -l url.txt -o result.xlsx -t 10
```

## 输入文件

| 文件 | 说明 |
|------|------|
| `ip.txt` | 目标IP地址列表，每行一个IP |
| `ports.txt` / `port.txt` | 端口列表文件 |
| `dirv2.txt` / `dirv3.txt` | 目录扫描字典 |
| `finger.json` | 指纹识别规则库 |

## 输出文件

扫描结果按日期组织到文件夹中（格式：`MMDD`）：

### 端口扫描结果
- `port_scan_report_YYYYMMDD_HHMMSS.xlsx` - 端口扫描详细报告

### 目录扫描结果
- `spray_original_YYYYMMDD.json` - Spray原始输出
- `spray_processed_YYYYMMDD.xlsx` - 处理后的Excel报告
- `YYYYMMDD_status200_urls_N.txt` - 状态码200的URL列表

### 指纹识别结果
- `ehole_result_YYYYMMDD.xlsx` - 指纹识别详细报告

## 工作流程

```
┌─────────────────────────────────────────────────────────┐
│                    完整扫描流程                           │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  1. 端口扫描 (ts.exe)                                    │
│     ip.txt ──→ 扫描 ──→ url.txt, port.txt               │
│                                                         │
│  2. 结果处理 (ppp.py)                                    │
│     port.txt ──→ 解析 ──→ Excel报告                      │
│                                                         │
│  3. 目录扫描 (spray.exe)                                 │
│     url.txt + dirv2.txt ──→ 扫描 ──→ res.json           │
│                                                         │
│  4. 数据处理 (process_data.py)                           │
│     res.json ──→ 处理 ──→ Excel + URL列表                │
│                                                         │
│  5. 指纹识别 (ehole.exe)                                 │
│     URL列表 ──→ 识别 ──→ 指纹报告.xlsx                    │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

## 配置文件

### config.yaml
主要配置项：
- 扫描超时设置 (`Cfgtimeout`, `CfgWebTimeout`)
- 代理配置 (`CfgGlobalProxy`, `CfgProxy`)
- 扫描引擎API密钥 (fofa, hunter, quake, shodan等)
- 端口策略 (`IpSelectedStrategy`: top100/top1000)
- 线程数配置 (`IpThreadStr`, `UrlThreadStr`)

## 依赖项

### Python库
```batch
pip install pandas openpyxl psutil
```

### 必需库
- `pandas` - 数据处理
- `openpyxl` - Excel文件操作
- `psutil` - 进程监控

## 注意事项

1. **文件编码**: 所有文本文件使用UTF-8编码
2. **并发控制**: 线程数可在config.yaml中调整，默认100线程
3. **超时设置**: 可根据网络环境调整超时时间
4. **结果去重**: URL输出会自动去重
5. **进程监控**: 自动化脚本会监控子进程执行状态

## 常见问题

**Q: 扫描结果为空？**
A: 检查ip.txt格式，确保每行一个有效IP地址

**Q: ehole未生成结果？**
A: 可能是快速完成，脚本已处理此情况

**Q: Excel打不开？**
A: 确保安装了openpyxl库：`pip install openpyxl`
