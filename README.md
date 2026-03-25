# ip_monitor_pyqt6

一个基于 PyQt6 的桌面公网 IP 监测工具（Windows 优先适配）。

## 功能说明

- 启动后自动检测当前公网 IPv4（双源模式：主源 + 网页对齐源）。
- 查询并展示 IP 详情：IP、国家、地区、城市、ISP、组织、AS、时区。
- 仅当当前 IP 与 Excel 最后一条记录不同，才追加新记录。
- 每次检测后（无论是否变化、是否失败）自动打开 `https://whatismyipaddress.com`。
- 提供图形界面：
  - 当前公网 IP
  - 上次记录 IP
  - ISP
  - 地区/城市
  - 当前状态（未变化 / 已变化并写入Excel / 检测失败）
  - 网页源 IPv4、双源一致性
  - 只读日志区域
  - 按钮：立即检测、打开Excel、打开网站、设置Excel路径、退出
  - 状态栏：Excel 路径、上次检测时间

## 项目结构

```text
ip_monitor_pyqt6/
├─ main.py
├─ requirements.txt
├─ README.md
├─ core/
│  ├─ __init__.py
│  ├─ config_service.py
│  ├─ excel_service.py
│  ├─ ip_service.py
│  └─ logger_service.py
├─ ui/
│  ├─ __init__.py
│  └─ main_window.py
└─ data/
   └─ config.json
```

## 环境要求

- Python 3.10+
- Windows（优先）

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行方式

在项目根目录执行：

```bash
python main.py
```

## 本地运行说明（Windows）

1. 进入项目目录：

```powershell
cd ip_monitor_pyqt6
```

2. 可选：创建并激活虚拟环境：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

3. 安装依赖：

```powershell
pip install -r requirements.txt
```

4. 启动程序：

```powershell
python main.py
```

或使用虚拟环境解释器直接启动：

```powershell
..\.venv\Scripts\python.exe main.py
```

5. 首次运行行为：

- 若 `data/config.json` 不存在，会自动创建默认配置。
- 若 `excel_path` 为空，会自动回退为默认路径 `data/ip_log.xlsx`。
- 若未配置 `ip_check_urls`，程序会自动生成双源默认配置。
- 若 Excel 文件不存在，会自动创建并写入表头。

## Excel 记录说明

- 默认路径：`data/ip_log.xlsx`
- 文件不存在时自动创建。
- 表头字段：
  - 记录时间
  - IP地址
  - 国家
  - 地区
  - 城市
  - ISP
  - 组织
  - AS
  - 时区
  - 是否变化
  - 备注

## 配置文件

配置文件路径：`data/config.json`

示例：

```json
{
  "excel_path": "data/ip_log.xlsx",
  "ip_check_url": "https://api.ipify.org?format=json",
  "ip_check_urls": [
    {
      "name": "主源(ipify)",
      "url": "https://api.ipify.org?format=json"
    },
    {
      "name": "网页对齐源(ifconfig)",
      "url": "https://ifconfig.me/ip"
    }
  ],
  "ip_info_url_template": "http://ip-api.com/json/{ip}?lang=zh-CN",
  "auto_open_url": "https://whatismyipaddress.com"
}
```

## 异常处理

已覆盖以下常见异常：

- 网络请求失败或接口返回异常。
- 配置文件不存在、损坏或读取失败。
- Excel 被占用导致无法读写。
- 未预期异常的弹窗提示与日志输出。

## 常见问题（Windows）

1. 报错 `ImportError: DLL load failed while importing QtWidgets/QtCore`

- 先确认在项目虚拟环境中运行，而不是系统 Python：

```powershell
..\.venv\Scripts\python.exe main.py
```

- 重新安装已验证版本：

```powershell
pip uninstall -y PyQt6 PyQt6-Qt6 PyQt6-sip
pip install --no-cache-dir PyQt6==6.7.1 PyQt6-Qt6==6.7.3 PyQt6-sip==13.8.0
```
