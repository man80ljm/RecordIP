# 公网 IP 监测工具（PyQt6）

一个面向 Windows 的桌面小工具，用于持续监测公网 IPv4 变化，并将变化记录到 Excel。

## 功能亮点

- 双源 IP 检测：
  - 主源：ipify
  - 网页对齐源：ifconfig（通常更接近 whatismyipaddress 的展示结果）
- 自动监测：按配置间隔轮询（默认 5 分钟）
- 变化入库：只要任一来源 IP 变化，即写入 Excel 新记录
- 手动检测：点击“立即检测”可立刻检测，并自动打开配置网站
- 定时检测：后台静默检测，不自动打开网站
- 托盘运行：点击窗口关闭按钮后最小化到系统托盘，右键托盘图标可退出程序
- 归属地信息：国家、地区、城市、ISP、组织、AS、时区

## 技术栈

- Python 3.10+
- PyQt6 6.7.1（已验证稳定）
- requests
- openpyxl

## 项目结构

```text
ip_monitor_pyqt6/
├─ main.py
├─ requirements.txt
├─ READ.md
├─ README.md
├─ core/
│  ├─ config_service.py
│  ├─ excel_service.py
│  ├─ ip_service.py
│  └─ logger_service.py
├─ ui/
│  └─ main_window.py
├─ data/
│  └─ config.json
└─ dist/
```

## 安装与运行

在项目目录执行：

```powershell
pip install -r requirements.txt
python main.py
```

如果你使用本地虚拟环境：

```powershell
..\.venv\Scripts\python.exe main.py
```

## 配置说明

配置文件路径：data/config.json

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
  "auto_open_url": "https://whatismyipaddress.com",
  "auto_detect_interval_minutes": 5
}
```

说明：

- auto_detect_interval_minutes > 0：开启定时检测
- auto_detect_interval_minutes = 0：仅手动检测

## Excel 记录格式

默认文件：data/ip_log.xlsx

表头字段共 13 列：

1. 记录时间
2. 主源IP
3. 网页源IP
4. 国家
5. 地区
6. 城市
7. ISP
8. 组织
9. AS
10. 时区
11. 双源一致
12. 是否变化
13. 备注

## 行为规则

- 程序启动后会自动检测一次
- 定时检测仅后台执行，不弹浏览器
- 点击“立即检测”会执行检测并打开网站
- 点击窗口右上角关闭：不会退出程序，而是缩到系统托盘
- 需要真正退出时：右键托盘图标，选择“退出程序”

## 打包（Windows / PyInstaller）

单文件（onefile）：

```powershell
pyinstaller --clean -F -w --icon="..\ip.ico" --add-data "..\ip.ico;." --name "IP监测工具" main.py
```

输出文件：dist/IP监测工具.exe

## 常见问题

### 1. PyQt6 导入失败（QtWidgets/QtCore DLL）

建议使用已验证版本：

```powershell
pip uninstall -y PyQt6 PyQt6-Qt6 PyQt6-sip
pip install --no-cache-dir PyQt6==6.7.1 PyQt6-Qt6==6.7.3 PyQt6-sip==13.8.0
```

### 2. 任务管理器图标未刷新

可能是 Windows 图标缓存问题。可尝试：

- 关闭所有同名进程后重新启动
- 重启资源管理器或系统
- 使用包含多尺寸图层的 ico（16/24/32/48/64/128/256）

## License

仅用于学习与个人使用。可按你的需要补充正式开源协议（如 MIT）。
