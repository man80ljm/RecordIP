import json
from pathlib import Path
from typing import Any


class ConfigService:
    """配置服务，负责加载和保存配置。"""

    DEFAULT_CONFIG = {
        "excel_path": "data/ip_log.xlsx",
        # 保留旧字段，兼容历史配置。
        "ip_check_url": "https://api.ipify.org?format=json",
        # 双源模式：主源 + 网页对齐源。
        "ip_check_urls": [
            {"name": "主源(ipify)", "url": "https://api.ipify.org?format=json"},
            {"name": "网页对齐源(ifconfig)", "url": "https://ifconfig.me/ip"},
        ],
        # 国内可访问回退源（当默认双源不可达时使用）。
        "ip_check_urls_cn": [
            {"name": "国内主源(3322)", "url": "https://ip.3322.net"},
            {"name": "国内网页源(ipip)", "url": "https://myip.ipip.net"},
        ],
        # 是否启用国内源自动回退。
        "enable_cn_fallback": True,
        "ip_info_url_template": "http://ip-api.com/json/{ip}?lang=zh-CN",
        "auto_open_url": "https://whatismyipaddress.com",
        # 国内 IP 时跳转的网站。
        "auto_open_url_cn": "https://myip.ipip.net/",
        # 延迟测试目标 URL（默认 Google，适合 VPN 用户）。
        "latency_test_url": "https://www.google.com",
        # 国内 IP 时的延迟测试目标。
        "latency_test_url_cn": "https://www.baidu.com",
        # 自动检测间隔（分钟），0 表示禁用自动检测。
        "auto_detect_interval_minutes": 5,
    }

    def __init__(self, project_root: Path):
        self.project_root = project_root
        self.data_dir = self.project_root / "data"
        self.config_path = self.data_dir / "config.json"

    def load_config(self) -> dict[str, Any]:
        """加载配置文件，不存在时自动创建默认配置。"""
        self.data_dir.mkdir(parents=True, exist_ok=True)

        if not self.config_path.exists():
            config = dict(self.DEFAULT_CONFIG)
            self.save_config(config)
            return config

        try:
            with self.config_path.open("r", encoding="utf-8") as f:
                user_config = json.load(f)
        except (json.JSONDecodeError, OSError):
            # 配置文件损坏或读取失败时回退到默认配置。
            config = dict(self.DEFAULT_CONFIG)
            self.save_config(config)
            return config

        config = dict(self.DEFAULT_CONFIG)
        config.update(user_config if isinstance(user_config, dict) else {})

        # 空字符串配置回退到默认值，避免运行时路径或地址为空。
        for key, default_value in self.DEFAULT_CONFIG.items():
            value = config.get(key)
            if value is None:
                config[key] = default_value
                continue
            if isinstance(default_value, str) and str(value).strip() == "":
                config[key] = default_value

        # 兼容旧版配置: 当未配置双源时，自动由旧字段生成。
        if not self._is_valid_source_list(config.get("ip_check_urls")):
            legacy_url = str(config.get("ip_check_url", self.DEFAULT_CONFIG["ip_check_url"])).strip()
            if not legacy_url:
                legacy_url = self.DEFAULT_CONFIG["ip_check_url"]
            config["ip_check_urls"] = [
                {"name": "主源(ipify)", "url": legacy_url},
                {"name": "网页对齐源(ifconfig)", "url": "https://ifconfig.me/ip"},
            ]

        # 兼容旧版配置: 当未配置国内回退双源时，自动补齐默认值。
        if not self._is_valid_source_list(config.get("ip_check_urls_cn")):
            config["ip_check_urls_cn"] = list(self.DEFAULT_CONFIG["ip_check_urls_cn"])

        # 自动回写，确保新增字段被持久化。
        self.save_config(config)
        return config

    def save_config(self, config: dict[str, Any]) -> None:
        """保存配置文件。"""
        self.data_dir.mkdir(parents=True, exist_ok=True)
        with self.config_path.open("w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def _is_valid_source_list(self, source_list: Any) -> bool:
        """校验双源配置结构是否有效。"""
        if not isinstance(source_list, list) or len(source_list) < 2:
            return False
        for item in source_list:
            if not isinstance(item, dict):
                return False
            url = str(item.get("url", "")).strip()
            if not url:
                return False
        return True

    def get_excel_path(self, config: dict[str, Any]) -> Path:
        """获取 Excel 绝对路径。支持相对路径与绝对路径。"""
        raw_path = str(config.get("excel_path", "")).strip()
        if not raw_path:
            raw_path = self.DEFAULT_CONFIG["excel_path"]
        excel_path = Path(raw_path)
        if not excel_path.is_absolute():
            excel_path = self.project_root / excel_path
        return excel_path
