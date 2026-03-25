from typing import Any
import re

import requests


class IPServiceError(Exception):
    """IP 查询相关异常。"""


class IPService:
    """IP 服务，负责获取公网 IP 及其详细信息。"""

    def __init__(self, timeout: int = 10):
        self.timeout = timeout

    def fetch_current_ipv4(self, ip_check_url: str) -> str:
        """获取当前公网 IPv4。"""
        try:
            response = requests.get(ip_check_url, timeout=self.timeout)
            response.raise_for_status()
        except requests.RequestException as exc:
            raise IPServiceError(f"获取公网 IP 失败: {exc}") from exc

        ip = self._extract_ip_from_response(response)
        if not ip:
            raise IPServiceError("公网 IP 为空")

        return ip

    def fetch_dual_source_ipv4(self, source_configs: list[dict[str, str]]) -> dict[str, Any]:
        """双源模式获取公网 IPv4，返回主源、网页源与一致性。"""
        if not source_configs or len(source_configs) < 2:
            raise IPServiceError("双源配置不足，至少需要两个来源")

        primary = source_configs[0]
        web_like = source_configs[1]

        primary_name = str(primary.get("name", "主源")).strip() or "主源"
        primary_url = str(primary.get("url", "")).strip()
        web_name = str(web_like.get("name", "网页对齐源")).strip() or "网页对齐源"
        web_url = str(web_like.get("url", "")).strip()

        if not primary_url or not web_url:
            raise IPServiceError("双源配置URL为空")

        primary_ip = self.fetch_current_ipv4(primary_url)
        web_ip = self.fetch_current_ipv4(web_url)

        return {
            "primary_name": primary_name,
            "primary_url": primary_url,
            "primary_ip": primary_ip,
            "web_name": web_name,
            "web_url": web_url,
            "web_ip": web_ip,
            "is_consistent": (primary_ip == web_ip),
        }

    def _extract_ip_from_response(self, response: requests.Response) -> str:
        """从响应中解析 IP，兼容 JSON 和纯文本。"""
        content_type = response.headers.get("Content-Type", "")
        if "application/json" in content_type:
            data = response.json()
            ip = str(data.get("ip", "")).strip()
            if ip:
                return ip

            # 兼容一些接口返回字段不叫 ip 的情况。
            for key in ("query", "ipAddress", "origin"):
                value = str(data.get(key, "")).strip()
                if value:
                    match = re.search(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", value)
                    if match:
                        return match.group(0)
            return ""

        # 兼容纯文本接口，如 "当前 IP：1.2.3.4 来自于..."。
        text = response.text.strip()
        match = re.search(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", text)
        if match:
            return match.group(0)
        return text

    def fetch_ip_info(self, ip_info_url_template: str, ip: str) -> dict[str, Any]:
        """根据 IP 获取归属地和运营商信息。"""
        if "{ip}" not in ip_info_url_template:
            raise IPServiceError("IP 信息 URL 模板缺少 {ip} 占位符")

        url = ip_info_url_template.format(ip=ip)

        try:
            response = requests.get(url, timeout=self.timeout)
            response.raise_for_status()
            data = response.json()
        except requests.RequestException as exc:
            raise IPServiceError(f"获取 IP 详细信息失败: {exc}") from exc
        except ValueError as exc:
            raise IPServiceError("IP 详细信息返回内容不是有效 JSON") from exc

        # 兼容 ip-api 风格字段。
        return {
            "ip": ip,
            "country": str(data.get("country", "")),
            "region": str(data.get("regionName", data.get("region", ""))),
            "city": str(data.get("city", "")),
            "isp": str(data.get("isp", "")),
            "org": str(data.get("org", "")),
            "as": str(data.get("as", "")),
            "timezone": str(data.get("timezone", "")),
        }
