import os
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QObject, QThread, QTimer, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QStatusBar,
    QSystemTrayIcon,
    QMenu,
    QVBoxLayout,
    QWidget,
)

from core.config_service import ConfigService
from core.excel_service import ExcelService, ExcelServiceError
from core.ip_service import IPService, IPServiceError
from core.logger_service import LoggerService


class DetectWorker(QObject):
    """后台执行检测逻辑，避免阻塞 UI 主线程。"""

    succeeded = pyqtSignal(dict)
    failed = pyqtSignal(str, bool)
    finished = pyqtSignal()

    def __init__(
        self,
        *,
        excel_path: Path,
        source_configs: list[dict[str, str]],
        fallback_source_configs: list[dict[str, str]] | None,
        ip_info_url_template: str,
        ip_service: IPService,
        excel_service: ExcelService,
    ):
        super().__init__()
        self.excel_path = excel_path
        self.source_configs = source_configs
        self.fallback_source_configs = fallback_source_configs
        self.ip_info_url_template = ip_info_url_template
        self.ip_service = ip_service
        self.excel_service = excel_service

    @pyqtSlot()
    def run(self) -> None:
        """在线程中执行检测并通过信号返回结果。"""
        try:
            logs: list[str] = []
            used_fallback = False
            try:
                dual_result = self.ip_service.fetch_dual_source_ipv4(self.source_configs)
            except IPServiceError as primary_exc:
                if not self.fallback_source_configs:
                    raise
                logs.append(f"默认源不可达，尝试国内源回退: {primary_exc}")
                dual_result = self.ip_service.fetch_dual_source_ipv4(self.fallback_source_configs)
                used_fallback = True

            primary_ip = dual_result["primary_ip"]
            web_ip = dual_result["web_ip"]
            is_consistent = bool(dual_result["is_consistent"])

            # IP 信息查询失败时不终止主流程，优先保证 IP 检测可用。
            try:
                ip_info = self.ip_service.fetch_ip_info(self.ip_info_url_template, web_ip)
            except IPServiceError as info_exc:
                logs.append(f"IP 详情查询失败，已忽略: {info_exc}")
                ip_info = {
                    "country": "",
                    "region": "",
                    "city": "",
                    "isp": "",
                    "org": "",
                    "as": "",
                    "timezone": "",
                }

            last_ips = self.excel_service.get_last_ips(self.excel_path)
            last_primary = last_ips.get("primary_ip", "")
            last_web = last_ips.get("web_ip", "")

            # 任意来源发生变化时，均写入 Excel 作为新记录。
            changed = (primary_ip != last_primary) or (web_ip != last_web)

            if changed:
                is_first = not last_primary and not last_web
                note_prefix = "首次记录" if is_first else "IP发生变化"
                consistency_text = "一致" if is_consistent else "不一致"
                note = f"{note_prefix}; 主源={primary_ip}; 网页源={web_ip}; 双源{consistency_text}"
                record = {
                    "记录时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "主源IP": primary_ip,
                    "网页源IP": web_ip,
                    "国家": ip_info.get("country", ""),
                    "地区": ip_info.get("region", ""),
                    "城市": ip_info.get("city", ""),
                    "ISP": ip_info.get("isp", ""),
                    "组织": ip_info.get("org", ""),
                    "AS": ip_info.get("as", ""),
                    "时区": ip_info.get("timezone", ""),
                    "双源一致": "是" if is_consistent else "否",
                    "是否变化": "是",
                    "备注": note,
                }
                self.excel_service.append_record(self.excel_path, record)
                status_text = "已变化并写入Excel"
                if is_first:
                    logs.append(f"首次记录: 主源={primary_ip} | 网页源={web_ip}")
                else:
                    logs.append(
                        f"IP 变化: 主源 {last_primary or '-'} → {primary_ip} | "
                        f"网页源 {last_web or '-'} → {web_ip}"
                    )
                logs.append("已写入新记录到 Excel")
                display_last_web = web_ip
            else:
                status_text = "未变化"
                logs.append(f"IP 未变化: 主源={primary_ip} | 网页源={web_ip}")
                display_last_web = last_web

            logs.append(
                f"双源: {dual_result['primary_name']}={primary_ip} | "
                f"{dual_result['web_name']}={web_ip} | "
                f"一致: {'是' if is_consistent else '否'}"
            )
            if used_fallback:
                logs.append("本次检测已自动切换到国内源")

            payload: dict[str, Any] = {
                "status_text": status_text,
                "primary_ip": primary_ip,
                "web_ip": web_ip,
                "display_last_web": display_last_web,
                "isp": ip_info.get("isp", "-") or "-",
                "region": ip_info.get("region", "") or "-",
                "city": ip_info.get("city", "") or "-",
                "logs": logs,
            }
            self.succeeded.emit(payload)
        except (IPServiceError, ExcelServiceError, ValueError) as exc:
            self.failed.emit(str(exc), False)
        except Exception as exc:
            self.failed.emit(f"发生未预期错误: {exc}", True)
        finally:
            self.finished.emit()


class MainWindow(QMainWindow):
    """主界面窗口。"""

    def __init__(self, project_root: Path):
        super().__init__()

        self.project_root = project_root
        self.config_service = ConfigService(project_root=project_root)
        self.config = self.config_service.load_config()

        self.ip_service = IPService(timeout=10)
        self.excel_service = ExcelService()
        self.logger = LoggerService()
        self.logger.add_listener(self.append_log)

        self.last_check_time_text = "未检测"
        self._allow_close = False
        self._tray_tip_shown = False
        self.tray_icon: QSystemTrayIcon | None = None
        self._is_detecting = False
        self._pending_open_site = False
        self._detect_thread: QThread | None = None
        self._detect_worker: DetectWorker | None = None

        self.setWindowTitle("公网 IP 监测工具 - ip_monitor_pyqt6")
        app_icon = QApplication.windowIcon()
        if not app_icon.isNull():
            self.setWindowIcon(app_icon)
        self.setMinimumSize(820, 560)
        self.resize(820, 560)

        self._init_ui()
        self._apply_styles()
        self._init_tray_icon()
        self._set_status_text("未检测")
        self._refresh_statusbar()

        # 自动检测定时器（仅初始化，由 _setup_auto_timer 决定是否启动）。
        self.auto_detect_timer = QTimer(self)
        self.auto_detect_timer.timeout.connect(self.detect_now)
        self._setup_auto_timer()

        # 启动后自动执行一次检测。
        QTimer.singleShot(300, self.detect_now)

    def _init_ui(self) -> None:
        """初始化界面布局与控件。"""
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 14, 16, 12)
        main_layout.setSpacing(12)

        info_group = QGroupBox("网络概览")
        info_layout = QVBoxLayout(info_group)
        info_layout.setContentsMargins(12, 14, 12, 12)
        info_layout.setSpacing(10)

        # 第一层级：网页源IP 卡片 | 主源IP 卡片 | 状态卡片。
        top_cards_layout = QHBoxLayout()
        top_cards_layout.setSpacing(10)

        # 网页对齐源 IP 卡片（与 whatismyipaddress 最接近）。
        web_ip_card = QFrame()
        web_ip_card.setObjectName("primaryCard")
        web_ip_card_layout = QVBoxLayout(web_ip_card)
        web_ip_card_layout.setContentsMargins(14, 12, 14, 12)
        web_ip_card_layout.setSpacing(4)

        web_ip_title = QLabel("网页对齐源 IPv4")
        web_ip_title.setObjectName("cardTitle")
        self.current_ip_value = QLabel("-")
        self.current_ip_value.setObjectName("currentIpValue")

        web_ip_card_layout.addWidget(web_ip_title)
        web_ip_card_layout.addWidget(self.current_ip_value)

        # 主源 IP 卡片（路由参考源）。
        primary_ref_card = QFrame()
        primary_ref_card.setObjectName("infoCard")
        primary_ref_card_layout = QVBoxLayout(primary_ref_card)
        primary_ref_card_layout.setContentsMargins(14, 12, 14, 12)
        primary_ref_card_layout.setSpacing(4)

        primary_ref_title = QLabel("主源 IPv4")
        primary_ref_title.setObjectName("cardTitle")
        self.primary_ip_value = QLabel("-")
        self.primary_ip_value.setObjectName("primaryIpValue")

        primary_ref_card_layout.addWidget(primary_ref_title)
        primary_ref_card_layout.addWidget(self.primary_ip_value)

        # 当前状态卡片。
        status_card = QFrame()
        status_card.setObjectName("statusCard")
        status_card_layout = QVBoxLayout(status_card)
        status_card_layout.setContentsMargins(14, 12, 14, 12)
        status_card_layout.setSpacing(8)

        status_title = QLabel("当前状态")
        status_title.setObjectName("cardTitle")
        self.status_value = QLabel("未检测")
        self.status_value.setObjectName("statusBadge")
        self.status_value.setAlignment(self.status_value.alignment().AlignCenter)

        status_card_layout.addWidget(status_title)
        status_card_layout.addWidget(self.status_value)
        status_card_layout.addStretch(1)

        top_cards_layout.addWidget(web_ip_card, 5)
        top_cards_layout.addWidget(primary_ref_card, 4)
        top_cards_layout.addWidget(status_card, 3)

        # 第二层级：辅助信息卡片。
        bottom_cards_layout = QGridLayout()
        bottom_cards_layout.setHorizontalSpacing(10)
        bottom_cards_layout.setVerticalSpacing(10)

        self.last_ip_value = QLabel("-")
        self.last_ip_value.setObjectName("cardValue")
        self.isp_value = QLabel("-")
        self.isp_value.setObjectName("cardValue")
        self.location_value = QLabel("-")
        self.location_value.setObjectName("cardValue")

        bottom_cards_layout.addWidget(self._create_info_card("上次记录 IP", self.last_ip_value), 0, 0)
        bottom_cards_layout.addWidget(self._create_info_card("ISP", self.isp_value), 0, 1)
        bottom_cards_layout.addWidget(self._create_info_card("地区 / 城市", self.location_value), 0, 2)

        info_layout.addLayout(top_cards_layout)
        info_layout.addLayout(bottom_cards_layout)

        log_group = QGroupBox("日志输出")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(10, 14, 10, 10)
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMinimumHeight(180)
        self.log_output.setMaximumHeight(230)
        log_layout.addWidget(self.log_output)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        self.detect_btn = QPushButton("立即检测")
        self.detect_btn.setObjectName("primaryButton")
        self.open_excel_btn = QPushButton("打开Excel")
        self.open_excel_btn.setObjectName("secondaryButton")
        self.open_site_btn = QPushButton("打开网站")
        self.open_site_btn.setObjectName("secondaryButton")
        self.set_excel_btn = QPushButton("设置Excel路径")
        self.set_excel_btn.setObjectName("secondaryButton")
        self.exit_btn = QPushButton("退出")
        self.exit_btn.setObjectName("secondaryButton")
        self.interval_label = QLabel("自动(分)")
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(0, 1440)
        self.interval_spin.setSingleStep(1)
        self.interval_spin.setValue(int(self.config.get("auto_detect_interval_minutes", 5)))
        self.interval_spin.setFixedHeight(36)
        self.interval_spin.setFixedWidth(92)
        self.interval_spin.setToolTip("0 表示关闭自动检测")

        self.detect_btn.setFixedHeight(36)
        self.open_excel_btn.setFixedHeight(36)
        self.open_site_btn.setFixedHeight(36)
        self.set_excel_btn.setFixedHeight(36)
        self.exit_btn.setFixedHeight(36)
        self.interval_label.setFixedHeight(36)

        self.detect_btn.setMinimumWidth(140)
        self.open_excel_btn.setMinimumWidth(116)
        self.open_site_btn.setMinimumWidth(116)
        self.set_excel_btn.setMinimumWidth(132)
        self.exit_btn.setMinimumWidth(96)

        self.detect_btn.clicked.connect(lambda: self.detect_now(open_site=True))
        self.open_excel_btn.clicked.connect(self.open_excel)
        self.open_site_btn.clicked.connect(self.open_site)
        self.set_excel_btn.clicked.connect(self.set_excel_path)
        self.interval_spin.valueChanged.connect(self.set_auto_detect_interval)
        self.exit_btn.clicked.connect(self.close)

        button_layout.addWidget(self.detect_btn)
        button_layout.addWidget(self.open_excel_btn)
        button_layout.addWidget(self.open_site_btn)
        button_layout.addWidget(self.set_excel_btn)
        button_layout.addWidget(self.interval_label)
        button_layout.addWidget(self.interval_spin)
        button_layout.addStretch(1)
        button_layout.addWidget(self.exit_btn)

        main_layout.addWidget(info_group)
        main_layout.addWidget(log_group)
        main_layout.addLayout(button_layout)
        main_layout.addStretch(1)

        status_bar = QStatusBar(self)
        self.setStatusBar(status_bar)
        self.status_excel_label = QLabel("Excel路径: -")
        self.status_time_label = QLabel("上次检测时间: 未检测")
        status_bar.addWidget(self.status_excel_label, 1)
        status_bar.addPermanentWidget(self.status_time_label)

    def _create_info_card(self, title: str, value_label: QLabel) -> QFrame:
        """构建统一样式的信息卡片。"""
        card = QFrame()
        card.setObjectName("infoCard")

        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        title_label = QLabel(title)
        title_label.setObjectName("cardTitle")

        layout.addWidget(title_label)
        layout.addWidget(value_label)

        return card

    def _apply_styles(self) -> None:
        """应用更紧凑、现代的桌面工具风格。"""
        self.setStyleSheet(
            """
            QWidget {
                font-size: 13px;
                color: #1f2937;
                font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", sans-serif;
            }
            QGroupBox {
                font-weight: 600;
                border: 1px solid #d6dde7;
                border-radius: 10px;
                margin-top: 8px;
                padding-top: 12px;
                background: #fbfcfe;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 4px;
            }
            QFrame#primaryCard,
            QFrame#statusCard,
            QFrame#infoCard {
                border: 1px solid #d9e2ec;
                border-radius: 10px;
                background: #ffffff;
            }
            QLabel#cardTitle {
                color: #607086;
                font-size: 12px;
            }
            QLabel#currentIpValue {
                font-size: 30px;
                font-weight: 700;
                letter-spacing: 0.5px;
                color: #0f172a;
            }
            QLabel#cardValue {
                font-size: 16px;
                font-weight: 600;
                color: #172554;
            }
            QLabel#statusBadge {
                border-radius: 8px;
                padding: 8px 12px;
                font-size: 14px;
                font-weight: 700;
                border: 1px solid #cfd8e3;
                background: #eef2f7;
                color: #475569;
            }
            QLabel#statusBadge[state="idle"] {
                background: #eef2f7;
                color: #64748b;
                border: 1px solid #d5dde8;
            }
            QLabel#statusBadge[state="changed"] {
                background: #e9f9ee;
                color: #167a3f;
                border: 1px solid #c6ebd3;
            }
            QLabel#statusBadge[state="error"] {
                background: #fdecef;
                color: #c62828;
                border: 1px solid #f7c9d1;
            }
            QLabel#statusBadge[state="checking"] {
                background: #eaf3ff;
                color: #1e5aa7;
                border: 1px solid #c9ddfb;
            }
            QLabel#primaryIpValue {
                font-size: 20px;
                font-weight: 600;
                color: #334155;
            }
            QPlainTextEdit {
                border: 1px solid #d3dae6;
                border-radius: 8px;
                background: #ffffff;
                padding: 4px;
            }
            QPushButton#primaryButton {
                border: 1px solid #1f6fd9;
                border-radius: 8px;
                background: #287be6;
                color: #ffffff;
                padding: 4px 16px;
                font-weight: 600;
            }
            QPushButton#primaryButton:hover {
                background: #1f6fd9;
            }
            QPushButton#primaryButton:pressed {
                background: #155ab8;
            }
            QPushButton#secondaryButton {
                border: 1px solid #cfd7e5;
                border-radius: 8px;
                background: #ffffff;
                color: #1f2937;
                padding: 4px 14px;
                font-weight: 500;
            }
            QPushButton#secondaryButton:hover {
                background: #f4f7fc;
            }
            QPushButton#secondaryButton:pressed {
                background: #e9eef7;
            }
            QSpinBox {
                border: 1px solid #cfd7e5;
                border-radius: 8px;
                background: #ffffff;
                padding: 4px 8px;
                color: #1f2937;
            }
            QSpinBox::up-button,
            QSpinBox::down-button {
                width: 16px;
            }
            QStatusBar {
                border-top: 1px solid #d7dfeb;
                background: #f8fafc;
                color: #4b5563;
            }
            """
        )

    def _setup_auto_timer(self) -> None:
        """根据配置启动或停止自动定时检测。"""
        self.auto_detect_timer.stop()
        interval_min = int(self.config.get("auto_detect_interval_minutes", 0))
        if interval_min > 0:
            self.auto_detect_timer.start(interval_min * 60 * 1000)

    def _set_status_text(self, text: str) -> None:
        """更新状态文本并按状态应用颜色。"""
        self.status_value.setText(text)

        if text == "已变化并写入Excel":
            state = "changed"
        elif text == "检测失败":
            state = "error"
        elif text == "检测中...":
            state = "checking"
        else:
            state = "idle"

        self.status_value.setProperty("state", state)
        self.status_value.style().unpolish(self.status_value)
        self.status_value.style().polish(self.status_value)

    def append_log(self, text: str) -> None:
        """向日志区域追加文本。"""
        self.log_output.appendPlainText(text)

    def _refresh_statusbar(self) -> None:
        """刷新底部状态栏文本。"""
        excel_path = self.config_service.get_excel_path(self.config)
        self.status_excel_label.setText(f"Excel: {self._short_excel_path(excel_path)}")
        interval_min = int(self.config.get("auto_detect_interval_minutes", 0))
        interval_text = f"每{interval_min}分钟" if interval_min > 0 else "手动"
        self.status_time_label.setText(
            f"自动检测: {interval_text}  |  上次检测: {self.last_check_time_text}"
        )

    def _init_tray_icon(self) -> None:
        """初始化系统托盘与右键菜单。"""
        if not QSystemTrayIcon.isSystemTrayAvailable():
            self.logger.warning("系统不支持托盘，点击关闭将直接退出")
            return

        tray_icon = self.windowIcon()
        if tray_icon.isNull():
            fallback_icon = self.project_root.parent / "ip.ico"
            if fallback_icon.exists():
                tray_icon = QIcon(str(fallback_icon))

        self.tray_icon = QSystemTrayIcon(tray_icon, self)
        self.tray_icon.setToolTip("公网 IP 监测工具")

        menu = QMenu(self)
        show_action = menu.addAction("显示主界面")
        hide_action = menu.addAction("最小化到托盘")
        menu.addSeparator()
        quit_action = menu.addAction("退出程序")

        show_action.triggered.connect(self._show_from_tray)
        hide_action.triggered.connect(self._hide_to_tray)
        quit_action.triggered.connect(self._quit_from_tray)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self._on_tray_activated)
        self.tray_icon.show()

    def _show_from_tray(self) -> None:
        """从托盘恢复窗口。"""
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def _hide_to_tray(self) -> None:
        """将窗口最小化到托盘。"""
        self.hide()

    def _quit_from_tray(self) -> None:
        """通过托盘菜单真正退出程序。"""
        self._allow_close = True
        if self.tray_icon:
            self.tray_icon.hide()
        self.close()

    def _on_tray_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        """双击托盘图标恢复窗口。"""
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self._show_from_tray()

    def closeEvent(self, event) -> None:
        """拦截窗口关闭动作，默认隐藏到托盘。"""
        if self._allow_close or not self.tray_icon or not self.tray_icon.isVisible():
            event.accept()
            return

        event.ignore()
        self.hide()

        if not self._tray_tip_shown:
            self.tray_icon.showMessage(
                "程序已最小化到托盘",
                "仍在后台监测。右键托盘图标选择“退出程序”可完全关闭。",
                QSystemTrayIcon.MessageIcon.Information,
                2500,
            )
            self._tray_tip_shown = True

    def _short_excel_path(self, excel_path: Path) -> str:
        """将 Excel 路径缩短为更适合状态栏展示的文本。"""
        name = excel_path.name
        parent_name = excel_path.parent.name
        short_text = f".../{parent_name}/{name}" if parent_name else name
        if len(short_text) <= 42:
            return short_text
        return name

    def _normalize_source_list(self, source_list: Any) -> list[dict[str, str]]:
        """读取并规范化双源 IP 配置列表。"""

        valid_sources: list[dict[str, str]] = []
        if isinstance(source_list, list):
            for item in source_list:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("name", "")).strip()
                url = str(item.get("url", "")).strip()
                if not url:
                    continue
                valid_sources.append({"name": name or "来源", "url": url})

        return valid_sources

    def _get_ip_sources(self, config_key: str = "ip_check_urls") -> list[dict[str, str]]:
        """读取并规范化指定键对应的双源 IP 配置。"""
        source_list = self.config.get(config_key, [])
        valid_sources = self._normalize_source_list(source_list)

        # 向后兼容：没有双源配置时使用旧字段作为主源。
        if not valid_sources and config_key == "ip_check_urls":
            legacy_url = str(self.config.get("ip_check_url", "")).strip()
            if legacy_url:
                valid_sources.append({"name": "主源(ipify)", "url": legacy_url})

        # 保底网页对齐源。
        if len(valid_sources) == 1:
            valid_sources.append({"name": "网页对齐源(ifconfig)", "url": "https://ifconfig.me/ip"})

        if len(valid_sources) < 2:
            raise ValueError("配置缺少双源 IP 查询地址")

        return valid_sources[:2]

    def set_excel_path(self) -> None:
        """设置并保存 Excel 路径。"""
        current = str(self.config_service.get_excel_path(self.config))
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择 Excel 文件",
            current,
            "Excel Files (*.xlsx)",
        )

        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path = f"{file_path}.xlsx"

        self.config["excel_path"] = file_path
        try:
            self.config_service.save_config(self.config)
            self.logger.info(f"Excel 路径已更新: {file_path}")
            self._refresh_statusbar()
        except OSError as exc:
            self.logger.error(f"保存配置失败: {exc}")
            QMessageBox.critical(self, "错误", f"保存配置失败:\n{exc}")

    def set_auto_detect_interval(self, interval_min: int) -> None:
        """设置自动检测间隔（分钟），0 表示关闭自动检测。"""
        interval_min = max(0, int(interval_min))
        old_interval = int(self.config.get("auto_detect_interval_minutes", 0))
        if interval_min == old_interval:
            return

        self.config["auto_detect_interval_minutes"] = interval_min
        self._setup_auto_timer()
        self._refresh_statusbar()

        try:
            self.config_service.save_config(self.config)
            if interval_min > 0:
                self.logger.info(f"自动检测间隔已更新为每 {interval_min} 分钟")
            else:
                self.logger.info("自动检测已关闭（仅手动检测）")
        except OSError as exc:
            self.logger.error(f"保存自动检测设置失败: {exc}")
            QMessageBox.warning(self, "提示", f"保存自动检测设置失败:\n{exc}")

    def open_excel(self) -> None:
        """打开 Excel 文件。"""
        excel_path = self.config_service.get_excel_path(self.config)

        try:
            self.excel_service.ensure_workbook(excel_path)
        except ExcelServiceError as exc:
            self.logger.error(str(exc))
            QMessageBox.critical(self, "错误", str(exc))
            return

        try:
            # Windows 优先使用系统默认程序打开 Excel。
            os.startfile(str(excel_path))
            self.logger.info(f"已打开 Excel: {excel_path}")
        except OSError as exc:
            self.logger.error(f"打开 Excel 失败: {exc}")
            QMessageBox.warning(self, "提示", f"无法打开 Excel 文件:\n{exc}")

    def open_site(self) -> None:
        """手动打开网站。"""
        self._open_auto_site(reason="手动")

    def _open_auto_site(self, reason: str) -> None:
        """打开配置中的网站地址。"""
        url = str(self.config.get("auto_open_url", "https://whatismyipaddress.com")).strip()
        if not url:
            self.logger.warning("网站地址为空，跳过打开")
            return

        try:
            webbrowser.open(url, new=2)
            self.logger.info(f"已打开网站({reason}): {url}")
        except OSError as exc:
            self.logger.error(f"打开网站失败: {exc}")

    def detect_now(self, *, open_site: bool = False) -> None:
        """执行一次 IP 检测流程。"""
        if self._is_detecting:
            self.logger.warning("检测任务仍在执行中，已跳过本次请求")
            return

        self.detect_btn.setEnabled(False)
        self._set_status_text("检测中...")
        self.logger.info("开始执行 IP 检测")
        try:
            source_configs = self._get_ip_sources()
            fallback_source_configs: list[dict[str, str]] | None = None
            if bool(self.config.get("enable_cn_fallback", True)):
                fallback_source_configs = self._normalize_source_list(
                    self.config.get("ip_check_urls_cn", [])
                )
                if len(fallback_source_configs) < 2:
                    fallback_source_configs = None
            ip_info_url_template = str(self.config.get("ip_info_url_template", "")).strip()
            excel_path = self.config_service.get_excel_path(self.config)

            if not ip_info_url_template:
                raise ValueError("配置缺少 IP 详情查询地址")

        except (IPServiceError, ExcelServiceError, ValueError) as exc:
            self._set_status_text("检测失败")
            self.primary_ip_value.setText("-")
            self.logger.error(str(exc))
            QMessageBox.warning(self, "检测失败", str(exc))
            self.last_check_time_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._refresh_statusbar()
            self.detect_btn.setEnabled(True)
            return

        self._pending_open_site = open_site
        self._is_detecting = True

        self._detect_thread = QThread(self)
        self._detect_worker = DetectWorker(
            excel_path=excel_path,
            source_configs=source_configs,
            fallback_source_configs=fallback_source_configs,
            ip_info_url_template=ip_info_url_template,
            ip_service=self.ip_service,
            excel_service=self.excel_service,
        )
        self._detect_worker.moveToThread(self._detect_thread)

        self._detect_thread.started.connect(self._detect_worker.run)
        self._detect_worker.succeeded.connect(self._on_detect_success)
        self._detect_worker.failed.connect(self._on_detect_failed)
        self._detect_worker.finished.connect(self._on_detect_finished)
        self._detect_worker.finished.connect(self._detect_thread.quit)
        self._detect_worker.finished.connect(self._detect_worker.deleteLater)
        self._detect_thread.finished.connect(self._detect_thread.deleteLater)

        self._detect_thread.start()

    def _on_detect_success(self, payload: dict[str, Any]) -> None:
        """在主线程处理检测成功结果并更新 UI。"""
        self.current_ip_value.setText(str(payload.get("web_ip", "-") or "-"))
        self.primary_ip_value.setText(str(payload.get("primary_ip", "-") or "-"))
        self.last_ip_value.setText(str(payload.get("display_last_web", "-") or "-"))
        self.isp_value.setText(str(payload.get("isp", "-") or "-"))
        region = str(payload.get("region", "-") or "-")
        city = str(payload.get("city", "-") or "-")
        self.location_value.setText(f"{region} / {city}")
        self._set_status_text(str(payload.get("status_text", "未变化")))

        for line in payload.get("logs", []):
            self.logger.info(str(line))

    def _on_detect_failed(self, message: str, is_unexpected: bool) -> None:
        """在主线程处理检测失败结果并提示用户。"""
        self._set_status_text("检测失败")
        self.primary_ip_value.setText("-")
        self.logger.error(message)
        if is_unexpected:
            QMessageBox.critical(self, "错误", f"发生未预期错误:\n{message}")
        else:
            QMessageBox.warning(self, "检测失败", message)

    def _on_detect_finished(self) -> None:
        """收尾逻辑：恢复按钮状态、更新时间与状态栏。"""
        self._is_detecting = False
        self.last_check_time_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._refresh_statusbar()
        self.detect_btn.setEnabled(True)

        should_open_site = self._pending_open_site
        self._pending_open_site = False
        if should_open_site:
            self._open_auto_site(reason="手动")

        self._detect_worker = None
        self._detect_thread = None
