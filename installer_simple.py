#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
雪之下倾城 - 简单向导式安装程序
使用系统默认经典样式，有说明、下一步、默认目录、完成
"""
import os
import sys
import shutil
import time
import subprocess
import tempfile

from PySide6.QtWidgets import (
    QApplication, QWizard, QWizardPage, QLabel, QVBoxLayout,
    QHBoxLayout, QLineEdit, QPushButton, QProgressBar, QTextEdit,
    QMessageBox, QCheckBox, QDialog
)
from PySide6.QtCore import Qt, QTimer, Signal, QThread
from PySide6.QtGui import QIcon, QFont

APP_NAME = "雪之下倾城"
INSTALL_DIR = r"D:\游戏菜单"
VERSION = "7.3.18"  # 内部版本号（数字格式，用于比较）
VERSION_DISPLAY = "V7.3.18"  # 显示格式
# 版本信息网页（获取最新版本号、公告、下载地址）
VERSION_URL = "http://xgadmin520.3vkj.vip/version.html"


def format_version(ver):
    """把版本号格式化为 V7.0 显示格式（去掉末尾的.0）"""
    try:
        parts = str(ver).lstrip('vV').split('.')
        if len(parts) >= 3 and parts[2] == '0':
            return "V" + ".".join(parts[:2])
        return "V" + str(ver).lstrip('vV')
    except Exception:
        return f"V{ver}"


def fetch_latest_version():
    """从服务器网页获取最新版本号/下载地址/访问密码（后台线程调用）"""
    try:
        import re, urllib.request, ssl
        # 用urllib直接获取网页内容，避免powershell减少杀毒误报
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(VERSION_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            html_text = resp.read().decode("utf-8", errors="ignore")
        # 兼容多种写法：class="version"、class='version'、class="version-badge" 内、纯文本 V7.3.0 / 7.3.0
        patterns = [
            r'class=["\']version["\']\s*>\s*([^<]+)',
            r'class=["\']version-badge["\'][^>]*>\s*<[^>]*>\s*([^<]+)',
            r'最新版本[^0-9]*([Vv]?\d+\.\d+[\.\d]*)',
            r'([Vv]\d+\.\d+[\.\d]*)',
            r'(\d+\.\d+\.\d+)',
        ]
        version_display = None
        for pat in patterns:
            m = re.search(pat, html_text)
            if m:
                ver = m.group(1).strip()
                if re.match(r'^[Vv]?\d', ver):
                    version_display = format_version(ver)
                    break
        # 下载地址
        url = ""
        m = re.search(r'class=["\']download_url["\']\s*>\s*([^<]+)', html_text)
        if m:
            url = m.group(1).strip()
        # 访问密码
        pwd = ""
        m = re.search(r'class=["\']download_password["\']\s*>\s*([^<]+)', html_text)
        if m:
            pwd = m.group(1).strip()
        return version_display, url, pwd
    except Exception as e:
        print(f"[fetch_latest_version error] {e}")
    return None, "", ""


def fetch_web_notice():
    """从服务器网页获取滚动公告文字（class=marquee / notice）"""
    import re
    try:
        import urllib.request, ssl
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(VERSION_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            html_text = resp.read().decode("utf-8", errors="ignore")
        text = ""
        for pat in (r'class=["\']marquee["\'][^>]*>([^<]+)',
                    r'class=["\']notice["\'][^>]*>(.*?)</div>',
                    r'class=["\']notice["\'][^>]*>([^<]+)'):
            m = re.search(pat, html_text, re.S)
            if m:
                text = re.sub(r'<[^>]+>', '', m.group(1))
                text = text.replace('&nbsp;', ' ').strip()
                if text:
                    break
        return text
    except Exception as e:
        print(f"[fetch_web_notice error] {e}")
        return ""


def launch_app_shell(exe_path, work_dir):
    """通过 ShellExecute 启动程序（不继承父进程权限/控制台，规避错误740需要提升）"""
    try:
        import ctypes
        rc = ctypes.windll.shell32.ShellExecuteW(
            None, "open", exe_path, None, work_dir, 1)  # SW_SHOWNORMAL=1
        if rc is not None and rc > 32:
            return True
    except Exception:
        pass
    # 回退：explorer 方式启动（完全脱离父进程）
    try:
        subprocess.Popen(
            ["cmd", "/c", "start", "", exe_path],
            cwd=work_dir, creationflags=0x08000000)
        return True
    except Exception:
        return False


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def create_shortcut(shortcut_path, target_path, work_dir, icon_path=None, arguments=""):
    """用win32com创建快捷方式，避免powershell调用减少杀毒误报"""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        sc = shell.CreateShortcut(shortcut_path)
        sc.TargetPath = target_path
        sc.WorkingDirectory = work_dir
        sc.Arguments = arguments
        if icon_path and os.path.exists(icon_path):
            sc.IconLocation = icon_path
        sc.Save()
        return os.path.exists(shortcut_path)
    except Exception:
        return False


def _resolve_lnk_dir(lnk_path):
    """解析 .lnk 快捷方式，返回目标程序所在目录（失败返回 None）"""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        sc = shell.CreateShortcut(lnk_path)
        t = (sc.TargetPath or "").strip().strip('"')
        if t:
            return os.path.dirname(t)
    except Exception:
        pass
    return None


def _reg_run_dir():
    """从注册表开机启动项读取游戏菜单安装目录"""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_READ)
        try:
            val, _ = winreg.QueryValueEx(key, APP_NAME)
        finally:
            winreg.CloseKey(key)
        val = (val or "").strip().strip('"')
        if val:
            return os.path.dirname(val)
    except Exception:
        pass
    return None


def _running_exe_dirs():
    """获取正在运行的游戏菜单进程所在目录（用tasklist避免powershell）"""
    dirs = []
    try:
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq 雪之下倾城.exe", "/FO", "CSV", "/NH"],
                           capture_output=True, timeout=8, creationflags=0x08000000)
        out = (r.stdout or b"").decode("gbk", errors="ignore") if isinstance(r.stdout, bytes) else (r.stdout or "")
        if "雪之下倾城.exe" in out:
            # tasklist不直接给路径，用wmic
            r2 = subprocess.run(["wmic", "process", "where", "name='雪之下倾城.exe'", "get", "ExecutablePath", "/FORMAT:LIST"],
                                capture_output=True, timeout=8, creationflags=0x08000000)
            out2 = (r2.stdout or b"").decode("gbk", errors="ignore") if isinstance(r2.stdout, bytes) else (r2.stdout or "")
            for line in out2.splitlines():
                line = line.strip()
                if line.startswith("ExecutablePath="):
                    p = line.split("=", 1)[1].strip().strip('"')
                    if p:
                        dirs.append(os.path.dirname(p))
    except Exception:
        pass
    return dirs


def find_latest_game_config(preferred_dir):
    """
    在所有可能的安装位置中查找 game_menu_config.json，
    返回 (最新配置路径, 候选目录列表)。优先用户选择的目录，
    但以文件修改时间最新为准，避免备份到旧目录的旧数据。
    """
    candidates = []
    seen = set()

    def _add(d):
        if not d:
            return
        d = os.path.normpath(d)
        if d.lower() not in seen:
            seen.add(d.lower())
            candidates.append(d)

    _add(preferred_dir)
    _add(_reg_run_dir())
    for d in _running_exe_dirs():
        _add(d)
    home = os.path.expanduser("~")
    desktop = os.path.join(home, "Desktop")
    if not os.path.exists(desktop):
        desktop = r"C:\Users\Public\Desktop"
    lnk_dirs = [
        desktop,
        os.path.join(home, r"AppData\Roaming\Microsoft\Windows\Start Menu\Programs"),
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
        r"C:\Users\Public\Desktop",
    ]
    for ld in lnk_dirs:
        try:
            if os.path.isdir(ld):
                for root, _, files in os.walk(ld):
                    for fn in files:
                        if fn.lower().endswith(".lnk") and (APP_NAME in fn or "游戏菜单" in fn):
                            _add(_resolve_lnk_dir(os.path.join(root, fn)))
        except Exception:
            pass
    for d in [r"D:\游戏菜单", r"C:\游戏菜单",
              os.path.join(home, "游戏菜单"),
              r"C:\Program Files\游戏菜单",
              r"C:\Program Files (x86)\游戏菜单"]:
        _add(d)

    configs = []
    for d in candidates:
        cf = os.path.join(d, "game_menu_config.json")
        if os.path.isfile(cf):
            try:
                configs.append((os.path.getmtime(cf), cf))
            except Exception:
                pass
    if not configs:
        return None, candidates
    configs.sort(key=lambda x: x[0], reverse=True)
    return configs[0][1], candidates


def _load_banner(w=580, h=300):
    """加载大司命横幅图（随安装包内置），宽度适配窗口、居中裁剪"""
    try:
        from PySide6.QtGui import QPixmap
        path = resource_path("installer_banner.png")
        if not os.path.exists(path):
            return None
        pm = QPixmap(path)
        if pm.isNull():
            return None
        pm2 = pm.scaled(w, h, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
        x = (pm2.width() - w) // 2
        y = (pm2.height() - h) // 2
        return pm2.copy(x, y, w, h)
    except Exception:
        return None


def _load_banner_keep(w=500, h=380, fname="installer_banner.png"):
    """加载大司命完整立绘（保持比例不裁剪），宽度适配窗口"""
    try:
        from PySide6.QtGui import QPixmap
        path = resource_path(fname)
        if not os.path.exists(path):
            return None
        pm = QPixmap(path)
        if pm.isNull():
            return None
        return pm.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    except Exception:
        return None


def _load_banner_fill(w=620, h=280, fname="installer_banner.png"):
    """加载banner图片，宽度拉满，高度按比例自适应"""
    try:
        from PySide6.QtGui import QPixmap
        path = resource_path(fname)
        if not os.path.exists(path):
            return None
        pm = QPixmap(path)
        if pm.isNull():
            return None
        # 宽度拉满w，高度按原图比例计算
        ratio = pm.height() / pm.width()
        new_h = int(w * ratio)
        return pm.scaled(w, new_h, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
    except Exception:
        return None


class WelcomePage(QWizardPage):
    """BILIBILI LIVE 风格单页：大图 + 立即安装 + 目录选择 + 协议勾选"""

    def __init__(self):
        super().__init__()
        root = QVBoxLayout()
        root.setContentsMargins(16, 8, 16, 8)
        root.setSpacing(12)

        # 顶部大图（大司命横版 banner）- 宽度拉满
        self.banner_lbl = QLabel()
        _banner = _load_banner_fill(620, 280)
        if _banner is not None:
            self.banner_lbl.setPixmap(_banner)
        self.banner_lbl.setAlignment(Qt.AlignCenter)
        self.banner_lbl.setStyleSheet("border-radius:12px;")
        root.addWidget(self.banner_lbl)

        # 立即安装按钮（白色胶囊，居中）
        self.install_btn = QPushButton("立即安装")
        self.install_btn.setFixedSize(180, 46)
        self.install_btn.setCursor(Qt.PointingHandCursor)
        self.install_btn.setStyleSheet("""
            QPushButton { background:#ffffff; color:#d63384; border:none;
                border-radius:23px; font-size:18px; font-weight:bold; }
            QPushButton:hover { background:#ffeaf4; }
        """)
        self.install_btn.clicked.connect(self._do_install)
        btn_row = QHBoxLayout()
        btn_row.addStretch(); btn_row.addWidget(self.install_btn); btn_row.addStretch()
        root.addLayout(btn_row)

        # 异常提示
        tip = QLabel("若安装过程出现异常，请重新下载安装包并运行")
        tip.setStyleSheet("color:#9a9a9a; font-size:11px;")
        tip.setAlignment(Qt.AlignCenter)
        root.addWidget(tip)

        # 安装目录行
        dir_row = QHBoxLayout()
        dir_lbl = QLabel("安装目录：")
        dir_lbl.setStyleSheet("color:#ffffff; font-size:13px; font-weight:bold;")
        self.dir_edit = QLineEdit(INSTALL_DIR)
        self.dir_edit.setMinimumHeight(34)
        self.dir_edit.setStyleSheet("""
            QLineEdit { background:rgba(255,255,255,25); color:#ffffff;
                border:1px solid rgba(255,255,255,55); border-radius:6px;
                padding:5px 10px; font-size:13px; }
        """)
        custom_btn = QPushButton("浏览")
        custom_btn.setFixedHeight(34)
        custom_btn.setFixedWidth(80)
        custom_btn.setStyleSheet("""
            QPushButton {
                background:rgba(255,51,170,180); color:#fff; border:none;
                border-radius:6px; font-size:13px; font-weight:bold;
            }
            QPushButton:hover { background:rgba(255,51,170,220); }
        """)
        custom_btn.setCursor(Qt.PointingHandCursor)
        custom_btn.clicked.connect(self._browse)
        dir_row.addWidget(dir_lbl)
        dir_row.addWidget(self.dir_edit, 4)
        dir_row.addWidget(custom_btn)
        root.addLayout(dir_row)

        # 协议勾选（默认勾选）+ 可点击查看协议/公告
        agree_row = QHBoxLayout()
        self.agree_chk = QCheckBox("已阅读并同意")
        self.agree_chk.setChecked(True)
        self.agree_chk.setStyleSheet("color:#e8b8d8; font-size:12px;")
        self.agree_link = QPushButton("《游戏菜单 许可及服务协议》")
        self.agree_link.setCursor(Qt.PointingHandCursor)
        self.agree_link.setFocusPolicy(Qt.NoFocus)
        self.agree_link.setAutoDefault(False)
        self.agree_link.setDefault(False)
        _lf = QFont()
        _lf.setUnderline(True)
        self.agree_link.setFont(_lf)
        self.agree_link.setStyleSheet("""
            QPushButton { background:transparent; color:#ff77c9; border:none;
                font-size:12px; padding:0; text-align:left; }
            QPushButton:hover { color:#ffaadd; }
        """)
        self.agree_link.clicked.connect(self._show_agreement)
        agree_row.addWidget(self.agree_chk)
        agree_row.addWidget(self.agree_link)
        agree_row.addStretch()
        root.addLayout(agree_row)

        root.addStretch()
        self.setLayout(root)

    def initializePage(self):
        # 隐藏标准导航按钮，界面更干净（靠"立即安装"和右上角X）
        QTimer.singleShot(0, self._hide_buttons)

    def _hide_buttons(self):
        try:
            wiz = self.wizard()
            if wiz:
                for _b in (QWizard.NextButton, QWizard.BackButton, QWizard.CancelButton):
                    _btn = wiz.button(_b)
                    if _btn:
                        _btn.hide()
        except Exception:
            pass

    def get_install_dir(self):
        return self.dir_edit.text().strip() or INSTALL_DIR

    def _browse(self):
        from PySide6.QtWidgets import QFileDialog
        d = QFileDialog.getExistingDirectory(
            self, "选择安装目录", self.dir_edit.text().strip() or "D:\\")
        if d:
            self.dir_edit.setText(d)

    def _do_install(self):
        if not self.agree_chk.isChecked():
            QMessageBox.warning(
                self, "提示", "请先勾选同意《游戏菜单 许可及服务协议》")
            return
        if not self.dir_edit.text().strip():
            QMessageBox.warning(self, "提示", "请选择安装目录！")
            return
        self.wizard().next()

    def _show_agreement(self, *args):
        """点击协议名：弹出协议/公告窗口，隐藏安装界面"""
        dlg = AgreementDialog(self.wizard())
        dlg.exec()


class NoticeWorker(QThread):
    """后台获取网页公告"""
    got = Signal(str)

    def run(self):
        try:
            txt = fetch_web_notice()
            self.got.emit(txt if txt else "暂时无法获取网页公告，请检查网络后重试。")
        except Exception:
            self.got.emit("暂时无法获取网页公告。")


class AgreementDialog(QDialog):
    """许可协议 + 网页公告。打开时隐藏安装向导，点返回恢复。"""

    def __init__(self, wizard):
        super().__init__()
        self._wizard = wizard
        self.setWindowTitle("游戏菜单许可及服务协议")
        self.setFixedSize(600, 580)
        self.setStyleSheet("""
            QDialog { background:#150828; }
            QLabel { color:#ffffff; background:transparent; }
            QTextEdit { background:rgba(10,5,25,240); color:#e8e8ff;
                border:1px solid #cc66ff; border-radius:8px;
                font-size:13px; font-family:'Microsoft YaHei'; }
            QPushButton { background:#e83e8c; color:#fff; border:none;
                border-radius:18px; font-size:14px; font-weight:bold;
                min-height:36px; }
            QPushButton:hover { background:#ff66b2; }
        """)
        lay = QVBoxLayout()
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(12)
        title = QLabel("游戏菜单许可及服务协议")
        title.setStyleSheet("color:#ff77c9; font-size:22px; font-weight:bold;")
        title.setAlignment(Qt.AlignCenter)
        lay.addWidget(title)

        agree = QTextEdit()
        agree.setReadOnly(True)
        agree.setText(
            "《游戏菜单 许可及服务协议》\n\n"
            "1. 本软件为免费的家庭游戏快捷方式管理工具，仅用于整理本机软件与游戏的快捷启动。\n"
            "2. 软件不修改系统关键设置，不收集任何个人隐私信息。\n"
            "3. 软件支持主题壁纸、漂浮特效、鼠标指针、灵动岛提示、天气预报等个性化功能。\n"
            "4. 如杀毒软件误报，请将软件加入信任白名单。\n"
            "5. 卸载时会先备份您的软件列表数据，可随时重新安装恢复。\n\n"
            "—————————— 家庭版使用说明 ——————————\n"
            "① 添加软件：把桌面或文件夹里的 exe / 快捷方式直接拖进游戏菜单，即可自动添加图标。\n"
            "② 管理图标：在图标上点鼠标右键，可重命名、复制、剪切、粘贴、打开文件所在位置、删除。\n"
            "③ 移动位置：按住图标拖动到网格任意位置，松开后自动保存，重启后位置不变。\n"
            "④ 图标大小：在设置里可切换大/中/小三档，图标位置保持不变。\n"
            "⑤ 主题壁纸：在设置里切换动漫、游戏等多种壁纸，设置与主界面同步生效。\n"
            "⑥ 漂浮特效与鼠标指针：可在设置里选择喜欢的漂浮特效和动漫鼠标指针。\n"
            "⑦ 灵动岛提示：切换主题、重命名、添加软件等操作会在屏幕顶部弹出灵动岛提示。\n"
            "⑧ 天气预报：点底部天气胶囊查看未来7日天气，满7天后自动获取新预报。\n"
            "⑨ 最小化与关闭：点最小化以星光粒子特效收进托盘，双击托盘图标可恢复窗口。\n"
            "⑩ 备份数据：软件列表自动备份，也可把备份文件拖进窗口一键恢复。"
        )
        lay.addWidget(agree)

        back = QPushButton("返回安装界面")
        back.setCursor(Qt.PointingHandCursor)
        back.clicked.connect(self._back)
        lay.addWidget(back)
        self.setLayout(lay)

    def showEvent(self, e):
        try:
            if self._wizard:
                self._wizard.hide()
        except Exception:
            pass
        super().showEvent(e)

    def _set_notice(self, text):
        self.notice_lbl.setText(str(text))

    def _back(self):
        try:
            if self._wizard:
                self._wizard.show()
                self._wizard.raise_()
        except Exception:
            pass
        self.accept()


class DirPage(QWizardPage):
    """选择安装目录页面（大司命宣传图在下方，宽度适配窗口）"""
    def __init__(self):
        super().__init__()
        self.setTitle("选择安装目录")
        self.setSubTitle("请选择游戏菜单的安装位置（默认D:\游戏菜单）")

        layout = QVBoxLayout()
        hint = QLabel("安装目录：")
        hint.setFont(QFont("Microsoft YaHei", 11))
        hint.setStyleSheet("color: #ffffff; font-weight: bold;")
        layout.addWidget(hint)

        dir_layout = QHBoxLayout()
        self.dir_edit = QLineEdit(INSTALL_DIR)
        self.dir_edit.setFont(QFont("Microsoft YaHei", 12))
        self.dir_edit.setMinimumHeight(40)
        self.dir_edit.setMinimumWidth(450)
        dir_layout.addWidget(self.dir_edit, 4)

        browse_btn = QPushButton("浏览...")
        browse_btn.setMinimumHeight(40)
        browse_btn.setFont(QFont("Microsoft YaHei", 12))
        browse_btn.setMinimumWidth(100)
        browse_btn.clicked.connect(self._browse)
        dir_layout.addWidget(browse_btn, 1)
        layout.addLayout(dir_layout)

        space_hint = QLabel("\n⚠ 提示：安装需要约100MB磁盘空间\n\n安装时自动备份软件列表数据到桌面，\n安装后自动恢复您的软件和图标位置，\n保留鼠标指针和上传的图标图片")
        space_hint.setStyleSheet("color: #ffd94d; font-size: 13px; font-weight: bold;")
        layout.addWidget(space_hint)

        # 目录页专用大司命立绘：放到下方（高度受控，保证底部按钮可见）
        _banner = _load_banner_keep(450, 380, "installer_dir_banner.png")
        if _banner is not None:
            self.banner_lbl = QLabel()
            self.banner_lbl.setPixmap(_banner)
            self.banner_lbl.setAlignment(Qt.AlignCenter)
            self.banner_lbl.setStyleSheet(
                "border:2px solid rgba(255,51,170,120); border-radius:10px; padding:4px; margin-top:10px;")
            layout.addWidget(self.banner_lbl)

        layout.addStretch()
        self.setLayout(layout)

    def isComplete(self):
        return True

    def validatePage(self):
        if not self.dir_edit.text().strip():
            QMessageBox.warning(self, "提示", "请输入安装目录！")
            return False
        return True

    def get_install_dir(self):
        return self.dir_edit.text().strip() or INSTALL_DIR

    def _browse(self):
        from PySide6.QtWidgets import QFileDialog
        d = QFileDialog.getExistingDirectory(self, "选择安装目录", self.dir_edit.text().strip() or "D:\\")
        if d:
            self.dir_edit.setText(d)
        self.completeChanged.emit()


class InstallWorker(QThread):
    """后台安装线程：杀进程/删文件/复制等耗时操作都在子线程，避免界面卡死(未响应)"""
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal()

    def __init__(self, install_dir):
        super().__init__()
        self.install_dir = install_dir

    def run(self):
        try:
            self._log("=" * 40)
            self._log(f"  {APP_NAME} - 安装程序")
            self._log("=" * 40)
            self._log("安装目录: " + self.install_dir)
            self._log("")
            steps = [
                ("[1/7] 正在强制关闭运行中的游戏菜单...", 10, self._step_kill),
                ("[2/7] 正在保存游戏数据（备份到桌面）...", 25, self._step_save_data),
                ("[3/7] 正在创建安装目录...", 40, self._step_mkdir),
                ("[4/7] 正在复制程序文件...", 60, self._step_copy_exe),
                ("[5/7] 正在复制资源文件...", 78, self._step_copy_res),
                ("[6/7] 正在创建桌面快捷方式...", 92, self._step_shortcuts),
                ("[7/7] 安装完成！", 100, self._step_done),
            ]
            for msg, prog, func in steps:
                self._log(msg)
                self.progress_signal.emit(prog)
                try:
                    func()
                except Exception as e:
                    self._log("  警告: " + str(e))
            self._log("")
            self._log("安装完成！点击「完成」即可启动游戏菜单...")
            self.progress_signal.emit(100)
        except Exception as e:
            self._log("安装失败: " + str(e))
        finally:
            self.finished_signal.emit()

    def _log(self, msg):
        self.log_signal.emit(str(msg))

    def _step_gentle_close(self):
        """强杀前先给运行中的游戏菜单发 WM_CLOSE，让其 closeEvent 保存最新配置（含图标位置）"""
        try:
            import ctypes
            from ctypes import wintypes
            u = ctypes.windll.user32
            WM_CLOSE = 0x0010
            targets = []

            import os as _os
            _my_pid = _os.getpid()
            _pid_ret = wintypes.DWORD()
            def _enum(hwnd, _l):
                if u.IsWindowVisible(hwnd):
                    cn = ctypes.create_unicode_buffer(256)
                    u.GetClassNameW(hwnd, cn, 256)
                    tb = ctypes.create_unicode_buffer(256)
                    u.GetWindowTextW(hwnd, tb, 256)
                    # 排除安装向导自己：标题含"安装"，或窗口属于当前进程
                    if "安装" in tb.value:
                        return True
                    u.GetWindowThreadProcessId(hwnd, ctypes.byref(_pid_ret))
                    if _pid_ret.value == _my_pid:
                        return True
                    if "Qt" in cn.value and (APP_NAME in tb.value or "游戏菜单" in tb.value):
                        targets.append(hwnd)
                return True

            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
            u.EnumWindows(WNDENUMPROC(_enum), 0)
            if targets:
                self._log("正在请求运行中的游戏菜单保存并退出...")
                for hwnd in targets:
                    try:
                        u.PostMessageW(hwnd, WM_CLOSE, 0, 0)
                    except Exception:
                        pass
                # closeEvent 有 1.2 秒灵动岛延迟 + 最多 2 秒保存，共等待 4 秒
                import time as _t
                _t.sleep(4.0)
        except Exception:
            pass

    def _step_kill(self):
        """强制结束所有占用游戏菜单文件的后台进程"""
        import time as _time
        # 先温和关闭，让游戏菜单把最新配置（软件/图标位置）写盘
        self._step_gentle_close()
        # 需要结束的进程名列表
        kill_names = [
            f"{APP_NAME}.exe",
            "雪之下倾城.exe",
            "游戏菜单.exe",
            "game_menu.exe",
            "雪之下倾城_游戏菜单安装包.exe",
        ]
        for name in kill_names:
            for attempt in range(3):
                try:
                    result = subprocess.run(
                        ["taskkill", "/F", "/IM", name, "/T"],
                        capture_output=True, timeout=10,
                        creationflags=0x08000000
                    )
                    # 128 = 找不到进程（正常），0 = 成功结束
                    if result.returncode in (0, 128):
                        break
                except Exception:
                    pass
                _time.sleep(0.5)
        # 额外：用 PowerShell 按进程名强杀（Stop-Process 失败再用 WMI Terminate 兜底）
        try:
            # 用taskkill代替powershell杀进程，减少杀毒误报
            subprocess.run(["taskkill", "/F", "/IM", f"{APP_NAME}.exe"],
                           capture_output=True, timeout=10, creationflags=0x08000000)
        except Exception:
            pass
        _time.sleep(1.5)
        # 验证主程序是否已退出
        for attempt in range(5):
            try:
                check = subprocess.run(
                    ["tasklist", "/FI", f"IMAGENAME eq {APP_NAME}.exe"],
                    capture_output=True, text=True, encoding="gbk", errors="ignore", timeout=5,
                    creationflags=0x08000000
                )
                if APP_NAME not in check.stdout or "INFO" in check.stdout:
                    break
            except Exception:
                break
            _time.sleep(0.5)

    def _step_save_data(self):
        """备份游戏数据：智能查找最新的 game_menu_config.json，备份到桌面（统一文件名+日期时间），安装后恢复"""
        import json as _json
        import datetime as _dt
        import tempfile
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.exists(desktop):
                desktop = r"C:\Users\Public\Desktop"
            self._backup_cfg_tmp = None
            self._desk_backup = os.path.join(desktop, "游戏菜单软件列表_备份导出数据.json")

            # 智能查找：在安装目录、注册表启动项、桌面/开始菜单快捷方式、运行进程、常见路径中
            # 找修改时间最新的 game_menu_config.json，避免备份到旧目录的旧数据
            cfg_file, candidates = find_latest_game_config(self.install_dir)

            if cfg_file and os.path.exists(cfg_file):
                # 读取原始配置（兼容BOM）
                with open(cfg_file, 'r', encoding='utf-8-sig') as f:
                    cfg_data = _json.load(f)
                # 安装完成自动勾选开机启动
                cfg_data["autostart"] = True

                # 统计软件数
                software_list = []
                for cat in ["网络游戏", "单机游戏", "聊天娱乐"]:
                    if cat in cfg_data.get("items", {}):
                        for item in cfg_data["items"][cat]:
                            software_list.append({
                                "category": cat,
                                "name": item.get("name", ""),
                                "path": item.get("path", ""),
                                "icon_path": item.get("icon_path", ""),
                                "x": item.get("x", 10),
                                "y": item.get("y", 10),
                            })

                # 来源配置的修改时间（证明备份的是最新内容）
                src_mtime = _dt.datetime.fromtimestamp(os.path.getmtime(cfg_file))
                now_str = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                src_dir = os.path.dirname(cfg_file)

                # 写到临时备份（安装后恢复用）
                self._backup_cfg_tmp = os.path.join(tempfile.gettempdir(), "_game_menu_cfg_backup.json")
                with open(self._backup_cfg_tmp, 'w', encoding='utf-8') as f:
                    _json.dump(cfg_data, f, ensure_ascii=False, indent=2)

                # 生成标准导出格式到桌面（用户可拖入软件恢复，只留一份，带日期时间）
                export_data = {
                    "export_time": now_str,
                    "app_title": cfg_data.get("app_title", APP_NAME),
                    "total": len(software_list),
                    "software": software_list,
                }
                # 先删除旧备份（确保只留一份）
                if os.path.exists(self._desk_backup):
                    try:
                        os.remove(self._desk_backup)
                    except Exception:
                        pass
                with open(self._desk_backup, 'w', encoding='utf-8') as f:
                    _json.dump(export_data, f, ensure_ascii=False, indent=2)
                # 清理历史多余备份文件
                for _f in os.listdir(desktop):
                    if _f.startswith("游戏菜单软件列表") and _f.endswith(".json") and _f != "游戏菜单软件列表_备份导出数据.json":
                        try:
                            os.remove(os.path.join(desktop, _f))
                            self._log("✓ 清理旧备份: " + _f)
                        except Exception:
                            pass
                self._log(f"✓ 数据来源: {src_dir}")
                self._log(f"✓ 配置最后修改: {src_mtime.strftime('%Y-%m-%d %H:%M:%S')}")
                self._log(f"✓ 已备份 {len(software_list)} 个软件（含图标位置）到桌面")
                self._log(f"✓ 备份时间: {now_str}")
            else:
                self._log("未检测到已有软件列表数据，无需备份")
        except Exception as e:
            self._log("保存数据警告: " + str(e))

    def _step_mkdir(self):
        os.makedirs(self.install_dir, exist_ok=True)
        # 先卸掉旧版：删除目录内的文件（不覆盖，先删后装），
        # 配置文件已备份到桌面，这里清空所有文件；保留 menu_assets（鼠标/背景图）、icons（上传图标）、使用说明.txt
        # 强制结束占用安装目录内 exe 的进程
        try:
            old_exe = os.path.join(self.install_dir, f"{APP_NAME}.exe")
            if os.path.exists(old_exe):
                self._force_kill_file_lock(old_exe)
                import time as _time
                _time.sleep(1)
        except Exception:
            pass
        try:
            for _name in os.listdir(self.install_dir):
                _fp = os.path.join(self.install_dir, _name)
                _low = _name.lower()
                if os.path.isfile(_fp):
                    if _low in ('使用说明.txt',):
                        continue
                    try:
                        os.remove(_fp)
                    except PermissionError:
                        # 文件被占用，强制结束进程后重试
                        self._force_kill_file_lock(_fp)
                        import time as _time
                        _time.sleep(0.5)
                        try:
                            os.remove(_fp)
                        except Exception:
                            pass
                    except Exception:
                        pass
                elif os.path.isdir(_fp):
                    # icons 目录保留（上传图标），menu_assets 清空内容（主题壁纸重新加载），其余子目录全部删除
                    if _low == 'icons':
                        continue
                    if _low == 'menu_assets':
                        # 清空 menu_assets 目录里的所有文件（主题壁纸图片），但保留目录本身
                        try:
                            for _sub in os.listdir(_fp):
                                _subfp = os.path.join(_fp, _sub)
                                if os.path.isfile(_subfp):
                                    try:
                                        os.remove(_subfp)
                                    except Exception:
                                        pass
                                elif os.path.isdir(_subfp):
                                    try:
                                        shutil.rmtree(_subfp, ignore_errors=True)
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                        continue
                    try:
                        shutil.rmtree(_fp, ignore_errors=True)
                    except Exception:
                        pass
        except Exception:
            pass

    def _force_kill_file_lock(self, target_path):
        """强制结束占用指定文件的进程"""
        import time as _time
        try:
            exe_name = os.path.basename(target_path)
            if not exe_name.lower().endswith('.exe'):
                return
            for _ in range(3):
                subprocess.run(
                    ["taskkill", "/F", "/IM", exe_name, "/T"],
                    capture_output=True, timeout=5,
                    creationflags=0x08000000
                )
                _time.sleep(0.4)
            # 用wmic杀指定路径进程，避免powershell
            try:
                subprocess.run(
                    ["wmic", "process", "where", f"ExecutablePath='{target_path}'", "delete"],
                    capture_output=True, timeout=8,
                    creationflags=0x08000000
                )
            except Exception:
                pass
            _time.sleep(0.5)
        except Exception:
            pass

    def _safe_copy(self, src, dst, max_retries=5):
        """带重试的文件复制，处理文件被占用的情况"""
        import time as _time
        for attempt in range(max_retries):
            try:
                # 如果目标文件存在，先尝试删除
                if os.path.exists(dst):
                    try:
                        os.remove(dst)
                    except PermissionError:
                        self._force_kill_file_lock(dst)
                        _time.sleep(1)
                        try:
                            os.remove(dst)
                        except Exception:
                            pass
                shutil.copy2(src, dst)
                return True
            except PermissionError:
                self._force_kill_file_lock(dst)
                self._force_kill_file_lock(src)
                _time.sleep(1.5)
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                _time.sleep(1)
        return False

    def _step_copy_exe(self):
        # 复制主程序 exe 到安装目录（带重试，强制结束占用进程）
        main_exe = resource_path(f"{APP_NAME}.exe")
        dst_exe = os.path.join(self.install_dir, f"{APP_NAME}.exe")
        if os.path.exists(main_exe):
            # 再次确保进程已结束
            self._force_kill_file_lock(dst_exe)
            ok = self._safe_copy(main_exe, dst_exe)
            if ok:
                self._log("✓ 主程序已复制到安装目录")
            else:
                self._log("⚠ 主程序复制失败，请手动关闭游戏菜单后重试")
        else:
            self._log("⚠ 主程序 exe 未找到")


    def _step_copy_res(self):
        icon_file = resource_path("app_icon.ico")
        if os.path.exists(icon_file):
            try:
                self._safe_copy(icon_file, os.path.join(self.install_dir, "app_icon.ico"))
            except Exception:
                shutil.copy2(icon_file, os.path.join(self.install_dir, "app_icon.ico"))
        uninstall_exe = resource_path("卸载游戏菜单.exe")
        if os.path.exists(uninstall_exe):
            try:
                self._safe_copy(uninstall_exe, os.path.join(self.install_dir, "卸载游戏菜单.exe"))
            except Exception:
                shutil.copy2(uninstall_exe, os.path.join(self.install_dir, "卸载游戏菜单.exe"))

    def _step_shortcuts(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop):
            try:
                import pythoncom
                from win32com.client import Dispatch
                pythoncom.CoInitialize()
                shell = Dispatch('WScript.Shell')
                desktop = shell.SpecialFolders("Desktop")
                pythoncom.CoUninitialize()
            except Exception:
                desktop = r"C:\Users\Public\Desktop"

        # 主程序 exe 在安装目录
        main_exe = os.path.join(self.install_dir, f"{APP_NAME}.exe")
        icon_path = os.path.join(self.install_dir, "app_icon.ico")
        shortcut_path = os.path.join(desktop, f"{APP_NAME}.lnk")
        create_shortcut(shortcut_path, main_exe, self.install_dir, icon_path, "")
        # 开机自启：启动文件夹快捷方式（指向打包后的exe，不是pythonw+脚本）
        try:
            startup = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
            os.makedirs(startup, exist_ok=True)
            create_shortcut(os.path.join(startup, f"{APP_NAME}.lnk"), main_exe, self.install_dir, icon_path, "")
        except Exception:
            pass
        # 开机自启：同时写入注册表 Run 项（双保险）
        try:
            import winreg as _wr
            _run_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            _key = _wr.OpenKey(_wr.HKEY_CURRENT_USER, _run_path, 0, _wr.KEY_SET_VALUE)
            _wr.SetValueEx(_key, APP_NAME, 0, _wr.REG_SZ, f'"{main_exe}"')
            _wr.CloseKey(_key)
        except Exception:
            pass

    def _step_done(self):
        import json as _json
        # 从临时备份恢复完整配置
        try:
            tmp = getattr(self, '_backup_cfg_tmp', None)
            if tmp and os.path.exists(tmp):
                dst_cfg = os.path.join(self.install_dir, "game_menu_config.json")
                shutil.copy2(tmp, dst_cfg)
                self._log("✓ 已恢复软件列表和图标位置")
                try:
                    os.remove(tmp)
                except Exception:
                    pass
        except Exception as e:
            self._log("⚠ 恢复配置失败: " + str(e))
        # 强制开启开机自启（防止备份cfg里autostart=False覆盖）
        try:
            dst_cfg = os.path.join(self.install_dir, "game_menu_config.json")
            if os.path.exists(dst_cfg):
                with open(dst_cfg, 'r', encoding='utf-8') as f:
                    _cfg = _json.load(f)
                _cfg["autostart"] = True
                with open(dst_cfg, 'w', encoding='utf-8') as f:
                    _json.dump(_cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        # 重新写入注册表Run项和启动文件夹快捷方式（确保开机自启）
        try:
            import winreg as _wr
            main_exe = os.path.join(self.install_dir, f"{APP_NAME}.exe")
            _key = _wr.OpenKey(_wr.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, _wr.KEY_SET_VALUE)
            _wr.SetValueEx(_key, APP_NAME, 0, _wr.REG_SZ, f'"{main_exe}"')
            _wr.CloseKey(_key)
            self._log("✓ 已设置开机自动启动（注册表）")
        except Exception as e:
            self._log("⚠ 注册表写入失败: " + str(e))
        # 启动文件夹快捷方式
        try:
            import os as _os
            startup = _os.path.join(_os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
            _os.makedirs(startup, exist_ok=True)
            lnk_path = _os.path.join(startup, f"{APP_NAME}.lnk")
            main_exe = _os.path.join(self.install_dir, f"{APP_NAME}.exe")
            icon_path = _os.path.join(self.install_dir, "app_icon.ico")
            create_shortcut(lnk_path, main_exe, self.install_dir, icon_path, "")
            self._log("✓ 已设置开机自动启动（启动文件夹）")
        except Exception as e:
            self._log("⚠ 启动文件夹写入失败: " + str(e))
        # 自动添加杀毒软件信任区/白名单
        try:
            import winreg
            _main_exe_trust = os.path.join(self.install_dir, APP_NAME + ".exe")
            trust_items = [self.install_dir, _main_exe_trust]
            for reg_path in [r"Software\Huorong\Sysdiag\trust", r"Software\Huorong\Sysdiag\allow"]:
                try:
                    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)
                    for d in trust_items:
                        winreg.SetValueEx(key, d, 0, winreg.REG_SZ, "1")
                    winreg.CloseKey(key)
                except Exception:
                    pass
            self._log("✓ 已添加火绒信任区")
        except Exception:
            pass
        # Windows Defender 白名单（Add-MpPreference，确定有效）
        try:
            import subprocess
            _main_exe_df = os.path.join(self.install_dir, APP_NAME + ".exe")
            for _item in [self.install_dir, _main_exe_df]:
                try:
                    subprocess.run(
                        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                         "Add-MpPreference -ExclusionPath '%s'" % _item],
                        creationflags=0x08000000, timeout=20)
                except Exception:
                    pass
            self._log("✓ 已添加 Windows Defender 白名单")
        except Exception:
            pass
        # 清除兼容性"以管理员身份运行"(RUNASADMIN)标志。
        # 该标志会让主程序以高完整性运行，普通权限的资源管理器无法向窗口拖放文件
        # （Windows UIPI 拦截，显示禁止图标），是拖放失效的根因。只删主程序这一项。
        try:
            import winreg as _wr2
            main_exe_path = os.path.join(self.install_dir, APP_NAME + ".exe")
            _layers = r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
            try:
                _k = _wr2.OpenKey(_wr2.HKEY_CURRENT_USER, _layers, 0,
                                  _wr2.KEY_READ | _wr2.KEY_WRITE)
            except OSError:
                _k = None
            if _k is not None:
                try:
                    _v, _ = _wr2.QueryValueEx(_k, main_exe_path)
                    if _v and "RUNASADMIN" in str(_v).upper():
                        _tokens = [t for t in str(_v).replace("~", " ").split()
                                   if t and t.upper() != "RUNASADMIN"]
                        if _tokens:
                            _wr2.SetValueEx(_k, main_exe_path, 0, _wr2.REG_SZ,
                                            "~ " + " ".join(_tokens))
                        else:
                            try:
                                _wr2.DeleteValue(_k, main_exe_path)
                            except OSError:
                                _wr2.SetValueEx(_k, main_exe_path, 0, _wr2.REG_SZ, "")
                        self._log("✓ 已清除管理员兼容性标志（保证拖放可用）")
                except OSError:
                    pass
                finally:
                    _wr2.CloseKey(_k)
        except Exception:
            pass


class InstallPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("")
        self.setSubTitle("")
        self.setFixedSize(680, 480)

        layout = QVBoxLayout()
        layout.addStretch(1)
        self.stage_title = QLabel("正在安装")
        self.stage_title.setAlignment(Qt.AlignCenter)
        self.stage_title.setStyleSheet("""
            color:#ff88dd; font-size:26px; font-weight:bold;
            font-family:'Microsoft YaHei'; letter-spacing:3px;
        """)
        layout.addWidget(self.stage_title)
        self.stage_lbl = QLabel("加载安装数据文件…")
        self.stage_lbl.setAlignment(Qt.AlignCenter)
        self.stage_lbl.setStyleSheet("color:#cfeeff; font-size:14px;")
        layout.addWidget(self.stage_lbl)
        layout.addSpacing(24)
        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setMaximum(100)
        self.progress.setTextVisible(True)
        self.progress.setMinimumHeight(36)
        self.progress.setValue(0)
        self.progress.setStyleSheet("""
            QProgressBar { background: rgba(255,255,255,15); border: 1px solid rgba(255,51,170,150); border-radius: 18px; text-align: center; color: #fff; font-size: 15px; font-weight: bold; }
            QProgressBar::chunk { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #ff51aa, stop:0.3 #ffd76b, stop:0.6 #51e6ff, stop:1 #51ffaa); border-radius: 16px; }
        """)
        layout.addWidget(self.progress)
        layout.addStretch(3)
        self.setLayout(layout)

        self.install_done = False
        self.install_dir = INSTALL_DIR
        self._started = False
        self.worker = None
        self.target_progress = 0

        # 注意：安装不能在 __init__ 启动。addPage() 会实例化全部页面，
        # 若在这里启动，向导刚打开就后台静默安装、装完直接弹主程序，跳过整个向导。
        # 安装统一在进入本页 initializePage() 时才开始。

        # 持续禁用"上一步"
        self._disable_back_timer = QTimer()
        self._disable_back_timer.timeout.connect(self._disable_back)
        self._disable_back_timer.start(300)

        # 进度条平滑动画
        self.anim_timer = QTimer()
        self.anim_timer.timeout.connect(self._animate_progress)
        self.anim_timer.start(30)

    def _disable_back(self):
        try:
            wiz = self.wizard()
            if not wiz:
                return
            for _b in (QWizard.BackButton, QWizard.NextButton, QWizard.CancelButton, QWizard.FinishButton):
                _btn = wiz.button(_b)
                if _btn:
                    _btn.hide()
        except Exception:
            pass

    def initializePage(self):
        # 进入安装进度页：隐藏按钮
        try:
            wiz = self.wizard()
            if wiz:
                for btn_id in (QWizard.BackButton, QWizard.NextButton, QWizard.CancelButton):
                    btn = wiz.button(btn_id)
                    if btn:
                        btn.hide()
        except Exception:
            pass
        QTimer.singleShot(300, self._start_install)

    def cleanupPage(self):
        # 离开安装进度页：恢复按钮
        try:
            wiz = self.wizard()
            if wiz:
                for btn_id in (QWizard.BackButton, QWizard.NextButton, QWizard.CancelButton):
                    btn = wiz.button(btn_id)
                    if btn:
                        btn.show()
        except Exception:
            pass

    def isComplete(self):
        return self.install_done

    def _start_install(self):
        if self._started:
            return
        self._started = True
        try:
            if self.wizard():
                _btn = self.wizard().button(QWizard.BackButton)
                if _btn:
                    _btn.setEnabled(False)
        except Exception:
            pass
        # 获取安装目录
        try:
            if self.wizard():
                for page_id in self.wizard().pageIds():
                    page = self.wizard().page(page_id)
                    if hasattr(page, 'get_install_dir'):
                        self.install_dir = page.get_install_dir()
                        break
        except Exception:
            pass
        # 启动后台安装线程
        self.worker = InstallWorker(self.install_dir)
        self.worker.log_signal.connect(self._log)
        self.worker.progress_signal.connect(self._set_target)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.start()

    def _set_target(self, v):
        self.target_progress = int(v)

    def _animate_progress(self):
        try:
            cur = self.progress.value()
            if cur < self.target_progress:
                self.progress.setValue(min(cur + 1, self.target_progress))
            elif cur > self.target_progress:
                self.progress.setValue(self.target_progress)
        except Exception:
            pass

    def _is_app_running(self):
        try:
            chk = subprocess.run(
                ["tasklist", "/FI", f"IMAGENAME eq {APP_NAME}.exe"],
                capture_output=True, text=True, encoding="gbk", errors="ignore", timeout=5,
                creationflags=0x08000000)
            return APP_NAME in (chk.stdout or "")
        except Exception:
            return False

    def _launch_app(self):
        """安装完成后自动启动游戏菜单（ShellExecute，避免权限继承报错740）"""
        exe_path = os.path.join(self.install_dir, f"{APP_NAME}.exe")
        try:
            if os.path.exists(exe_path) and not self._is_app_running():
                if launch_app_shell(exe_path, self.install_dir):
                    self._log("✓ 游戏菜单已自动启动")
                else:
                    self._log("⚠ 自动启动失败，请手动从桌面启动")
        except Exception as e:
            self._log("⚠ 自动启动失败: " + str(e))

    def _on_finished(self):
        self.install_done = True
        self.target_progress = 100
        self.progress.setValue(100)
        try:
            self._disable_back_timer.stop()
        except Exception:
            pass
        # 安装完成后显示"下一步"和"取消"按钮
        try:
            wiz = self.wizard()
            if wiz:
                for btn_id in (QWizard.NextButton, QWizard.CancelButton):
                    btn = wiz.button(btn_id)
                    if btn:
                        btn.show()
                        btn.setEnabled(True)
        except Exception:
            pass
        self.completeChanged.emit()
        # 不自动启动游戏菜单，点完成才启动
        self._log("✓ 安装完成")
        QTimer.singleShot(1200, self._goto_finish)

    def _goto_finish(self):
        try:
            wiz = self.wizard()
            if wiz:
                wiz.next()
                # 延迟3秒自动点完成按钮，点完成后直接启动游戏菜单
                def _auto_finish():
                    try:
                        btn = wiz.button(QWizard.FinishButton)
                        if btn:
                            btn.click()
                    except Exception:
                        pass
                    # 直接启动游戏菜单（不依赖finished信号）
                    try:
                        inst_dir = r"D:\游戏菜单"
                        for p in wiz.pageIds():
                            pg = wiz.page(p)
                            if hasattr(pg, 'get_install_dir'):
                                inst_dir = pg.get_install_dir()
                                break
                        exe_path = os.path.join(inst_dir, f"{APP_NAME}.exe")
                        if os.path.exists(exe_path):
                            launch_app_shell(exe_path, inst_dir)
                        # 退出安装程序
                        wiz.reject()
                    except Exception:
                        pass
                QTimer.singleShot(3000, _auto_finish)
        except Exception:
            pass

    def _log(self, msg):
        try:
            self.stage_lbl.setText(str(msg))
        except Exception:
            pass


class FinishPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("安装完成")
        self.setSubTitle(f"雪之下倾城游戏菜单已成功安装到您的电脑\n版权所有 © 2026 雪之下倾城  版本 {VERSION_DISPLAY}")
        self.setFixedSize(680, 560)

        layout = QVBoxLayout()
        info = QLabel(
            f"\n"
            f"  ✓ 程序文件已复制\n"
            f"  ✓ 桌面快捷方式已创建\n"
            f"  ✓ 卸载快捷方式已创建\n\n"
            f"  安装目录: {INSTALL_DIR}\n\n"
            f"  您可以通过以下方式启动游戏菜单：\n"
            f"  • 双击桌面「{APP_NAME}」快捷方式\n"
            f"  • 从安装目录运行「{APP_NAME}.exe」\n\n"
            f"  点击「完成」将自动启动游戏菜单。"
        )
        info.setFont(QFont("Microsoft YaHei", 12))
        info.setStyleSheet("color:#e8e8e8; line-height:1.6;")
        info.setWordWrap(True)
        layout.addWidget(info)
        self.setLayout(layout)

    def isComplete(self):
        return True


class InstallWizard(QWizard):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} - 安装程序")
        self.setFixedSize(700, 640)
        self.setWizardStyle(QWizard.ModernStyle)

        # 朋克主题样式（精致美观）
        self.setStyleSheet("""
            QWizard {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #2a1050, stop:0.5 #4a2080, stop:1 #2a1050);
            }
            QWizard * {
                background: transparent;
            }
            QWizard QFrame {
                background: transparent;
            }
            QWizardPage {
                background: transparent;
            }
            QWizard::title {
                color:#ff88dd; font-size:22px; font-weight:bold;
                padding:12px; background:rgba(255,102,204,20);
                border-bottom:2px solid rgba(255,102,204,120);
            }
            QWizard::subTitle {
                color:#ddaaff; font-size:14px; padding:6px 12px;
            }
            QLabel {
                color:#ffeeff; font-size:14px; font-family:'Microsoft YaHei';
                background: transparent;
            }
            QPushButton {
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ff33aa, stop:1 #cc3399);
                color:#ffffff; border:2px solid rgba(255,136,221,180); border-radius:10px;
                padding:8px 20px; font-size:13px; font-weight:bold;
                min-width:80px;
                min-height:32px;
            }
            QPushButton:hover {
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ff66cc, stop:1 #ee44aa);
                border-color:#ffaaee;
            }
            QPushButton:pressed {
                background:qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #cc3399, stop:1 #ff33aa);
            }
            QPushButton:disabled {
                background:#444; color:#888; border-color:#555;
            }
            QLineEdit {
                background:rgba(20,10,40,200); color:#ffffff;
                border:2px solid rgba(204,102,255,120); border-radius:8px;
                padding:10px 14px; font-size:14px;
            }
            QLineEdit:focus { border-color:#ff66cc; }
            QProgressBar {
                background:rgba(20,10,40,200); border:2px solid rgba(204,102,255,120);
                border-radius:10px; text-align:center; color:#ffffff;
                height:36px; font-size:15px; font-weight:bold;
            }
            QProgressBar::chunk {
                background:qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #ff33aa, stop:0.5 #00ffff, stop:1 #ff33aa);
                border-radius:8px;
            }
            QTextEdit {
                background:rgba(10,5,25,230); color:#88ffdd;
                border:2px solid rgba(204,102,255,120); border-radius:8px;
                font-family:'Consolas'; font-size:13px;
            }
            QCheckBox {
                color:#e8b8d8; font-size:13px;
            }
        """)

        self.setButtonText(QWizard.BackButton, "上一步")
        self.setButtonText(QWizard.NextButton, "下一步")
        self.setButtonText(QWizard.FinishButton, "完成")
        self.setButtonText(QWizard.CancelButton, "取消")

        self.addPage(WelcomePage())
        self.addPage(InstallPage())
        self.addPage(FinishPage())
        self.finished.connect(self._after_install)
        self._drag_pos = None

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self._drag_pos and event.buttons() & Qt.LeftButton:
            self.move(event.globalPos() - self._drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_pos = None

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)

    def closeEvent(self, event):
        """拦截关闭按钮：安装进行中时给确认提示，防止误关导致安装中断"""
        try:
            for pid in self.pageIds():
                pg = self.page(pid)
                if isinstance(pg, InstallPage):
                    if pg._started and not pg.install_done:
                        # 安装进行中
                        reply = QMessageBox.question(
                            self, "确认退出",
                            "安装正在进行中，关闭将中断安装！\n确定要退出吗？",
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.No
                        )
                        if reply != QMessageBox.Yes:
                            event.ignore()
                            return
                    elif pg._started and pg.install_done:
                        # 安装已完成，正常关闭即可
                        pass
                    break
        except Exception:
            pass
        event.accept()

    def _after_install(self, result):
        """安装结束后自动启动游戏菜单"""
        inst_dir = r"D:\游戏菜单"
        try:
            for p in self.pageIds():
                pg = self.page(p)
                if hasattr(pg, 'get_install_dir'):
                    inst_dir = pg.get_install_dir()
                    break
        except Exception:
            pass
        exe_path = os.path.join(inst_dir, f"{APP_NAME}.exe")
        try:
            log_path = os.path.join(inst_dir, "install_autostart.log")
            with open(log_path, "w", encoding="utf-8") as _f:
                _f.write(f"安装目录: {inst_dir}\n")
                _f.write(f"exe路径: {exe_path}\n")
                _f.write(f"exe存在: {os.path.exists(exe_path)}\n")
            if os.path.exists(exe_path):
                # ShellExecute 启动
                ok = launch_app_shell(exe_path, inst_dir)
                if not ok:
                    # 备选：os.startfile
                    try:
                        os.startfile(exe_path, cwd=inst_dir)
                        ok = True
                    except Exception:
                        pass
                with open(log_path, "a", encoding="utf-8") as _f:
                    _f.write("启动成功！\n" if ok else "启动失败\n")
            else:
                with open(log_path, "a", encoding="utf-8") as _f:
                    _f.write("启动失败：exe不存在\n")
        except Exception as e:
            try:
                log_path = os.path.join(inst_dir, "install_autostart.log")
                with open(log_path, "a", encoding="utf-8") as _f:
                    _f.write(f"启动异常: {repr(e)}\n")
            except Exception:
                pass


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(f"{APP_NAME}安装程序")

    # 加载Qt中文翻译（QMessageBox等标准按钮中文化：确定/取消/是/否）
    try:
        from PySide6.QtCore import QTranslator
        _trans = QTranslator()
        _meipass = getattr(sys, "_MEIPASS", "")
        _candidates = [
            resource_path(os.path.join("translations", "qtbase_zh_CN.qm")),
            os.path.join(_meipass, "PySide6", "translations", "qtbase_zh_CN.qm") if _meipass else "",
        ]
        for _tp in _candidates:
            if _tp and os.path.isfile(_tp) and _trans.load(_tp):
                app.installTranslator(_trans)
                break
    except Exception:
        pass

    icon_path = resource_path("app_icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    wizard = InstallWizard()
    wizard.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
