"""
工作日报助手 - Fluent Design 主程序

使用 PyQt-Fluent-Widgets 库实现现代化界面
支持响应式布局，随窗口大小自动调整
支持系统DPI缩放检测和手动调整
"""

import sys
import os
import platform
import sys; sys.setrecursionlimit(sys.getrecursionlimit() * 5)

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 顶部导入 PyQt5 组件（供 LoginWindow 使用）
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QApplication, QCheckBox, QGraphicsOpacityEffect, QTextBrowser)
from PyQt5.QtCore import Qt, pyqtSignal, QPoint, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QPixmap, QPainter, QPainterPath, QLinearGradient, QPen, QBrush, QColor
from datetime import datetime, timedelta, timezone

# 东八区时区
CST = timezone(timedelta(hours=8))


def get_now():
    """获取东八区当前时间"""
    return datetime.now(CST)


def get_today():
    """获取东八区今天的日期字符串"""
    return get_now().strftime('%Y-%m-%d')


def get_current_time():
    """获取东八区当前时间字符串"""
    return get_now().strftime('%H:%M:%S')


def _write_xlsx(file_path, sheets):
    """
    用标准库生成 .xlsx 文件（无第三方依赖，Excel/WPS 可直接打开）

    参数:
        file_path: 输出文件路径
        sheets: 列表，每个元素为 dict:
            {
                "name":   工作表名称,
                "headers": 表头列表（可为空列表）,
                "rows":    二维数据列表（每行一个列表）,
                "widths":  可选，每列宽度列表
            }
    """
    import zipfile
    from xml.sax.saxutils import escape

    def col_letter(idx):
        """列序号转字母：0->A, 25->Z, 26->AA"""
        letters = ""
        idx += 1
        while idx:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    def cell_xml(r, c, value, style=0):
        ref = f"{col_letter(c)}{r}"
        if isinstance(value, bool):
            value = 1 if value else 0
        if isinstance(value, (int, float)):
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
        text = escape(str(value))
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'

    def safe_sheet_name(name):
        for ch in '[]:*?/\\':
            name = name.replace(ch, '_')
        return (name[:31] or "Sheet")

    sheet_names = [safe_sheet_name(s["name"]) for s in sheets]

    # ---- 生成每个工作表的 XML ----
    worksheets_xml = []
    for s in sheets:
        headers = s.get("headers") or []
        rows = s.get("rows") or []
        widths = s.get("widths") or []
        lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        lines.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        if widths:
            cols = "".join(
                f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
                for i, w in enumerate(widths)
            )
            lines.append(f"<cols>{cols}</cols>")
        lines.append("<sheetData>")
        r = 1
        if headers:
            cells = "".join(cell_xml(r, c, h, style=1) for c, h in enumerate(headers))
            lines.append(f'<row r="{r}">{cells}</row>')
            r += 1
        for row in rows:
            cells = "".join(cell_xml(r, c, v, style=0) for c, v in enumerate(row))
            lines.append(f'<row r="{r}">{cells}</row>')
            r += 1
        lines.append("</sheetData>")
        lines.append("</worksheet>")
        worksheets_xml.append("\n".join(lines))

    # ---- 固定 XML 部件 ----
    n_sheets = len(sheets)
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, n_sheets + 1)
        )
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets>'
        + "".join(
            f'<sheet name="{escape(name)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
            for i, name in enumerate(sheet_names)
        )
        + '</sheets></workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
            for i in range(n_sheets)
        )
        + f'<Relationship Id="rId{n_sheets + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2563EB"/></patternFill></fill>'
        '</fills>'
        '<borders count="1">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )

    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', workbook)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/styles.xml', styles)
        for i, xml in enumerate(worksheets_xml, start=1):
            zf.writestr(f'xl/worksheets/sheet{i}.xml', xml)


SCALE_FACTOR = 1.0


def load_scale_setting():
    try:
        config_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'config.txt')
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.startswith('scale_factor='):
                        return float(line.strip().split('=')[1])
    except:
        pass
    return None


def get_system_dpi_scale():
    try:
        if platform.system() == 'Windows':
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(2)
                hdc = ctypes.windll.user32.GetDC(0)
                dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
                ctypes.windll.user32.ReleaseDC(0, hdc)
                scale = dpi / 96.0
                return round(scale * 4) / 4
            except:
                return 1.0
        else:
            return 1.0
    except:
        return 1.0


# ==================== 登录窗口 ====================

# API 配置
API_BASE_URL = "http://129.204.12.226"

class LoginWindow(QWidget):
    """登录/注册窗口 - 高端美观设计"""
    login_success = pyqtSignal()  # 登录成功信号
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.is_login_mode = True  # True=登录模式, False=注册模式
        self.is_forgot_mode = False  # True=找回密码模式
        self._is_logging_in = False  # 是否正在登录（防止closeEvent误触发）
        self._anim_widgets = []  # 用于入场动画的控件列表
        self.setup_ui()
        self._load_remember_account()
        # 延迟播放入场动画
        QTimer.singleShot(50, self._play_entrance_animations)
    
    def paintEvent(self, event):
        """绘制渐变背景和装饰元素"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        w, h = self.width(), self.height()
        
        # 主渐变背景：从浅绿到白
        gradient = QLinearGradient(0, 0, w, h)
        gradient.setColorAt(0.0, QColor("#E8F5E9"))
        gradient.setColorAt(0.5, QColor("#F1F8F2"))
        gradient.setColorAt(1.0, QColor("#FFFFFF"))
        painter.fillRect(self.rect(), gradient)
        
        # 顶部装饰大圆（右上角，浅绿）
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(22, 163, 74, 18))
        painter.drawEllipse(w - 140, -100, 280, 280)
        
        # 左上角小圆
        painter.setBrush(QColor(22, 163, 74, 25))
        painter.drawEllipse(-60, -60, 160, 160)
        
        # 左下角装饰圆
        painter.setBrush(QColor(76, 175, 80, 15))
        painter.drawEllipse(-80, h - 120, 220, 220)
        
        # 右下角小圆点装饰
        painter.setBrush(QColor(22, 163, 74, 30))
        painter.drawEllipse(w - 60, h - 80, 40, 40)
        painter.setBrush(QColor(76, 175, 80, 20))
        painter.drawEllipse(w - 110, h - 40, 24, 24)
        
        # 顶部细绿条装饰
        painter.setBrush(QColor(22, 163, 74, 200))
        painter.drawRoundedRect(0, 0, w, 4, 0, 0)
        
        painter.end()
    
    def _play_entrance_animations(self):
        """播放入场动画（交错淡入+上移）"""
        for i, widget in enumerate(self._anim_widgets):
            # 淡入
            opacity = QGraphicsOpacityEffect(widget)
            opacity.setOpacity(0)
            widget.setGraphicsEffect(opacity)
            
            fade = QPropertyAnimation(opacity, b"opacity")
            fade.setDuration(500)
            fade.setStartValue(0.0)
            fade.setEndValue(1.0)
            fade.setEasingCurve(QEasingCurve.OutCubic)
            
            # 保存引用防止被回收
            if not hasattr(self, '_entrance_anims'):
                self._entrance_anims = []
            self._entrance_anims.append(opacity)
            self._entrance_anims.append(fade)
            
            QTimer.singleShot(i * 90, fade.start)
    
    def setup_ui(self):
        """设置界面"""
        self.setWindowTitle("工作日报助手 - 登录")
        self.setFixedSize(440, 580)
        self.setAttribute(Qt.WA_StyledBackground, True)
        
        # 主布局
        mainLayout = QVBoxLayout(self)
        mainLayout.setContentsMargins(44, 36, 44, 28)
        mainLayout.setSpacing(0)
        
        # ========== 顶部 Logo 区域 ==========
        logoLayout = QHBoxLayout()
        logoLayout.setSpacing(14)
        
        # Logo 图标（使用图片，带绿色光环）
        logoIcon = QLabel()
        logoIcon.setFixedSize(52, 52)
        logoIcon.setAlignment(Qt.AlignCenter)
        
        avatar_path = r"C:\Users\20057\Desktop\frog.jpg"
        if os.path.exists(avatar_path):
            pixmap = QPixmap(avatar_path)
            pixmap = pixmap.scaled(44, 44, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
            
            # 创建带绿色光环的圆形头像
            rounded = QPixmap(52, 52)
            rounded.fill(Qt.transparent)
            painter = QPainter(rounded)
            painter.setRenderHint(QPainter.Antialiasing)
            # 外圈绿色光环
            painter.setPen(QPen(QColor("#16A34A"), 3))
            painter.setBrush(QBrush(QColor("#E8F5E9")))
            painter.drawEllipse(2, 2, 48, 48)
            # 内圈头像
            path = QPainterPath()
            path.addEllipse(6, 6, 40, 40)
            painter.setClipPath(path)
            painter.setPen(Qt.NoPen)
            painter.drawPixmap(6, 6, pixmap)
            painter.end()
            
            logoIcon.setPixmap(rounded)
        else:
            logoIcon.setStyleSheet("""
                QLabel {
                    background-color: #16A34A;
                    border-radius: 26px;
                    font-size: 26px;
                    color: white;
                }
            """)
            logoIcon.setText("🐸")
        
        logoLayout.addWidget(logoIcon)
        
        # 标题
        titleLayout = QVBoxLayout()
        titleLayout.setSpacing(3)
        
        titleLabel = QLabel("工作日报助手")
        titleLabel.setStyleSheet("font-size: 22px; font-weight: 800; color: #14532D; border: none; background: transparent; letter-spacing: 1px;")
        titleLayout.addWidget(titleLabel)
        
        subtitleLabel = QLabel("AI 驱动的智能工作报告工具")
        subtitleLabel.setStyleSheet("font-size: 11px; color: #6B7280; border: none; background: transparent; letter-spacing: 0.5px;")
        titleLayout.addWidget(subtitleLabel)
        
        logoLayout.addLayout(titleLayout)
        logoLayout.addStretch()
        
        mainLayout.addLayout(logoLayout)
        self._anim_widgets.append(logoIcon)
        mainLayout.addSpacing(28)
        
        # ========== 登录/注册标题行 ==========
        headerLayout = QHBoxLayout()
        headerLayout.setSpacing(8)
        
        self.modeTitle = QLabel("登录")
        self.modeTitle.setStyleSheet("font-size: 26px; font-weight: 800; color: #1a1a1a; border: none; background: transparent;")
        headerLayout.addWidget(self.modeTitle)
        
        # 标题下方绿色装饰条
        self.titleAccent = QLabel()
        self.titleAccent.setFixedSize(36, 4)
        self.titleAccent.setStyleSheet("background-color: #16A34A; border-radius: 2px; border: none;")
        
        headerLayout.addStretch()
        
        # 切换按钮
        self.switchBtn = QPushButton("注册")
        self.switchBtn.setCursor(Qt.PointingHandCursor)
        self.switchBtn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #16A34A;
                font-size: 14px;
                font-weight: bold;
                border: none;
                padding: 4px 10px;
                border-radius: 6px;
            }
            QPushButton:hover {
                color: #15803D;
                background-color: rgba(22, 163, 74, 0.08);
            }
        """)
        self.switchBtn.clicked.connect(self.toggle_mode)
        headerLayout.addWidget(self.switchBtn)
        
        mainLayout.addLayout(headerLayout)
        mainLayout.addWidget(self.titleAccent)
        mainLayout.addSpacing(6)
        
        # 副标题
        self.modeSubtitle = QLabel("请输入您的账号信息")
        self.modeSubtitle.setStyleSheet("font-size: 12px; color: #6B7280; border: none; background: transparent;")
        mainLayout.addWidget(self.modeSubtitle)
        mainLayout.addSpacing(22)
        
        # ========== 邮箱输入框 ==========
        emailLabel = QLabel("邮箱")
        emailLabel.setStyleSheet("font-size: 12px; font-weight: 700; color: #374151; border: none; background: transparent;")
        mainLayout.addWidget(emailLabel)
        mainLayout.addSpacing(8)
        
        emailContainer = QWidget()
        emailContainer.setStyleSheet("background: transparent; border: none;")
        emailLayout = QHBoxLayout(emailContainer)
        emailLayout.setContentsMargins(0, 0, 0, 0)
        emailLayout.setSpacing(8)
        
        self.emailInput = QLineEdit()
        self.emailInput.setPlaceholderText("请输入邮箱地址")
        self.emailInput.setFixedHeight(44)
        self.emailInput.setStyleSheet("""
            QLineEdit {
                background-color: white;
                border: 1.5px solid #E5E7EB;
                border-radius: 10px;
                padding: 0 14px;
                font-size: 13px;
                color: #1a1a1a;
            }
            QLineEdit:hover {
                border: 1.5px solid #A7F3D0;
            }
            QLineEdit:focus {
                border: 2px solid #16A34A;
                background-color: #FAFFFB;
            }
        """)
        emailLayout.addWidget(self.emailInput, 1)
        
        # 发送验证码按钮（注册/找回密码模式显示）
        self.sendCodeBtn = QPushButton("发送验证码")
        self.sendCodeBtn.setCursor(Qt.PointingHandCursor)
        self.sendCodeBtn.setFixedSize(104, 44)
        self.sendCodeBtn.setStyleSheet("""
            QPushButton {
                background-color: #16A34A;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #15803D;
            }
            QPushButton:pressed {
                background-color: #14532D;
            }
            QPushButton:disabled {
                background-color: #A7F3D0;
                color: #6EE7B7;
            }
        """)
        self.sendCodeBtn.clicked.connect(self.send_verification_code)
        self.sendCodeBtn.setVisible(False)
        emailLayout.addWidget(self.sendCodeBtn)
        
        mainLayout.addWidget(emailContainer)
        mainLayout.addSpacing(18)
        
        # ========== 密码输入框 ==========
        self.passwordLabel = QLabel("密码")
        self.passwordLabel.setStyleSheet("font-size: 12px; font-weight: 700; color: #374151; border: none; background: transparent;")
        mainLayout.addWidget(self.passwordLabel)
        mainLayout.addSpacing(8)
        
        self.passwordInput = QLineEdit()
        self.passwordInput.setPlaceholderText("请输入密码")
        self.passwordInput.setFixedHeight(44)
        self.passwordInput.setEchoMode(QLineEdit.Password)
        self.passwordInput.setStyleSheet("""
            QLineEdit {
                background-color: white;
                border: 1.5px solid #E5E7EB;
                border-radius: 10px;
                padding: 0 14px;
                font-size: 13px;
                color: #1a1a1a;
            }
            QLineEdit:hover {
                border: 1.5px solid #A7F3D0;
            }
            QLineEdit:focus {
                border: 2px solid #16A34A;
                background-color: #FAFFFB;
            }
        """)
        mainLayout.addWidget(self.passwordInput)
        mainLayout.addSpacing(10)
        
        # ========== 记住账号密码（登录模式显示）==========
        self.rememberCheckBox = QCheckBox("记住账号密码")
        self.rememberCheckBox.setStyleSheet("""
            QCheckBox {
                font-size: 12px;
                color: #6B7280;
                border: none;
                background: transparent;
                spacing: 6px;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 2px solid #D1D5DB;
                border-radius: 4px;
                background: white;
            }
            QCheckBox::indicator:hover {
                border-color: #16A34A;
            }
            QCheckBox::indicator:checked {
                background-color: #16A34A;
                border-color: #16A34A;
            }
        """)
        mainLayout.addWidget(self.rememberCheckBox)
        mainLayout.addSpacing(14)
        
        # ========== 验证码输入框（注册/找回密码模式显示）==========
        self.codeLabel = QLabel("验证码")
        self.codeLabel.setStyleSheet("font-size: 12px; font-weight: 700; color: #374151; border: none; background: transparent;")
        self.codeLabel.setVisible(False)
        mainLayout.addWidget(self.codeLabel)
        mainLayout.addSpacing(8)
        
        self.codeInput = QLineEdit()
        self.codeInput.setPlaceholderText("请输入验证码")
        self.codeInput.setFixedHeight(44)
        self.codeInput.setStyleSheet("""
            QLineEdit {
                background-color: white;
                border: 1.5px solid #E5E7EB;
                border-radius: 10px;
                padding: 0 14px;
                font-size: 13px;
                color: #1a1a1a;
                letter-spacing: 4px;
                font-weight: bold;
            }
            QLineEdit:hover {
                border: 1.5px solid #A7F3D0;
            }
            QLineEdit:focus {
                border: 2px solid #16A34A;
                background-color: #FAFFFB;
            }
        """)
        self.codeInput.setVisible(False)
        mainLayout.addWidget(self.codeInput)
        
        mainLayout.addSpacing(22)
        
        # ========== 登录/注册按钮 ==========
        self.submitBtn = QPushButton("登  录")
        self.submitBtn.setFixedHeight(46)
        self.submitBtn.setCursor(Qt.PointingHandCursor)
        self.submitBtn.setStyleSheet("""
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #16A34A, stop:1 #22C55E);
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 15px;
                font-weight: bold;
                letter-spacing: 2px;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #15803D, stop:1 #16A34A);
            }
            QPushButton:pressed {
                background-color: #14532D;
            }
            QPushButton:disabled {
                background-color: #A7F3D0;
                color: #6EE7B7;
            }
        """)
        self.submitBtn.clicked.connect(self.on_submit)
        mainLayout.addWidget(self.submitBtn)
        
        # ========== 忘记密码链接 ==========
        forgotLayout = QHBoxLayout()
        forgotLayout.addStretch()
        
        self.forgotBtn = QPushButton("忘记密码？")
        self.forgotBtn.setCursor(Qt.PointingHandCursor)
        self.forgotBtn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #16A34A;
                font-size: 12px;
                border: none;
                padding: 6px 4px;
            }
            QPushButton:hover {
                color: #15803D;
                text-decoration: underline;
            }
        """)
        self.forgotBtn.clicked.connect(self.toggle_forgot_mode)
        forgotLayout.addWidget(self.forgotBtn)
        forgotLayout.addStretch()
        
        mainLayout.addLayout(forgotLayout)
        mainLayout.addSpacing(8)
        
        # ========== 底部提示 ==========
        bottomLayout = QHBoxLayout()
        bottomLayout.setSpacing(4)
        
        tipLabel = QLabel("登录即表示同意")
        tipLabel.setStyleSheet("font-size: 11px; color: #9CA3AF; border: none; background: transparent;")
        bottomLayout.addWidget(tipLabel)
        
        termsBtn = QPushButton("服务条款")
        termsBtn.setCursor(Qt.PointingHandCursor)
        termsBtn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #16A34A;
                font-size: 11px;
                border: none;
                padding: 0;
            }
            QPushButton:hover {
                color: #15803D;
                text-decoration: underline;
            }
        """)
        bottomLayout.addWidget(termsBtn)
        
        andLabel = QLabel("和")
        andLabel.setStyleSheet("font-size: 11px; color: #9CA3AF; border: none; background: transparent;")
        bottomLayout.addWidget(andLabel)
        
        privacyBtn = QPushButton("隐私政策")
        privacyBtn.setCursor(Qt.PointingHandCursor)
        privacyBtn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #16A34A;
                font-size: 11px;
                border: none;
                padding: 0;
            }
            QPushButton:hover {
                color: #15803D;
                text-decoration: underline;
            }
        """)
        bottomLayout.addWidget(privacyBtn)
        bottomLayout.addStretch()
        
        mainLayout.addLayout(bottomLayout)
        
        # 收集需要入场动画的控件
        self._anim_widgets.extend([
            self.modeTitle, self.modeSubtitle,
            self.emailInput, self.passwordInput, self.submitBtn
        ])
    
    def toggle_mode(self):
        """切换登录/注册/找回密码模式"""
        if self.is_forgot_mode:
            # 从找回密码返回登录
            self.is_forgot_mode = False
            self.is_login_mode = True
        elif self.is_login_mode:
            # 从登录切换到注册
            self.is_login_mode = False
        else:
            # 从注册切换到登录
            self.is_login_mode = True
        
        # 更新UI
        if self.is_forgot_mode:
            self.toggle_forgot_mode()
        elif self.is_login_mode:
            self.modeTitle.setText("登录")
            self.switchBtn.setText("注册")
            self.switchBtn.setVisible(True)
            self.modeSubtitle.setText("请输入您的账号信息")
            self.submitBtn.setText("登录")
            self.sendCodeBtn.setVisible(False)
            self.codeLabel.setVisible(False)
            self.codeInput.setVisible(False)
            self.passwordInput.setVisible(True)
            self.passwordLabel.setVisible(True)
            self.passwordLabel.setText("密码")
            self.passwordInput.setPlaceholderText("请输入密码")
            self.passwordInput.setEchoMode(QLineEdit.Password)
            self.rememberCheckBox.setVisible(True)
            self.forgotBtn.setVisible(True)
        else:
            # 注册模式
            self.modeTitle.setText("注册")
            self.switchBtn.setText("登录")
            self.switchBtn.setVisible(True)
            self.modeSubtitle.setText("创建新账号开始使用")
            self.submitBtn.setText("注册")
            self.sendCodeBtn.setVisible(True)
            self.codeLabel.setVisible(True)
            self.codeInput.setVisible(True)
            self.passwordInput.setVisible(True)
            self.passwordLabel.setVisible(True)
            self.passwordLabel.setText("密码")
            self.passwordInput.setPlaceholderText("设置密码（至少6位）")
            self.passwordInput.setEchoMode(QLineEdit.Password)
            self.rememberCheckBox.setVisible(False)
            self.forgotBtn.setVisible(False)
        
        # 清空输入框
        self.emailInput.clear()
        self.passwordInput.clear()
        self.codeInput.clear()
    
    def toggle_forgot_mode(self):
        """切换到找回密码模式"""
        self.is_forgot_mode = True
        self.is_login_mode = False
        
        self.modeTitle.setText("找回密码")
        self.switchBtn.setText("返回登录")
        self.switchBtn.setVisible(True)
        self.modeSubtitle.setText("通过验证码重置您的密码")
        self.submitBtn.setText("重置密码")
        
        # 隐藏忘记密码按钮
        self.forgotBtn.setVisible(False)
        
        # 显示验证码输入框和发送按钮
        self.sendCodeBtn.setVisible(True)
        self.codeLabel.setVisible(True)
        self.codeInput.setVisible(True)
        
        # 修改密码输入框为新密码
        self.passwordLabel.setText("新密码")
        self.passwordInput.setPlaceholderText("设置新密码（至少6位）")
        self.passwordInput.setVisible(True)
        self.passwordLabel.setVisible(True)
        self.passwordInput.setEchoMode(QLineEdit.Password)
        self.rememberCheckBox.setVisible(False)
        
        # 清空输入框
        self.emailInput.clear()
        self.passwordInput.clear()
        self.codeInput.clear()
    
    def send_verification_code(self):
        """发送验证码"""
        import requests
        
        email = self.emailInput.text().strip()
        if not email:
            from qfluentwidgets import InfoBar, InfoBarPosition
            InfoBar.warning(
                title="提示",
                content="请输入邮箱地址",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
            return
        
        # 调用发送验证码接口
        self.sendCodeBtn.setEnabled(False)
        self.sendCodeBtn.setText("发送中...")
        
        try:
            response = requests.post(
                f"{API_BASE_URL}/api/send-code",
                json={"email": email},
                timeout=10
            )
            result = response.json()
            
            from qfluentwidgets import InfoBar, InfoBarPosition
            if result.get('success'):
                InfoBar.success(
                    title="发送成功",
                    content=result.get('message', f"验证码已发送到 {email}"),
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                self.sendCodeBtn.setText("已发送")
                # 60秒后重新启用
                from PyQt5.QtCore import QTimer
                QTimer.singleShot(60000, lambda: [
                    self.sendCodeBtn.setEnabled(True),
                    self.sendCodeBtn.setText("发送验证码")
                ])
            else:
                InfoBar.error(
                    title="发送失败",
                    content=result.get('message', '发送验证码失败'),
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                self.sendCodeBtn.setEnabled(True)
                self.sendCodeBtn.setText("发送验证码")
        except Exception as e:
            from qfluentwidgets import InfoBar, InfoBarPosition
            InfoBar.error(
                title="连接失败",
                content=f"无法连接到服务器: {str(e)}",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
            self.sendCodeBtn.setEnabled(True)
            self.sendCodeBtn.setText("发送验证码")
    
    def on_submit(self):
        """提交登录/注册/找回密码"""
        import requests
        
        email = self.emailInput.text().strip()
        password = self.passwordInput.text().strip()
        
        if not email:
            from qfluentwidgets import InfoBar, InfoBarPosition
            InfoBar.warning(
                title="提示",
                content="请输入邮箱地址",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
            return
        
        if self.is_forgot_mode:
            # 找回密码逻辑
            code = self.codeInput.text().strip()
            if not password:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.warning(
                    title="提示",
                    content="请输入新密码",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
                return
            if not code:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.warning(
                    title="提示",
                    content="请输入验证码",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
                return
            
            self.submitBtn.setEnabled(False)
            self.submitBtn.setText("重置中...")
            
            try:
                response = requests.post(
                    f"{API_BASE_URL}/api/reset-password",
                    json={"email": email, "code": code, "new_password": password},
                    timeout=10
                )
                result = response.json()
                
                from qfluentwidgets import InfoBar, InfoBarPosition
                if result.get('success'):
                    InfoBar.success(
                        title="重置成功",
                        content=result.get('message', '密码已重置，请重新登录'),
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                    self.toggle_mode()  # 返回登录模式
                else:
                    InfoBar.error(
                        title="重置失败",
                        content=result.get('message', '重置失败'),
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                self.submitBtn.setEnabled(True)
                self.submitBtn.setText("重置密码")
            except Exception as e:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.error(
                    title="连接失败",
                    content=f"无法连接到服务器: {str(e)}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                self.submitBtn.setEnabled(True)
                self.submitBtn.setText("重置密码")
        
        elif self.is_login_mode:
            # 登录逻辑
            if not password:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.warning(
                    title="提示",
                    content="请输入密码",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
                return
            
            self.submitBtn.setEnabled(False)
            self.submitBtn.setText("登录中...")
            
            try:
                response = requests.post(
                    f"{API_BASE_URL}/api/login",
                    json={"email": email, "password": password},
                    timeout=10
                )
                result = response.json()
                
                from qfluentwidgets import InfoBar, InfoBarPosition
                if result.get('success'):
                    InfoBar.success(
                        title="登录成功",
                        content=result.get('message', '欢迎回来'),
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=2000,
                        parent=self
                    )
                    # 保存登录状态（包含密码用于数据同步）
                    self.save_login_state(email, password)
                    # 记住账号密码到 secret.json
                    self._save_remember_account(email, password)
                    self._is_logging_in = True  # 标记正在登录，防止closeEvent退出程序
                    self.login_success.emit()
                    self.close()
                else:
                    InfoBar.error(
                        title="登录失败",
                        content=result.get('message', '邮箱或密码错误'),
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                    self.submitBtn.setEnabled(True)
                    self.submitBtn.setText("登录")
            except Exception as e:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.error(
                    title="连接失败",
                    content=f"无法连接到服务器: {str(e)}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                self.submitBtn.setEnabled(True)
                self.submitBtn.setText("登录")
        else:
            # 注册逻辑
            code = self.codeInput.text().strip()
            if not password:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.warning(
                    title="提示",
                    content="请输入密码",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
                return
            if not code:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.warning(
                    title="提示",
                    content="请输入验证码",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
                return
            
            self.submitBtn.setEnabled(False)
            self.submitBtn.setText("注册中...")
            
            try:
                response = requests.post(
                    f"{API_BASE_URL}/api/register",
                    json={"email": email, "password": password, "code": code},
                    timeout=10
                )
                result = response.json()
                
                from qfluentwidgets import InfoBar, InfoBarPosition
                if result.get('success'):
                    InfoBar.success(
                        title="注册成功",
                        content=result.get('message', '请使用新账号登录'),
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                    self.toggle_mode()  # 切换回登录模式
                else:
                    InfoBar.error(
                        title="注册失败",
                        content=result.get('message', '注册失败'),
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                self.submitBtn.setEnabled(True)
                self.submitBtn.setText("注册")
            except Exception as e:
                from qfluentwidgets import InfoBar, InfoBarPosition
                InfoBar.error(
                    title="连接失败",
                    content=f"无法连接到服务器: {str(e)}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                self.submitBtn.setEnabled(True)
                self.submitBtn.setText("注册")
    
    def closeEvent(self, event):
        """关闭窗口时终止程序（仅在用户手动关闭时）"""
        if not self._is_logging_in:
            QApplication.quit()
        event.accept()
    
    def save_login_state(self, email, password=None):
        """保存登录状态"""
        import json
        from datetime import datetime
        config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
        os.makedirs(config_dir, exist_ok=True)
        config_file = os.path.join(config_dir, 'login_state.json')
        
        state = {
            'email': email,
            'login_time': get_now().strftime('%Y-%m-%d %H:%M:%S')
        }
        if password:
            state['password'] = password
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False)
    
    def check_login_state(self):
        """检查登录状态"""
        import json
        config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
        config_file = os.path.join(config_dir, 'login_state.json')
        
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    state = json.load(f)
                    return state.get('email')
            except:
                pass
        return None
    
    def _save_remember_account(self, email, password):
        """将记住的账号密码加密保存到 secret.json，取消勾选时删除记录"""
        import json
        config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
        secret_file = os.path.join(config_dir, 'secret.json')
        
        if not self.rememberCheckBox.isChecked():
            # 取消勾选，删除已保存的记录
            if os.path.exists(secret_file):
                os.remove(secret_file)
            return
        
        try:
            from crypto_utils import encrypt_text
            os.makedirs(config_dir, exist_ok=True)
            
            data = {
                'email': encrypt_text(email),
                'password': encrypt_text(password)
            }
            with open(secret_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception as e:
            print(f"保存记住账号失败: {e}")
    
    def _load_remember_account(self):
        """从 secret.json 加载已记住的账号密码并填充到输入框"""
        try:
            import json
            from crypto_utils import decrypt_text
            config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
            secret_file = os.path.join(config_dir, 'secret.json')
            
            if os.path.exists(secret_file):
                with open(secret_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                email = decrypt_text(data.get('email', ''))
                password = decrypt_text(data.get('password', ''))
                if email and password:
                    self.emailInput.setText(email)
                    self.passwordInput.setText(password)
                    self.rememberCheckBox.setChecked(True)
        except Exception as e:
            print(f"加载记住账号失败: {e}")


def main():
    global SCALE_FACTOR
    
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtCore import Qt, QSize, QTimer, QPropertyAnimation, QEasingCurve, QPoint
    from PyQt5.QtGui import QFont, QColor, QPixmap, QPainter, QPainterPath, QBrush, QPen, QIcon
    
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    
    system_scale = get_system_dpi_scale()
    saved_scale = load_scale_setting()
    SCALE_FACTOR = saved_scale if saved_scale else system_scale
    
    base_font_size = max(9, int(12 / SCALE_FACTOR))
    app.setFont(QFont("Microsoft YaHei", base_font_size))
    
    # 统一 ToolTip 样式：浅灰背景、更小字体、圆角细边框，观感更精致
    # （报告列表的 Token 悬停详情等多行提示均生效）
    _tooltip_font_size = max(10, min(13, base_font_size))
    app.setStyleSheet(f"""
        QToolTip {{
            background-color: #F5F5F5;
            color: #444444;
            font-size: {_tooltip_font_size}px;
            font-family: "Microsoft YaHei", sans-serif;
            border: 1px solid #DCDCDC;
            border-radius: 6px;
            padding: 5px 9px;
        }}
    """)
    
    from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                                 QLabel, QFrame, QScrollArea, QCheckBox,
                                 QGraphicsDropShadowEffect, QGraphicsOpacityEffect,
                                 QSizePolicy, QPushButton, QTableWidget, QTableWidgetItem,
                                 QLineEdit, QDateEdit, QComboBox, QApplication,
                                 QMessageBox, QSystemTrayIcon, QMenu, QAction, QDialog,
                                 QTextEdit, QLayout, QSplitter, QTextBrowser, QSpinBox)
    from PyQt5.QtCore import Qt, QSize, QTimer, QDate, QPropertyAnimation, QEasingCurve, QDateTime, QThread, pyqtSignal, QRect, QPoint
    from PyQt5.QtGui import QFont, QColor, QPixmap, QPainter, QPainterPath, QBrush, QPen, QIcon
    from qfluentwidgets import (FluentWindow, NavigationItemPosition, StrongBodyLabel,
                                TitleLabel, SubtitleLabel, BodyLabel, CaptionLabel,
                                PrimaryPushButton, TransparentPushButton, PillPushButton,
                                SimpleCardWidget, HeaderCardWidget, TableWidget,
                                FluentIcon, ComboBox, CalendarPicker, SearchLineEdit,
                                InfoBar, InfoBarPosition, ToolButton, FluentIconBase)
    from store import init_db, get_daily_summary, get_daily_records, read_records
    from screenshot import run_and_store, get_today_stats, get_monitor_info, start_monitor, stop_monitor, set_use_glm, set_ollama_config, test_glm_connection, test_ollama_connection, set_test_mode, is_test_mode
    from datetime import datetime, timedelta

    # ==================== 工具函数 ====================
    
    def create_circle_avatar(image_path, size=60):
        """创建圆形头像"""
        pixmap = QPixmap(image_path)
        if pixmap.isNull():
            pixmap = QPixmap(size, size)
            pixmap.fill(QColor("#E3F2FD"))
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.Antialiasing)
            painter.setPen(QPen(QColor("#1976D2"), 2))
            painter.setBrush(QBrush(QColor("#BBDEFB")))
            painter.drawEllipse(2, 2, size - 4, size - 4)
            painter.setPen(QColor("#1976D2"))
            painter.setFont(QFont("Microsoft YaHei", size // 3))
            painter.drawText(pixmap.rect(), Qt.AlignCenter, "👤")
            painter.end()
        
        pixmap = pixmap.scaled(size, size, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
        
        rounded = QPixmap(size, size)
        rounded.fill(Qt.transparent)
        
        painter = QPainter(rounded)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QPen(QColor("#4CAF50"), 3))
        painter.setBrush(QBrush(pixmap))
        painter.drawEllipse(2, 2, size - 4, size - 4)
        painter.end()
        
        return rounded

    # ==================== 自定义组件 ====================
    
    class StatCard(SimpleCardWidget):
        """统计卡片组件 - 紧凑型"""
        def __init__(self, title, value, icon=None, parent=None):
            super().__init__(parent)
            self.setMinimumSize(120, 70)
            self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(10)
            shadow.setXOffset(0)
            shadow.setYOffset(2)
            shadow.setColor(QColor(0, 0, 0, 20))
            self.setGraphicsEffect(shadow)
            
            layout = QHBoxLayout(self)
            layout.setSpacing(8)
            layout.setContentsMargins(10, 8, 10, 8)
            
            if icon:
                iconLabel = QLabel(self)
                iconLabel.setFixedSize(28, 28)
                iconLabel.setAlignment(Qt.AlignCenter)
                iconLabel.setStyleSheet("""
                    background-color: #E3F2FD;
                    border-radius: 14px;
                    font-size: 13px;
                """)
                iconLabel.setText(icon)
                layout.addWidget(iconLabel)
            
            textLayout = QVBoxLayout()
            textLayout.setSpacing(2)
            
            self.titleLabel = BodyLabel(title, self)
            self.titleLabel.setStyleSheet("color: #888888; font-size: 9px;")
            
            self.valueLabel = TitleLabel(value, self)
            self.valueLabel.setStyleSheet("color: #1a1a1a; font-size: 15px; font-weight: bold;")
            
            textLayout.addWidget(self.titleLabel)
            textLayout.addWidget(self.valueLabel)
            layout.addLayout(textLayout)
        
        def updateValue(self, value):
            self.valueLabel.setText(value)

    class SectionTitle(QWidget):
        """章节标题"""
        def __init__(self, title, parent=None):
            super().__init__(parent)
            layout = QHBoxLayout(self)
            layout.setContentsMargins(0, 5, 0, 3)
            
            self.label = StrongBodyLabel(title, self)
            self.label.setStyleSheet("color: #333333; font-size: 11px; font-weight: bold;")
            layout.addWidget(self.label)
            layout.addStretch()

    # ==================== 今日工作页面 ====================
    
    class TodayWorkPage(QWidget):
        """今日工作页面"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setObjectName("todayWorkPage")
            self.cards = []  # 存储卡片用于动画
            
            # 主滚动区域
            scrollLayout = QVBoxLayout(self)
            scrollLayout.setContentsMargins(0, 0, 0, 0)
            scrollLayout.setSpacing(0)
            
            scrollArea = QScrollArea()
            scrollArea.setWidgetResizable(True)
            scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scrollArea.setStyleSheet("QScrollArea { border: none; background-color: #F5F5F5; }")
            
            contentWidget = QWidget()
            contentWidget.setStyleSheet("background-color: #F5F5F5; border: none;")
            layout = QVBoxLayout(contentWidget)
            layout.setContentsMargins(16, 12, 16, 12)
            layout.setSpacing(10)
            
            # ========== 头部信息区域 ==========
            headerCard = QFrame()
            headerCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 10px;
                    border: none;
                }
            """)
            self.cards.append(headerCard)
            
            headerLayout = QHBoxLayout(headerCard)
            headerLayout.setContentsMargins(16, 12, 16, 12)
            headerLayout.setSpacing(12)
            
            # 圆形头像
            avatarPath = r"C:\Users\20057\Desktop\frog.jpg"
            avatarPixmap = create_circle_avatar(avatarPath, 45)
            avatarLabel = QLabel()
            avatarLabel.setPixmap(avatarPixmap)
            avatarLabel.setFixedSize(45, 45)
            avatarLabel.setStyleSheet("border: none;")
            headerLayout.addWidget(avatarLabel)
            
            # 右侧文字信息
            infoLayout = QVBoxLayout()
            infoLayout.setSpacing(4)
            
            mainTitle = QLabel("告别加班写周报")
            mainTitle.setStyleSheet("font-size: 16px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            infoLayout.addWidget(mainTitle)
            
            subTitle = QLabel("一周工作内容自动汇总，AI帮你梳理亮点，周五准时下班。")
            subTitle.setStyleSheet("font-size: 10px; color: #888888; border: none; background: transparent;")
            subTitle.setWordWrap(True)
            infoLayout.addWidget(subTitle)
            
            # 标签行
            tagsLayout = QHBoxLayout()
            tagsLayout.setSpacing(8)
            
            tags = [
                ("🔒", "截图分析后即刻销毁"),
                ("💾", "数据仅存本地，不上传云端"),
                ("👤", "你的工作内容只属于你")
            ]
            
            for icon, text in tags:
                tagLabel = QLabel(f"{icon} {text}")
                tagLabel.setStyleSheet("""
                    QLabel {
                        background-color: #E8F5E9;
                        color: #2E7D32;
                        padding: 3px 8px;
                        border-radius: 8px;
                        font-size: 9px;
                        border: none;
                    }
                """)
                tagsLayout.addWidget(tagLabel)
            
            tagsLayout.addStretch()
            infoLayout.addLayout(tagsLayout)
            
            headerLayout.addLayout(infoLayout)
            layout.addWidget(headerCard)
            
            # ========== 工作概览卡片 ==========
            overviewCard = QFrame()
            overviewCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 10px;
                    border: none;
                }
            """)
            self.cards.append(overviewCard)
            
            overviewLayout = QVBoxLayout(overviewCard)
            overviewLayout.setContentsMargins(16, 12, 16, 12)
            overviewLayout.setSpacing(10)
            
            # 标题
            overviewTitle = QLabel("工作概览")
            overviewTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            overviewLayout.addWidget(overviewTitle)
            
            # 描述文本
            self.overviewDesc = QLabel("加载中...")
            self.overviewDesc.setStyleSheet("font-size: 10px; color: #666666; line-height: 1.6; border: none; background: transparent;")
            self.overviewDesc.setWordWrap(True)
            overviewLayout.addWidget(self.overviewDesc)
            
            # 分隔线
            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setFixedHeight(1)
            separator.setStyleSheet("background-color: #E0E0E0; border: none;")
            overviewLayout.addWidget(separator)
            
            # 三个统计项
            statsLayout = QHBoxLayout()
            statsLayout.setSpacing(16)
            
            self.recordCountLabel = QLabel("0")
            self.recordCountLabel.setStyleSheet("font-size: 22px; font-weight: bold; color: #4CAF50; border: none; background: transparent;")
            self.recordCountLabel.setAlignment(Qt.AlignCenter)
            
            self.durationLabel = QLabel("0h")
            self.durationLabel.setStyleSheet("font-size: 22px; font-weight: bold; color: #2196F3; border: none; background: transparent;")
            self.durationLabel.setAlignment(Qt.AlignCenter)
            
            self.mainWorkLabel = QLabel("暂无")
            self.mainWorkLabel.setStyleSheet("font-size: 16px; font-weight: bold; color: #FF9800; border: none; background: transparent;")
            self.mainWorkLabel.setAlignment(Qt.AlignCenter)
            
            # 统计项布局
            for label, sub_text in [(self.recordCountLabel, "记录条数"), 
                                     (self.durationLabel, "专注时长"),
                                     (self.mainWorkLabel, "主要工作")]:
                statWidget = QWidget()
                statWidget.setStyleSheet("border: none; background: transparent;")
                statLayout = QVBoxLayout(statWidget)
                statLayout.setSpacing(3)
                statLayout.addWidget(label, 0, Qt.AlignCenter)
                
                subLabel = QLabel(sub_text)
                subLabel.setStyleSheet("font-size: 9px; color: #999999; border: none; background: transparent;")
                subLabel.setAlignment(Qt.AlignCenter)
                statLayout.addWidget(subLabel, 0, Qt.AlignCenter)
                
                statsLayout.addWidget(statWidget)
            
            overviewLayout.addLayout(statsLayout)
            layout.addWidget(overviewCard)
            
            # ========== 时段记录卡片 ==========
            timeCard = QFrame()
            timeCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 12px;
                    border: none;
                }
            """)
            self.cards.append(timeCard)
            
            timeLayout = QVBoxLayout(timeCard)
            timeLayout.setContentsMargins(20, 20, 20, 20)
            timeLayout.setSpacing(12)
            
            # 标题栏
            timeHeaderLayout = QHBoxLayout()
            timeTitle = QLabel("时段记录")
            timeTitle.setStyleSheet("font-size: 16px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            timeHeaderLayout.addWidget(timeTitle)
            timeHeaderLayout.addStretch()
            
            # 图例
            legendLayout = QHBoxLayout()
            legendLayout.setSpacing(5)
            legendLabel = QLabel("少")
            legendLabel.setStyleSheet("font-size: 10px; color: #999999; border: none; background: transparent;")
            legendLayout.addWidget(legendLabel)
            
            for intensity in range(5):
                block = QLabel()
                block.setFixedSize(12, 12)
                block.setStyleSheet(f"background-color: rgba(76, 175, 80, {50 + intensity * 50}); border-radius: 2px; border: none;")
                legendLayout.addWidget(block)
            
            legendLabel2 = QLabel("多")
            legendLabel2.setStyleSheet("font-size: 10px; color: #999999; border: none; background: transparent;")
            legendLayout.addWidget(legendLabel2)
            
            timeHeaderLayout.addLayout(legendLayout)
            timeLayout.addLayout(timeHeaderLayout)
            
            # 热力图容器
            heatContainer = QWidget()
            heatContainer.setStyleSheet("background: transparent; border: none;")
            heatContainerLayout = QVBoxLayout(heatContainer)
            heatContainerLayout.setContentsMargins(0, 0, 0, 0)
            heatContainerLayout.setSpacing(4)
            
            # 热力图格子
            heatGrid = QGridLayout()
            heatGrid.setSpacing(4)
            heatGrid.setContentsMargins(0, 0, 0, 0)
            
            self.hourBlocks = []
            for h in range(24):
                block = QLabel("0")
                block.setMinimumSize(30, 28)
                block.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                block.setAlignment(Qt.AlignCenter)
                block.setStyleSheet("""
                    background-color: #E8F5E9;
                    border-radius: 6px;
                    font-size: 10px;
                    color: #666666;
                    border: none;
                """)
                self.hourBlocks.append(block)
                heatGrid.addWidget(block, 0, h)
            
            heatContainerLayout.addLayout(heatGrid)
            
            # 时间标签（使用相同的网格布局，居中对齐）
            timeLabelsGrid = QGridLayout()
            timeLabelsGrid.setSpacing(4)
            timeLabelsGrid.setContentsMargins(0, 0, 0, 0)
            
            for h in range(24):
                if h % 3 == 0:
                    label = QLabel(f"{h}:00")
                    label.setStyleSheet("font-size: 9px; color: #999999; border: none; background: transparent;")
                    label.setAlignment(Qt.AlignCenter)  # 居中对齐
                    timeLabelsGrid.addWidget(label, 0, h)
                else:
                    # 空占位符
                    spacer = QWidget()
                    spacer.setStyleSheet("background: transparent; border: none;")
                    timeLabelsGrid.addWidget(spacer, 0, h)
            
            heatContainerLayout.addLayout(timeLabelsGrid)
            timeLayout.addWidget(heatContainer)
            layout.addWidget(timeCard)
            
            # ========== 显示器信息卡片 ==========
            monitorCard = QFrame()
            monitorCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 12px;
                    border: none;
                }
            """)
            self.cards.append(monitorCard)
            
            monitorLayout = QVBoxLayout(monitorCard)
            monitorLayout.setContentsMargins(20, 20, 20, 20)
            monitorLayout.setSpacing(12)
            
            # 标题栏
            monitorHeaderLayout = QHBoxLayout()
            
            monitorIcon = QLabel("🖥️")
            monitorIcon.setStyleSheet("font-size: 16px; border: none; background: transparent;")
            monitorHeaderLayout.addWidget(monitorIcon)
            
            monitorTitle = QLabel("已连接显示器")
            monitorTitle.setStyleSheet("font-size: 16px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            monitorHeaderLayout.addWidget(monitorTitle)
            monitorHeaderLayout.addStretch()
            
            self.monitorCountLabel = QLabel("0台")
            self.monitorCountLabel.setStyleSheet("font-size: 12px; color: #999999; border: none; background: transparent;")
            monitorHeaderLayout.addWidget(self.monitorCountLabel)
            
            monitorLayout.addLayout(monitorHeaderLayout)
            
            # 显示器列表
            self.monitorListLayout = QVBoxLayout()
            self.monitorListLayout.setSpacing(8)
            monitorLayout.addLayout(self.monitorListLayout)
            
            layout.addWidget(monitorCard)
            
            # ========== 今日 Token 用量卡片 ==========
            tokenCard = QFrame()
            tokenCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 12px;
                    border: none;
                }
            """)
            self.cards.append(tokenCard)
            
            tokenCardLayout = QVBoxLayout(tokenCard)
            tokenCardLayout.setContentsMargins(20, 20, 20, 20)
            tokenCardLayout.setSpacing(12)
            
            # 标题栏
            tokenHeaderLayout = QHBoxLayout()
            tokenIcon = QLabel("🪙")
            tokenIcon.setStyleSheet("font-size: 16px; border: none; background: transparent;")
            tokenHeaderLayout.addWidget(tokenIcon)
            tokenTitle = QLabel("今日 Token 用量")
            tokenTitle.setStyleSheet("font-size: 16px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            tokenHeaderLayout.addWidget(tokenTitle)
            tokenHeaderLayout.addStretch()
            self.todayTokenTotalLabel = QLabel("0")
            self.todayTokenTotalLabel.setStyleSheet("font-size: 14px; font-weight: bold; color: #16A34A; border: none; background: transparent;")
            tokenHeaderLayout.addWidget(self.todayTokenTotalLabel)
            tokenCardLayout.addLayout(tokenHeaderLayout)
            
            # 分项：报告输出 / 活动分析
            tokenBreakdownLayout = QHBoxLayout()
            tokenBreakdownLayout.setSpacing(12)
            
            # 报告输出 token
            reportTokenWidget = QFrame()
            reportTokenWidget.setStyleSheet("QFrame { background-color: #F9FAFB; border-radius: 8px; border: 1px solid #F0F0F0; }")
            reportTokenLayout = QVBoxLayout(reportTokenWidget)
            reportTokenLayout.setContentsMargins(14, 12, 14, 12)
            reportTokenLayout.setSpacing(4)
            self.reportTokenLabel = QLabel("0")
            self.reportTokenLabel.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            reportTokenLayout.addWidget(self.reportTokenLabel)
            reportTokenSub = QLabel("报告生成输出")
            reportTokenSub.setStyleSheet("font-size: 11px; color: #999999; border: none; background: transparent;")
            reportTokenLayout.addWidget(reportTokenSub)
            tokenBreakdownLayout.addWidget(reportTokenWidget, 1)
            
            # 活动分析 token
            analysisTokenWidget = QFrame()
            analysisTokenWidget.setStyleSheet("QFrame { background-color: #F9FAFB; border-radius: 8px; border: 1px solid #F0F0F0; }")
            analysisTokenLayout = QVBoxLayout(analysisTokenWidget)
            analysisTokenLayout.setContentsMargins(14, 12, 14, 12)
            analysisTokenLayout.setSpacing(4)
            self.analysisTokenLabel = QLabel("0")
            self.analysisTokenLabel.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            analysisTokenLayout.addWidget(self.analysisTokenLabel)
            analysisTokenSub = QLabel("活动分析")
            analysisTokenSub.setStyleSheet("font-size: 11px; color: #999999; border: none; background: transparent;")
            analysisTokenLayout.addWidget(analysisTokenSub)
            tokenBreakdownLayout.addWidget(analysisTokenWidget, 1)
            
            tokenCardLayout.addLayout(tokenBreakdownLayout)
            
            layout.addWidget(tokenCard)
            
            # 添加弹性空间
            layout.addStretch()
            
            scrollArea.setWidget(contentWidget)
            scrollLayout.addWidget(scrollArea)
            
            # 启动入场动画
            QTimer.singleShot(100, self.playEntryAnimations)
            
            # 加载数据
            self.updateData()
        
        def playEntryAnimations(self):
            """播放卡片入场动画"""
            for i, card in enumerate(self.cards):
                # 设置初始状态：透明且向下偏移
                card.setStyleSheet(card.styleSheet() + "opacity: 0;")
                card.setGraphicsEffect(None)  # 移除旧效果
                
                # 创建淡入动画
                opacityEffect = QGraphicsOpacityEffect(card)
                opacityEffect.setOpacity(0)
                card.setGraphicsEffect(opacityEffect)
                
                anim = QPropertyAnimation(opacityEffect, b"opacity")
                anim.setDuration(400)
                anim.setStartValue(0.0)
                anim.setEndValue(1.0)
                anim.setEasingCurve(QEasingCurve.OutCubic)
                
                # 延迟启动
                QTimer.singleShot(i * 150, anim.start)
                
                # 保存动画引用
                if not hasattr(self, '_animations'):
                    self._animations = []
                self._animations.append(opacityEffect)
                self._animations.append(anim)
        
        def updateData(self):
            """更新页面数据"""
            stats = get_today_stats()
            
            # 更新工作概览
            time_range = stats['time_range']
            record_count = stats['record_count']
            duration = stats['duration_hours']
            
            if record_count > 0:
                # 计算下午时段产出
                hour_data = stats['hour_data']
                afternoon_count = sum(hour_data[12:18])
                overview_text = f"今天工作节奏紧凑，{time_range}，共记录{record_count}段活动，累计专注约{duration:.1f}小时。"
                if afternoon_count > record_count * 0.4:
                    overview_text = f"今天工作节奏紧凑，{time_range}，下午时段产出最多，共记录{record_count}段活动，累计专注约{duration:.1f}小时。"
            else:
                overview_text = "今天暂无工作记录，点击截图分析开始记录你的工作内容。"
            
            self.overviewDesc.setText(overview_text)
            
            # 更新统计数字
            self.recordCountLabel.setText(str(record_count))
            self.durationLabel.setText(f"{duration:.1f}h")
            self.mainWorkLabel.setText(stats['main_work'])
            
            # 更新热力图
            max_count = max(stats['hour_data']) if stats['hour_data'] else 1
            if max_count == 0:
                max_count = 1
            
            for h in range(24):
                count = stats['hour_data'][h]
                intensity = count / max_count if max_count > 0 else 0
                
                if count == 0:
                    # 无记录：浅灰色
                    bg_color = "#F0F0F0"
                    text_color = "#CCCCCC"
                else:
                    # 有记录：从白色渐变到绿色
                    # 白色: rgb(255, 255, 255)
                    # 深绿色: rgb(34, 139, 34)
                    r = int(255 - (255 - 34) * intensity)
                    g = int(255 - (255 - 139) * intensity)
                    b = int(255 - (255 - 34) * intensity)
                    bg_color = f"rgb({r}, {g}, {b})"
                    # 文字颜色：浅色背景用深色字，深色背景用白色字
                    text_color = "#FFFFFF" if intensity > 0.5 else "#333333"
                
                self.hourBlocks[h].setText(str(count))
                self.hourBlocks[h].setStyleSheet(f"""
                    background-color: {bg_color};
                    border-radius: 6px;
                    font-size: 10px;
                    color: {text_color};
                    font-weight: bold;
                    border: none;
                """)
            
            # 更新显示器信息
            monitors = get_monitor_info()
            self.monitorCountLabel.setText(f"{len(monitors)}台")
            
            # 清空旧的显示器列表
            while self.monitorListLayout.count():
                child = self.monitorListLayout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            
            # 添加显示器卡片
            for i, monitor in enumerate(monitors):
                monWidget = QFrame()
                monWidget.setStyleSheet("""
                    QFrame {
                        background-color: #F5F5F5;
                        border-radius: 8px;
                        border: none;
                    }
                """)
                monLayout = QHBoxLayout(monWidget)
                monLayout.setContentsMargins(12, 10, 12, 10)
                monLayout.setSpacing(12)
                
                # 序号
                numLabel = QLabel(str(i + 1))
                numLabel.setFixedSize(28, 28)
                numLabel.setAlignment(Qt.AlignCenter)
                numLabel.setStyleSheet("""
                    background-color: #4CAF50;
                    color: white;
                    border-radius: 14px;
                    font-size: 12px;
                    font-weight: bold;
                    border: none;
                """)
                monLayout.addWidget(numLabel)
                
                # 信息
                infoLayout = QVBoxLayout()
                infoLayout.setSpacing(2)
                
                nameLabel = QLabel(monitor['name'])
                nameLabel.setStyleSheet("font-size: 12px; font-weight: bold; color: #333333; border: none; background: transparent;")
                infoLayout.addWidget(nameLabel)
                
                for text in [monitor['resolution'], monitor['scale'], monitor['refresh_rate']]:
                    label = QLabel(text)
                    label.setStyleSheet("font-size: 10px; color: #888888; border: none; background: transparent;")
                    infoLayout.addWidget(label)
                
                monLayout.addLayout(infoLayout)
                monLayout.addStretch()
                
                self.monitorListLayout.addWidget(monWidget)
            
            # 更新今日 Token 用量（报告输出 + 活动分析）
            from store import get_token_stats, format_token_count
            try:
                ts = get_token_stats()
                report_out = ts.get('today_report_output', 0)
                analysis = ts.get('today_analysis', 0)
                self.reportTokenLabel.setText(format_token_count(report_out))
                self.reportTokenLabel.setToolTip(f"今日报告生成输出 {report_out} Token")
                self.analysisTokenLabel.setText(format_token_count(analysis))
                self.analysisTokenLabel.setToolTip(f"今日活动分析消耗 {analysis} Token")
                self.todayTokenTotalLabel.setText(format_token_count(report_out + analysis))
                self.todayTokenTotalLabel.setToolTip(
                    f"今日合计 {report_out + analysis} Token（报告输出 {report_out} + 活动分析 {analysis}）"
                )
            except Exception as e:
                print(f"[Token] 统计失败: {e}")

    # ==================== 截图识别工作线程 ====================
    
    class ScreenshotWorker(QThread):
        """截图识别工作线程 - 在后台执行识别任务，不阻塞UI"""
        # 定义信号
        finished = pyqtSignal(dict)  # 识别完成信号，传递结果
        error = pyqtSignal(str)      # 错误信号，传递错误信息
        
        def run(self):
            """线程执行函数"""
            try:
                # 在后台线程中执行截图识别
                result = run_and_store()
                # 发送完成信号
                self.finished.emit(result)
            except Exception as e:
                # 发送错误信号
                self.error.emit(str(e))

    # ==================== 截图分析页面 ====================
    
    class ScreenshotPage(QWidget):
        """截图分析页面"""
        def __init__(self, main_window, parent=None):
            super().__init__(parent)
            self.main_window = main_window
            self.setObjectName("screenshotPage")
            
            layout = QVBoxLayout(self)
            layout.setContentsMargins(15, 10, 15, 10)
            layout.setSpacing(12)
            
            title = SubtitleLabel("截图分析", self)
            title.setStyleSheet("font-size: 14px; font-weight: bold;")
            layout.addWidget(title)
            
            infoCard = SimpleCardWidget(self)
            infoLayout = QVBoxLayout(infoCard)
            infoLayout.setContentsMargins(15, 12, 15, 12)
            infoLayout.setSpacing(6)
            
            infoText = BodyLabel(
                "点击下方按钮，系统将自动截取当前屏幕画面，"
                "并通过AI识别当前正在进行的工作类型和内容。",
                infoCard
            )
            infoText.setWordWrap(True)
            infoText.setStyleSheet("color: #666666; font-size: 10px; line-height: 1.4;")
            infoLayout.addWidget(infoText)
            
            layout.addWidget(infoCard)
            
            btnCard = SimpleCardWidget(self)
            btnLayout = QVBoxLayout(btnCard)
            btnLayout.setContentsMargins(20, 15, 20, 15)
            btnLayout.setSpacing(8)
            btnLayout.setAlignment(Qt.AlignCenter)
            
            iconLabel = QLabel("📷", self)
            iconLabel.setAlignment(Qt.AlignCenter)
            iconLabel.setStyleSheet("font-size: 28px;")
            btnLayout.addWidget(iconLabel, 0, Qt.AlignCenter)
            
            self.captureBtn = PrimaryPushButton("开始截图分析", self)
            self.captureBtn.setFixedSize(160, 36)
            self.captureBtn.setStyleSheet("""
                PrimaryPushButton {
                    font-size: 12px;
                    font-weight: bold;
                    border-radius: 6px;
                }
            """)
            self.captureBtn.clicked.connect(self.startCapture)
            btnLayout.addWidget(self.captureBtn, 0, Qt.AlignCenter)
            
            self.statusLabel = CaptionLabel("等待操作...", btnCard)
            self.statusLabel.setAlignment(Qt.AlignCenter)
            self.statusLabel.setStyleSheet("color: #999999; font-size: 9px;")
            btnLayout.addWidget(self.statusLabel, 0, Qt.AlignCenter)
            
            layout.addWidget(btnCard)
            
            resultTitle = SectionTitle("分析结果", self)
            layout.addWidget(resultTitle)
            
            resultCard = SimpleCardWidget(self)
            resultLayout = QVBoxLayout(resultCard)
            resultLayout.setContentsMargins(12, 10, 12, 10)
            resultLayout.setSpacing(8)
            
            typeLayout = QHBoxLayout()
            typeLabel = BodyLabel("工作类型:", resultCard)
            typeLabel.setStyleSheet("color: #666666; font-weight: bold; font-size: 10px;")
            self.typeValue = BodyLabel("--", resultCard)
            self.typeValue.setStyleSheet("color: #0078d4; font-size: 12px; font-weight: bold;")
            typeLayout.addWidget(typeLabel)
            typeLayout.addWidget(self.typeValue)
            typeLayout.addStretch()
            resultLayout.addLayout(typeLayout)
            
            separator = QFrame(resultCard)
            separator.setFrameShape(QFrame.HLine)
            separator.setStyleSheet("background-color: #E0E0E0;")
            resultLayout.addWidget(separator)
            
            descLabel = BodyLabel("工作描述:", resultCard)
            descLabel.setStyleSheet("color: #666666; font-weight: bold; font-size: 10px;")
            resultLayout.addWidget(descLabel)
            
            self.descValue = BodyLabel("暂无分析结果", resultCard)
            self.descValue.setWordWrap(True)
            self.descValue.setStyleSheet("color: #333333; font-size: 10px; line-height: 1.4;")
            resultLayout.addWidget(self.descValue)
            
            layout.addWidget(resultCard, 1)
        
        def startCapture(self):
            """开始截图分析 - 使用多线程"""
            # 禁用按钮，防止重复点击
            self.captureBtn.setEnabled(False)
            self.captureBtn.setText("分析中...")
            self.statusLabel.setText("正在截图并分析，请稍候...")
            self.statusLabel.setStyleSheet("color: #FF9800; font-size: 9px;")
            
            # 创建并启动工作线程
            self.worker = ScreenshotWorker()
            self.worker.finished.connect(self.onCaptureSuccess)
            self.worker.error.connect(self.onCaptureError)
            self.worker.start()
        
        def onCaptureSuccess(self, result):
            """识别成功的回调函数"""
            self.typeValue.setText(result['type'])
            self.descValue.setText(result['description'])
            self.statusLabel.setText("分析完成！")
            self.statusLabel.setStyleSheet("color: #4CAF50; font-size: 9px;")
            
            InfoBar.success(
                title="分析完成",
                content=f"已识别为: {result['type']}",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
            
            # 更新其他页面数据
            self.main_window.todayPage.updateData()
            
            # 恢复按钮状态
            self.captureBtn.setEnabled(True)
            self.captureBtn.setText("开始截图分析")
        
        def onCaptureError(self, error_msg):
            """识别失败的回调函数"""
            self.statusLabel.setText(f"分析失败: {error_msg}")
            self.statusLabel.setStyleSheet("color: #F44336; font-size: 9px;")
            
            InfoBar.error(
                title="分析失败",
                content=error_msg,
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=5000,
                parent=self
            )
            
            # 恢复按钮状态
            self.captureBtn.setEnabled(True)
            self.captureBtn.setText("开始截图分析")

    # ==================== 工作记录页面 ====================
    
    class RecordsPage(QWidget):
        """工作记录页面 - Fluent Design 风格"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setObjectName("recordsPage")
            
            # 主布局
            mainLayout = QVBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            mainLayout.setSpacing(0)
            
            # 滚动区域
            scrollArea = QScrollArea()
            scrollArea.setWidgetResizable(True)
            scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scrollArea.setStyleSheet("QScrollArea { border: none; background-color: #F5F5F5; }")
            
            contentWidget = QWidget()
            contentWidget.setStyleSheet("background-color: #F5F5F5; border: none;")
            layout = QVBoxLayout(contentWidget)
            layout.setContentsMargins(20, 15, 20, 15)
            layout.setSpacing(15)
            
            # 页面标题
            headerCard = QFrame()
            headerCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            headerLayout = QHBoxLayout(headerCard)
            headerLayout.setContentsMargins(20, 15, 20, 15)
            
            title = QLabel("📋 工作记录")
            title.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            headerLayout.addWidget(title)
            headerLayout.addStretch()
            
            refreshBtn = QPushButton("🔄 刷新")
            refreshBtn.setCursor(Qt.PointingHandCursor)
            refreshBtn.setStyleSheet("""
                QPushButton {
                    background-color: #E3F2FD;
                    color: #1976D2;
                    padding: 8px 16px;
                    border-radius: 8px;
                    font-size: 12px;
                    font-weight: bold;
                    border: none;
                }
                QPushButton:hover { background-color: #BBDEFB; }
                QPushButton:pressed { background-color: #90CAF9; }
            """)
            refreshBtn.clicked.connect(self.updateData)
            headerLayout.addWidget(refreshBtn)
            
            layout.addWidget(headerCard)
            
            # 统计卡片
            statsCard = QFrame()
            statsCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            statsLayout = QHBoxLayout(statsCard)
            statsLayout.setContentsMargins(20, 12, 20, 12)
            statsLayout.setSpacing(30)
            
            self.countLabel = QLabel("📊 共 0 条记录")
            self.countLabel.setStyleSheet("font-size: 14px; color: #333333; font-weight: bold; border: none; background: transparent;")
            statsLayout.addWidget(self.countLabel)
            statsLayout.addStretch()
            
            layout.addWidget(statsCard)
            
            # 表格卡片
            tableCard = QFrame()
            tableCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            tableLayout = QVBoxLayout(tableCard)
            tableLayout.setContentsMargins(15, 15, 15, 15)
            
            # 使用 TableWidget (Fluent Design)
            self.recordsTable = TableWidget(self)
            self.recordsTable.setBorderRadius(8)
            self.recordsTable.setBorderVisible(True)
            self.recordsTable.setColumnCount(5)
            self.recordsTable.setHorizontalHeaderLabels(["序号", "时间", "类型", "描述", "时长"])
            self.recordsTable.horizontalHeader().setStretchLastSection(True)
            self.recordsTable.setColumnWidth(0, 60)
            self.recordsTable.setColumnWidth(1, 120)
            self.recordsTable.setColumnWidth(2, 80)
            self.recordsTable.setColumnWidth(3, 450)
            self.recordsTable.setColumnWidth(4, 80)
            self.recordsTable.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.recordsTable.setSelectionBehavior(QTableWidget.SelectRows)
            self.recordsTable.setAlternatingRowColors(True)
            
            tableLayout.addWidget(self.recordsTable)
            layout.addWidget(tableCard, 1)
            
            scrollArea.setWidget(contentWidget)
            mainLayout.addWidget(scrollArea)
        
        def updateData(self):
            """更新记录数据（按时间倒序排列）"""
            records = get_daily_records()
            
            # 按时间倒序排列（最近的在前面）
            records.sort(key=lambda r: r.get('时间', ''), reverse=True)
            
            self.recordsTable.setRowCount(len(records))
            self.countLabel.setText(f"📊 共 {len(records)} 条记录")
            
            for row, record in enumerate(records):
                # 序号
                idItem = QTableWidgetItem(str(row + 1))
                idItem.setTextAlignment(Qt.AlignCenter)
                self.recordsTable.setItem(row, 0, idItem)
                
                # 时间
                timeItem = QTableWidgetItem(record['时间'])
                timeItem.setTextAlignment(Qt.AlignCenter)
                self.recordsTable.setItem(row, 1, timeItem)
                
                # 类型
                typeItem = QTableWidgetItem(record['工作类型'])
                typeItem.setTextAlignment(Qt.AlignCenter)
                self.recordsTable.setItem(row, 2, typeItem)
                
                # 描述
                descItem = QTableWidgetItem(record['工作描述'])
                self.recordsTable.setItem(row, 3, descItem)
                
                # 时长
                duration = record.get('持续时长(分钟)', '0')
                try:
                    duration_min = float(duration)
                    duration_text = f"{duration_min:.0f}分钟" if duration_min > 0 else "-"
                except:
                    duration_text = "-"
                durationItem = QTableWidgetItem(duration_text)
                durationItem.setTextAlignment(Qt.AlignCenter)
                self.recordsTable.setItem(row, 4, durationItem)

    # ==================== 工作时间线页面 ====================
    
    # 工作类型颜色映射（固定颜色）
    TYPE_COLORS = {
        "开发": "#4CAF50",    # 绿色
        "沟通": "#2196F3",    # 蓝色
        "生活": "#FF9800",    # 橙色
        "学习": "#9C27B0",    # 紫色
        "设计": "#E91E63",    # 粉色
        "管理": "#00BCD4",    # 青色
        "文档": "#795548",    # 棕色
        "娱乐": "#F44336",    # 红色
        "产品": "#FF5722",    # 深橙色
        "会议": "#3F51B5",    # 靛蓝色
        "运维": "#009688",    # 青绿色
        "测试": "#FFC107",    # 琥珀色
        "数据分析": "#673AB7", # 深紫色
        "其他": "#607D8B",    # 灰蓝色
    }
    
    class PieChartWidget(QWidget):
        """饼状图组件"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.data = []  # [(name, value, color), ...]
            self.show_percentage = True  # True=显示占比, False=显示时长
            self.setMinimumSize(200, 200)
        
        def setData(self, data):
            """设置数据: [(name, value, color), ...]"""
            self.data = data
            self.update()  # 触发重绘
        
        def setShowPercentage(self, show_percentage):
            """设置是否显示占比"""
            self.show_percentage = show_percentage
            self.update()  # 触发重绘
        
        def paintEvent(self, event):
            """绘制饼状图"""
            if not self.data:
                return
            
            painter = QPainter(self)
            painter.setRenderHint(QPainter.Antialiasing)
            
            # 计算可用区域
            width = self.width()
            height = self.height()
            
            # 饼图区域（左侧）
            pie_size = min(width * 0.6, height * 0.9)
            pie_x = 20
            pie_y = (height - pie_size) / 2
            
            # 计算总值
            total = sum(item[1] for item in self.data)
            if total == 0:
                painter.end()
                return
            
            # 绘制饼图
            start_angle = 0
            for name, value, color in self.data:
                if value <= 0:
                    continue
                
                # 计算角度
                angle = int(360 * 16 * value / total)  # Qt使用1/16度为单位
                
                # 设置画刷
                painter.setBrush(QColor(color))
                painter.setPen(QPen(QColor("#FFFFFF"), 2))
                
                # 绘制扇形
                painter.drawPie(int(pie_x), int(pie_y), int(pie_size), int(pie_size), 
                               start_angle, angle)
                
                start_angle += angle
            
            # 绘制图例（右侧）
            legend_x = pie_x + pie_size + 20
            legend_y = pie_y + 10
            legend_spacing = 22
            
            for i, (name, value, color) in enumerate(self.data):
                if value <= 0:
                    continue
                
                y = legend_y + i * legend_spacing
                percentage = (value / total * 100) if total > 0 else 0
                hours = value / 60  # 转换为小时
                
                # 绘制颜色方块
                painter.setBrush(QColor(color))
                painter.setPen(Qt.NoPen)
                painter.drawRoundedRect(int(legend_x), int(y), 12, 12, 2, 2)
                
                # 绘制文字
                painter.setPen(QColor("#333333"))
                font = QFont("Microsoft YaHei", 9)
                painter.setFont(font)
                
                # 根据模式显示时长或占比
                if self.show_percentage:
                    value_text = f"{percentage:.1f}%"
                else:
                    value_text = f"{hours:.1f}h"
                
                painter.drawText(int(legend_x + 18), int(y + 10), 
                               f"{name} {value_text}")
            
            painter.end()
    
    class TimelinePage(QWidget):
        """工作时间线页面 - 替代数据统计页面"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setObjectName("timelinePage")
            self.main_window = parent
            
            # 主布局
            mainLayout = QVBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            mainLayout.setSpacing(0)
            
            # 滚动区域
            scrollArea = QScrollArea()
            scrollArea.setWidgetResizable(True)
            scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scrollArea.setStyleSheet("QScrollArea { border: none; background-color: #F9F9F9; }")
            
            contentWidget = QWidget()
            contentWidget.setStyleSheet("background-color: #F9F9F9; border: none;")
            layout = QVBoxLayout(contentWidget)
            layout.setSpacing(10)
            layout.setContentsMargins(16, 12, 16, 12)
            
            # ========== 1. 顶部筛选栏 ==========
            filterCard = QFrame()
            filterCard.setStyleSheet("QFrame { background-color: white; border-radius: 8px; border: none; }")
            filterLayout = QHBoxLayout(filterCard)
            filterLayout.setContentsMargins(12, 8, 12, 8)
            filterLayout.setSpacing(10)
            
            # 日期选择 - 开始日期（使用 Fluent CalendarPicker）
            self.startDatePicker = CalendarPicker()
            self.startDatePicker.setDate(QDate.currentDate())
            self.startDatePicker.setDateFormat("yyyy/MM/dd")
            self.startDatePicker.setFixedWidth(110)
            self.startDatePicker.dateChanged.connect(self.updateData)
            filterLayout.addWidget(self.startDatePicker)
            
            # "至" 文本
            toLabel = QLabel("至")
            toLabel.setStyleSheet("color: #666666; font-size: 10px; border: none; background: transparent;")
            filterLayout.addWidget(toLabel)
            
            # 日期选择 - 结束日期（使用 Fluent CalendarPicker）
            self.endDatePicker = CalendarPicker()
            self.endDatePicker.setDate(QDate.currentDate())
            self.endDatePicker.setDateFormat("yyyy/MM/dd")
            self.endDatePicker.setFixedWidth(110)
            self.endDatePicker.dateChanged.connect(self.updateData)
            filterLayout.addWidget(self.endDatePicker)
            
            filterLayout.addStretch()
            
            # 搜索框（使用 Fluent SearchLineEdit）
            self.searchInput = SearchLineEdit()
            self.searchInput.setPlaceholderText("搜索活动...")
            self.searchInput.setClearButtonEnabled(True)
            self.searchInput.setFixedWidth(180)
            self.searchInput.textChanged.connect(self.filterTimeline)
            filterLayout.addWidget(self.searchInput)
            
            layout.addWidget(filterCard)
            
            # ========== 2. 核心数据统计区 ==========
            statsCard = QFrame()
            statsCard.setStyleSheet("QFrame { background-color: white; border-radius: 8px; border: none; }")
            statsLayout = QHBoxLayout(statsCard)
            statsLayout.setContentsMargins(16, 10, 16, 10)
            statsLayout.setSpacing(30)
            
            # 记录条数
            recordWidget = QWidget()
            recordWidget.setStyleSheet("border: none; background: transparent;")
            recordLayout = QVBoxLayout(recordWidget)
            recordLayout.setSpacing(3)
            self.recordCountLabel = QLabel("0")
            self.recordCountLabel.setStyleSheet("font-size: 22px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            recordLayout.addWidget(self.recordCountLabel)
            recordSubLabel = QLabel("记录条数")
            recordSubLabel.setStyleSheet("font-size: 10px; color: #999999; border: none; background: transparent;")
            recordLayout.addWidget(recordSubLabel)
            statsLayout.addWidget(recordWidget)
            
            # 专注时长
            durationWidget = QWidget()
            durationWidget.setStyleSheet("border: none; background: transparent;")
            durationLayout = QVBoxLayout(durationWidget)
            durationLayout.setSpacing(3)
            self.durationLabel = QLabel("0h")
            self.durationLabel.setStyleSheet("font-size: 22px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            durationLayout.addWidget(self.durationLabel)
            durationSubLabel = QLabel("专注时长")
            durationSubLabel.setStyleSheet("font-size: 10px; color: #999999; border: none; background: transparent;")
            durationLayout.addWidget(durationSubLabel)
            statsLayout.addWidget(durationWidget)
            
            # 活跃时段
            activeWidget = QWidget()
            activeWidget.setStyleSheet("border: none; background: transparent;")
            activeLayout = QVBoxLayout(activeWidget)
            activeLayout.setSpacing(3)
            self.activeTimeLabel = QLabel("--:-- — --:--")
            self.activeTimeLabel.setStyleSheet("font-size: 16px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            activeLayout.addWidget(self.activeTimeLabel)
            activeSubLabel = QLabel("活跃时段")
            activeSubLabel.setStyleSheet("font-size: 10px; color: #999999; border: none; background: transparent;")
            activeLayout.addWidget(activeSubLabel)
            statsLayout.addWidget(activeWidget)
            
            # Token 消耗（活动分析）
            tokenWidget = QWidget()
            tokenWidget.setStyleSheet("border: none; background: transparent;")
            tokenStatLayout = QVBoxLayout(tokenWidget)
            tokenStatLayout.setSpacing(3)
            self.tokenCountLabel = QLabel("0")
            self.tokenCountLabel.setStyleSheet("font-size: 22px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            tokenStatLayout.addWidget(self.tokenCountLabel)
            self.tokenSubLabel = QLabel("今日Token")
            self.tokenSubLabel.setStyleSheet("font-size: 10px; color: #999999; border: none; background: transparent;")
            tokenStatLayout.addWidget(self.tokenSubLabel)
            statsLayout.addWidget(tokenWidget)
            
            statsLayout.addStretch()
            
            # 显示分类时长分布开关
            self.showDistCheckBox = QCheckBox("显示分类时长分布")
            self.showDistCheckBox.setChecked(True)
            self.showDistCheckBox.setStyleSheet("""
                QCheckBox {
                    font-size: 10px; color: #333333; border: none; background: transparent;
                    spacing: 4px;
                }
                QCheckBox::indicator {
                    width: 14px; height: 14px;
                    border: 2px solid #E0E0E0;
                    border-radius: 3px;
                    background-color: white;
                }
                QCheckBox::indicator:checked {
                    background-color: #4CAF50;
                    border: 2px solid #4CAF50;
                }
            """)
            self.showDistCheckBox.stateChanged.connect(self.toggleDistribution)
            statsLayout.addWidget(self.showDistCheckBox)
            
            layout.addWidget(statsCard)
            
            # ========== 3. 分类时长分布图 ==========
            self.distCard = QFrame()
            self.distCard.setStyleSheet("QFrame { background-color: white; border-radius: 8px; border: none; }")
            distLayout = QVBoxLayout(self.distCard)
            distLayout.setContentsMargins(16, 10, 16, 10)
            distLayout.setSpacing(10)
            
            # 标题栏（含切换按钮）
            distHeaderLayout = QHBoxLayout()
            distTitle = QLabel("📊 分类时长分布")
            distTitle.setStyleSheet("font-size: 12px; font-weight: bold; color: #333333; border: none; background: transparent;")
            distHeaderLayout.addWidget(distTitle)
            distHeaderLayout.addStretch()
            
            # 切换按钮样式
            btnStyle = """
                QPushButton {
                    background-color: #E3F2FD;
                    color: #1976D2;
                    padding: 3px 10px;
                    border-radius: 12px;
                    font-size: 9px;
                    border: none;
                }
                QPushButton:hover { background-color: #BBDEFB; }
            """
            
            # 时长/占比切换按钮
            self.distValueBtn = QPushButton("⏱️ 时长")
            self.distValueBtn.setCursor(Qt.PointingHandCursor)
            self.distValueBtn.setStyleSheet(btnStyle)
            self.distValueBtn.clicked.connect(self.toggleDistValueMode)
            distHeaderLayout.addWidget(self.distValueBtn)
            
            # 饼状图/条形图切换按钮
            self.distModeBtn = QPushButton("🥧 条形图")
            self.distModeBtn.setCursor(Qt.PointingHandCursor)
            self.distModeBtn.setStyleSheet(btnStyle)
            self.distModeBtn.clicked.connect(self.toggleDistMode)
            distHeaderLayout.addWidget(self.distModeBtn)
            
            distLayout.addLayout(distHeaderLayout)
            
            # 当前显示模式（False=条形图, True=饼状图）
            self.is_pie_mode = False
            # 当前数值模式（False=时长, True=占比）
            self.is_percentage_mode = False
            
            # 条形容器
            self.barContainer = QWidget()
            self.barContainer.setStyleSheet("border: none; background: transparent;")
            self.distListLayout = QVBoxLayout(self.barContainer)
            self.distListLayout.setSpacing(6)
            self.distListLayout.setContentsMargins(0, 0, 0, 0)
            distLayout.addWidget(self.barContainer)
            
            # 饼状图容器
            self.pieContainer = QWidget()
            self.pieContainer.setStyleSheet("border: none; background: transparent;")
            self.pieContainer.setVisible(False)
            self.pieLayout = QVBoxLayout(self.pieContainer)
            self.pieLayout.setContentsMargins(0, 0, 0, 0)
            self.pieChart = PieChartWidget()
            self.pieChart.setMinimumHeight(200)
            self.pieLayout.addWidget(self.pieChart)
            distLayout.addWidget(self.pieContainer)
            
            layout.addWidget(self.distCard)
            
            # ========== 4. 活动时间轴列表 ==========
            timelineCard = QFrame()
            timelineCard.setStyleSheet("QFrame { background-color: white; border-radius: 8px; border: none; }")
            timelineLayout = QVBoxLayout(timelineCard)
            timelineLayout.setContentsMargins(16, 10, 16, 10)
            timelineLayout.setSpacing(10)
            
            # 工具栏
            toolbarLayout = QHBoxLayout()
            
            toolbarTitle = QLabel("⏱️ 活动时间线")
            toolbarTitle.setStyleSheet("font-size: 12px; font-weight: bold; color: #333333; border: none; background: transparent;")
            toolbarLayout.addWidget(toolbarTitle)
            
            # 标签筛选下拉框（使用 Fluent ComboBox）
            self.tagFilterCombo = ComboBox()
            self.tagFilterCombo.addItems(["全部标签", "开发", "沟通", "生活", "学习", "设计", "管理", "文档", "娱乐", "产品", "会议", "运维", "测试", "数据分析", "其他"])
            self.tagFilterCombo.setCurrentIndex(0)
            self.tagFilterCombo.setFixedWidth(100)
            self.tagFilterCombo.currentTextChanged.connect(self.filterTimeline)
            toolbarLayout.addWidget(self.tagFilterCombo)
            
            toolbarLayout.addStretch()
            
            # 快速时间筛选按钮（使用 Fluent PillPushButton）
            for text in ["近30分", "近1小时", "近2小时", "今天"]:
                btn = PillPushButton(text)
                btn.setCursor(Qt.PointingHandCursor)
                btn.setCheckable(False)
                btn.clicked.connect(lambda checked, t=text: self.quickFilter(t))
                toolbarLayout.addWidget(btn)
            
            # 复制日志按钮（使用 Fluent TransparentPushButton）
            copyBtn = TransparentPushButton("复制日志")
            copyBtn.setIcon(FluentIcon.COPY)
            copyBtn.setCursor(Qt.PointingHandCursor)
            copyBtn.clicked.connect(self.copyLog)
            toolbarLayout.addWidget(copyBtn)
            
            timelineLayout.addLayout(toolbarLayout)
            
            # 时间轴列表容器
            self.timelineContainer = QWidget()
            self.timelineContainer.setStyleSheet("border: none; background: transparent;")
            self.timelineListLayout = QVBoxLayout(self.timelineContainer)
            self.timelineListLayout.setSpacing(0)
            self.timelineListLayout.setContentsMargins(0, 0, 0, 0)
            timelineLayout.addWidget(self.timelineContainer)
            
            layout.addWidget(timelineCard, 1)  # 时间轴占据更多空间
            
            # 添加弹性空间
            layout.addStretch()
            
            scrollArea.setWidget(contentWidget)
            mainLayout.addWidget(scrollArea)
            
            # 存储所有记录用于筛选
            self.all_records = []
        
        def toggleDistribution(self, state):
            """切换分类时长分布显示"""
            self.distCard.setVisible(state == Qt.Checked)
        
        def toggleDistMode(self):
            """切换条形图/饼状图模式"""
            self.is_pie_mode = not self.is_pie_mode
            
            if self.is_pie_mode:
                self.distModeBtn.setText("📊 饼状图")
                self.barContainer.setVisible(False)
                self.pieContainer.setVisible(True)
            else:
                self.distModeBtn.setText("🥧 条形图")
                self.barContainer.setVisible(True)
                self.pieContainer.setVisible(False)
            
            # 刷新数据
            self.updateDistribution()
        
        def toggleDistValueMode(self):
            """切换时长/占比显示模式"""
            self.is_percentage_mode = not self.is_percentage_mode
            
            if self.is_percentage_mode:
                self.distValueBtn.setText("📈 占比")
            else:
                self.distValueBtn.setText("⏱️ 时长")
            
            # 更新饼状图显示模式
            self.pieChart.setShowPercentage(self.is_percentage_mode)
            
            # 刷新数据
            self.updateDistribution()
        
        def quickFilter(self, timeText):
            """快速时间筛选"""
            now = QDate.currentDate()
            if timeText == "今天":
                self.startDatePicker.setDate(now)
                self.endDatePicker.setDate(now)
            elif timeText == "近30分":
                self.startDatePicker.setDate(now)
                self.endDatePicker.setDate(now)
            elif timeText == "近1小时":
                self.startDatePicker.setDate(now)
                self.endDatePicker.setDate(now)
            elif timeText == "近2小时":
                self.startDatePicker.setDate(now)
                self.endDatePicker.setDate(now)
        
        def filterTimeline(self):
            """筛选时间轴"""
            self.updateData()
        
        def copyLog(self):
            """复制日志到剪贴板"""
            records = self.getFilteredRecords()
            if not records:
                return
            
            log_text = ""
            for record in records:
                time = record.get('时间', '')
                work_type = record.get('工作类型', '')
                description = record.get('工作描述', '')
                log_text += f"[{time}] [{work_type}] {description}\n"
            
            QApplication.clipboard().setText(log_text)
            InfoBar.success(
                title="复制成功",
                content=f"已复制 {len(records)} 条记录到剪贴板",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
        
        def getFilteredRecords(self):
            """获取筛选后的记录"""
            records = self.all_records.copy()
            
            # 按标签筛选
            tag = self.tagFilterCombo.currentText()
            if tag != "全部标签":
                records = [r for r in records if r.get('工作类型') == tag]
            
            # 按搜索关键词筛选
            keyword = self.searchInput.text().strip()
            if keyword:
                records = [r for r in records if keyword.lower() in r.get('工作描述', '').lower()]
            
            return records
        
        def updateData(self):
            """更新页面数据"""
            # 获取日期范围内的记录
            start_date = self.startDatePicker.date.toString("yyyy-MM-dd")
            end_date = self.endDatePicker.date.toString("yyyy-MM-dd")
            
            # 验证日期范围
            if self.startDatePicker.date > self.endDatePicker.date:
                InfoBar.warning(
                    title="日期范围错误",
                    content="开始日期不能晚于结束日期",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                return
            
            # 读取所有记录
            all_records = read_records()
            
            # 筛选日期范围内的记录
            self.all_records = [r for r in all_records if start_date <= r.get('日期', '') <= end_date]
            
            # 更新统计数据
            self.updateStats()
            
            # 更新分类分布
            self.updateDistribution()
            
            # 更新时间轴
            self.updateTimeline()
        
        def updateStats(self):
            """更新统计数据"""
            records = self.all_records
            
            # 记录条数
            self.recordCountLabel.setText(str(len(records)))
            
            # 专注时长（总分钟数转小时）
            total_minutes = 0
            for r in records:
                try:
                    total_minutes += float(r.get('持续时长(分钟)', '0'))
                except:
                    pass
            hours = total_minutes / 60
            self.durationLabel.setText(f"{hours:.1f}h")
            
            # 活跃时段（最早和最晚时间）
            if records:
                times = [r.get('时间', '23:59:59') for r in records]
                earliest = min(times)
                latest = max(times)
                self.activeTimeLabel.setText(f"{earliest[:5]} — {latest[:5]}")
            else:
                self.activeTimeLabel.setText("--:-- — --:--")
            
            # Token 消耗（活动分析）：今日总数 + 全部总数
            from store import format_token_count
            all_recs = read_records()
            today_str = get_today()
            today_tokens = 0
            all_tokens = 0
            for r in all_recs:
                try:
                    tk = int(float(r.get('消耗token数', 0) or 0))
                except (TypeError, ValueError):
                    tk = 0
                all_tokens += tk
                if r.get('日期', '') == today_str:
                    today_tokens += tk
            self.tokenCountLabel.setText(format_token_count(today_tokens))
            self.tokenSubLabel.setText(f"今日Token · 全部 {format_token_count(all_tokens)}")
        
        def updateDistribution(self):
            """更新分类时长分布"""
            # 统计各类型时长
            type_hours = {}
            total_minutes = 0
            for r in self.all_records:
                work_type = r.get('工作类型', '其他')
                try:
                    minutes = float(r.get('持续时长(分钟)', '0'))
                except:
                    minutes = 0
                type_hours[work_type] = type_hours.get(work_type, 0) + minutes
                total_minutes += minutes
            
            # 按时长排序
            sorted_types = sorted(type_hours.items(), key=lambda x: x[1], reverse=True)
            
            # 更新条形图
            # 清空旧内容
            while self.distListLayout.count():
                child = self.distListLayout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            
            # 创建进度条
            for work_type, minutes in sorted_types:
                hours = minutes / 60
                percentage = (minutes / total_minutes * 100) if total_minutes > 0 else 0
                color = TYPE_COLORS.get(work_type, "#607D8B")
                
                # 单行容器
                rowWidget = QWidget()
                rowWidget.setStyleSheet("border: none; background: transparent;")
                rowLayout = QHBoxLayout(rowWidget)
                rowLayout.setContentsMargins(0, 0, 0, 0)
                rowLayout.setSpacing(12)
                
                # 类型名称
                nameLabel = QLabel(work_type)
                nameLabel.setFixedWidth(50)
                nameLabel.setStyleSheet(f"font-size: 12px; color: {color}; font-weight: bold; border: none; background: transparent;")
                rowLayout.addWidget(nameLabel)
                
                # 进度条背景
                progressBg = QFrame()
                progressBg.setFixedHeight(8)
                progressBg.setStyleSheet("background-color: #EDEDED; border-radius: 4px; border: none;")
                progressBgLayout = QHBoxLayout(progressBg)
                progressBgLayout.setContentsMargins(0, 0, 0, 0)
                progressBgLayout.setSpacing(0)
                
                # 进度条填充（左对齐）
                progressFill = QFrame()
                progressFill.setFixedHeight(8)
                progressFill.setStyleSheet(f"background-color: {color}; border-radius: 4px; border: none;")
                
                # 设置进度条宽度比例
                fillWidth = max(4, int(300 * percentage / 100))
                progressFill.setFixedWidth(fillWidth)
                
                # 添加到布局并左对齐
                progressBgLayout.addWidget(progressFill, 0, Qt.AlignLeft)
                progressBgLayout.addStretch(1)
                
                rowLayout.addWidget(progressBg, 1)
                
                # 时长/占比文本
                if self.is_percentage_mode:
                    valueText = f"{percentage:.1f}%"
                else:
                    valueText = f"{hours:.1f}h"
                timeLabel = QLabel(valueText)
                timeLabel.setFixedWidth(50)
                timeLabel.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                timeLabel.setStyleSheet("font-size: 12px; color: #999999; border: none; background: transparent;")
                rowLayout.addWidget(timeLabel)
                
                self.distListLayout.addWidget(rowWidget)
            
            # 更新饼状图
            pie_data = []
            for work_type, minutes in sorted_types:
                color = TYPE_COLORS.get(work_type, "#607D8B")
                pie_data.append((work_type, minutes, color))
            self.pieChart.setData(pie_data)
        
        def updateTimeline(self):
            """更新时间轴列表"""
            # 清空旧内容
            while self.timelineListLayout.count():
                child = self.timelineListLayout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            
            # 获取筛选后的记录
            records = self.getFilteredRecords()
            
            # 按时间倒序显示（最新的在上面）
            records = list(reversed(records))
            
            for i, record in enumerate(records):
                time = record.get('时间', '')
                date = record.get('日期', '')
                work_type = record.get('工作类型', '其他')
                description = record.get('工作描述', '')
                duration = record.get('持续时长(分钟)', '0')
                color = TYPE_COLORS.get(work_type, "#607D8B")
                
                # 计算结束时间
                try:
                    duration_min = float(duration)
                    start_dt = datetime.strptime(time, '%H:%M:%S')
                    end_dt = start_dt + timedelta(minutes=duration_min)
                    end_time = end_dt.strftime('%H:%M:%S')
                    time_range = f"{time[:5]} — {end_time[:5]}"
                except:
                    time_range = ""
                
                # 格式化日期显示（月-日）
                date_display = ""
                if date:
                    try:
                        date_obj = datetime.strptime(date, '%Y-%m-%d')
                        date_display = date_obj.strftime('%m/%d')
                    except:
                        date_display = ""
                
                # 列表项容器
                itemWidget = QWidget()
                itemWidget.setStyleSheet("border: none; background: transparent;")
                itemLayout = QHBoxLayout(itemWidget)
                itemLayout.setContentsMargins(0, 5, 0, 5)
                itemLayout.setSpacing(12)
                
                # 时间戳（包含日期和时间）
                timeLabel = QLabel(f"{date_display}\n{time[:5]}")
                timeLabel.setFixedWidth(55)
                timeLabel.setAlignment(Qt.AlignRight | Qt.AlignTop)
                timeLabel.setStyleSheet("font-size: 11px; color: #999999; border: none; background: transparent;")
                itemLayout.addWidget(timeLabel)
                
                # 时间轴指示器
                indicatorWidget = QWidget()
                indicatorWidget.setFixedWidth(20)
                indicatorWidget.setStyleSheet("border: none; background: transparent;")
                indicatorLayout = QVBoxLayout(indicatorWidget)
                indicatorLayout.setContentsMargins(0, 4, 0, 0)
                indicatorLayout.setSpacing(0)
                
                # 圆点
                dot = QLabel()
                dot.setFixedSize(10, 10)
                dot.setStyleSheet(f"background-color: {color}; border-radius: 5px; border: none;")
                indicatorLayout.addWidget(dot, 0, Qt.AlignHCenter)
                
                # 连接线（如果不是最后一个）
                if i < len(records) - 1:
                    line = QFrame()
                    line.setFixedWidth(2)
                    line.setStyleSheet("background-color: #E0E0E0; border: none;")
                    indicatorLayout.addWidget(line, 1, Qt.AlignHCenter)
                
                itemLayout.addWidget(indicatorWidget)
                
                # 内容卡片
                card = QFrame()
                card.setStyleSheet("""
                    QFrame {
                        background-color: #FAFAFA;
                        border-radius: 8px;
                        border: 1px solid #F0F0F0;
                    }
                """)
                cardLayout = QVBoxLayout(card)
                cardLayout.setContentsMargins(12, 10, 12, 10)
                cardLayout.setSpacing(8)
                
                # 文本内容
                descLabel = QLabel(description)
                descLabel.setWordWrap(True)
                descLabel.setStyleSheet("font-size: 12px; color: #333333; border: none; background: transparent;")
                cardLayout.addWidget(descLabel)
                
                # 底部标签栏
                tagsLayout = QHBoxLayout()
                tagsLayout.setSpacing(8)
                
                # 类型标签
                typeTag = QLabel(work_type)
                typeTag.setStyleSheet(f"""
                    QLabel {{
                        background-color: {color};
                        color: white;
                        padding: 2px 8px;
                        border-radius: 10px;
                        font-size: 10px;
                        font-weight: bold;
                        border: none;
                    }}
                """)
                tagsLayout.addWidget(typeTag)
                
                # 自动记录标签
                autoTag = QLabel("自动记录")
                autoTag.setStyleSheet("""
                    QLabel {
                        background-color: #E0E0E0;
                        color: #666666;
                        padding: 2px 8px;
                        border-radius: 10px;
                        font-size: 10px;
                        border: none;
                    }
                """)
                tagsLayout.addWidget(autoTag)
                
                # 时间段
                if time_range:
                    timeRangeLabel = QLabel(time_range)
                    timeRangeLabel.setStyleSheet("font-size: 10px; color: #CCCCCC; border: none; background: transparent;")
                    tagsLayout.addWidget(timeRangeLabel)
                
                # 消耗 token（显示在时间右侧）
                try:
                    rec_tokens = int(float(record.get('消耗token数', 0) or 0))
                except (TypeError, ValueError):
                    rec_tokens = 0
                if rec_tokens > 0:
                    from store import format_token_count
                    tokenLabel = QLabel(f"🪙 {format_token_count(rec_tokens)}")
                    tokenLabel.setStyleSheet("font-size: 10px; color: #9CA3AF; border: none; background: transparent;")
                    tokenLabel.setToolTip(f"本次活动分析消耗 {rec_tokens} Token")
                    tagsLayout.addWidget(tokenLabel)
                
                tagsLayout.addStretch()
                cardLayout.addLayout(tagsLayout)
                
                itemLayout.addWidget(card, 1)
                
                self.timelineListLayout.addWidget(itemWidget)
            
            # 如果没有记录
            if not records:
                emptyLabel = QLabel("暂无记录")
                emptyLabel.setAlignment(Qt.AlignCenter)
                emptyLabel.setStyleSheet("font-size: 14px; color: #CCCCCC; padding: 40px; border: none; background: transparent;")
                self.timelineListLayout.addWidget(emptyLabel)

    # ==================== 报告模板数据 ====================
    
    # 从 store 模块加载报告模板
    from store import read_templates, write_templates, add_template, delete_template, export_templates, import_templates
    REPORT_TEMPLATES = read_templates()
    
    # ==================== 报告生成工作线程 ====================
    
    class ReportGenerateWorker(QThread):
        """报告生成工作线程"""
        chunk_received = pyqtSignal(str)  # 接收到一块内容
        generation_finished = pyqtSignal(str)  # 生成完成
        generation_error = pyqtSignal(str)  # 生成出错
        
        def __init__(self, template_prompt, start_date, end_date, report_type):
            super().__init__()
            self.template_prompt = template_prompt
            self.start_date = start_date
            self.end_date = end_date
            self.report_type = report_type
            self.full_content = ""
            self.usage = {"input": 0, "output": 0, "total": 0}  # 本次生成消耗的 token
        
        def run(self):
            try:
                def on_chunk(chunk, is_finished):
                    if chunk:
                        self.full_content += chunk
                        self.chunk_received.emit(chunk)
                    # 完成信号延迟到拿到 token 用量后再发出
                
                from screenshot import generate_report_stream
                content, usage = generate_report_stream(
                    self.template_prompt,
                    self.start_date,
                    self.end_date,
                    self.report_type,
                    callback=on_chunk
                )
                self.full_content = content
                self.usage = usage or self.usage
                self.generation_finished.emit(self.full_content)
            
            except Exception as e:
                self.generation_error.emit(str(e))
    
    # ==================== 生成报告页面 ====================
    
    class TemplateCard(QFrame):
        """模板卡片组件"""
        clicked = pyqtSignal(int)  # 点击信号，传递模板索引
        preview_clicked = pyqtSignal(int)  # 预览按钮点击信号
        delete_clicked = pyqtSignal(int)  # 删除按钮点击信号
        
        def __init__(self, index, name, intro, is_cloud=True, parent=None):
            super().__init__(parent)
            self.index = index
            self.is_selected = False
            self.is_hovered = False
            self.setFixedSize(180, 90)
            self.setCursor(Qt.PointingHandCursor)
            self.setMouseTracking(True)
            
            # 主布局
            layout = QVBoxLayout(self)
            layout.setContentsMargins(10, 8, 10, 8)
            layout.setSpacing(4)
            
            # 标题行
            titleLayout = QHBoxLayout()
            titleLabel = QLabel(name)
            titleLabel.setStyleSheet("font-size: 11px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            titleLayout.addWidget(titleLabel)
            titleLayout.addStretch()
            
            # 选中对勾图标（默认隐藏）
            self.checkIcon = QLabel("✓")
            self.checkIcon.setFixedSize(14, 14)
            self.checkIcon.setAlignment(Qt.AlignCenter)
            self.checkIcon.setStyleSheet("""
                QLabel {
                    background-color: #16A34A;
                    color: white;
                    border-radius: 7px;
                    font-size: 9px;
                    font-weight: bold;
                    border: none;
                }
            """)
            self.checkIcon.setVisible(False)
            titleLayout.addWidget(self.checkIcon)
            layout.addLayout(titleLayout)
            
            # 简介
            introLabel = QLabel(intro)
            introLabel.setWordWrap(True)
            introLabel.setMaximumHeight(25)
            introLabel.setStyleSheet("font-size: 9px; color: #666666; border: none; background: transparent;")
            introLabel.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            layout.addWidget(introLabel)
            
            layout.addStretch()
            
            # 底部标签行
            tagLayout = QHBoxLayout()
            if is_cloud:
                tag = QLabel("云端")
                tag.setStyleSheet("""
                    QLabel {
                        background-color: #E3F7EA;
                        color: #16A34A;
                        padding: 2px 8px;
                        border-radius: 10px;
                        font-size: 10px;
                        font-weight: bold;
                        border: none;
                    }
                """)
                tagLayout.addWidget(tag)
            tagLayout.addStretch()
            
            # 删除按钮（默认隐藏）
            self.deleteBtn = QPushButton("🗑")
            self.deleteBtn.setFixedSize(28, 28)
            self.deleteBtn.setCursor(Qt.PointingHandCursor)
            self.deleteBtn.setStyleSheet("""
                QPushButton {
                    background-color: #FEE2E2;
                    border-radius: 14px;
                    font-size: 14px;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #FECACA;
                }
            """)
            self.deleteBtn.setVisible(False)
            self.deleteBtn.clicked.connect(lambda: self.delete_clicked.emit(self.index))
            tagLayout.addWidget(self.deleteBtn)
            
            # 预览按钮（默认隐藏）
            self.previewBtn = QPushButton("👁")
            self.previewBtn.setFixedSize(28, 28)
            self.previewBtn.setCursor(Qt.PointingHandCursor)
            self.previewBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F3F4F6;
                    border: 1px solid #E5E7EB;
                    border-radius: 14px;
                    font-size: 14px;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #E5E7EB;
                }
            """)
            self.previewBtn.setVisible(False)
            self.previewBtn.clicked.connect(lambda: self.preview_clicked.emit(self.index))
            tagLayout.addWidget(self.previewBtn)
            
            layout.addLayout(tagLayout)
            
            self.updateStyle()
        
        def setSelected(self, selected):
            self.is_selected = selected
            self.checkIcon.setVisible(selected)
            self.updateStyle()
        
        def updateStyle(self):
            if self.is_selected:
                self.setStyleSheet("""
                    TemplateCard {
                        background-color: #F0FBF4;
                        border: 2px solid #16A34A;
                        border-radius: 12px;
                    }
                """)
            else:
                self.setStyleSheet("""
                    TemplateCard {
                        background-color: white;
                        border: 1px solid #ECECEC;
                        border-radius: 12px;
                    }
                    TemplateCard:hover {
                        border: 1px solid #D1D5DB;
                    }
                """)
        
        def enterEvent(self, event):
            self.is_hovered = True
            self.previewBtn.setVisible(True)
            self.deleteBtn.setVisible(True)
            super().enterEvent(event)
        
        def leaveEvent(self, event):
            self.is_hovered = False
            self.previewBtn.setVisible(False)
            self.deleteBtn.setVisible(False)
            super().leaveEvent(event)
        
        def mousePressEvent(self, event):
            if not self.previewBtn.underMouse() and not self.deleteBtn.underMouse():
                self.clicked.emit(self.index)
            super().mousePressEvent(event)
    
    class TemplatePreviewDialog(QDialog):
        """模板提示词查看弹窗（4.1）"""
        prompt_updated = pyqtSignal(int, str)  # 信号：模板索引，新提示词
        
        def __init__(self, template_index, template_name, template_desc, prompt_text, parent=None):
            super().__init__(parent)
            self.template_index = template_index
            self.template_name = template_name
            self.is_editing = False
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setModal(True)
            
            # 主布局
            mainLayout = QHBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            
            # 半透明遮罩
            overlay = QWidget()
            overlay.setStyleSheet("background-color: rgba(0, 0, 0, 100);")
            overlayLayout = QVBoxLayout(overlay)
            overlayLayout.setAlignment(Qt.AlignCenter)
            
            # 弹窗卡片
            card = QFrame()
            card.setFixedSize(750, 600)
            card.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 16px;
                    border: 1px solid #ECECEC;
                }
            """)
            cardLayout = QVBoxLayout(card)
            cardLayout.setContentsMargins(24, 24, 24, 24)
            cardLayout.setSpacing(16)
            
            # 头部
            headerLayout = QHBoxLayout()
            titleLabel = QLabel(template_name)
            titleLabel.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            headerLayout.addWidget(titleLabel)
            headerLayout.addStretch()
            
            # 修改按钮
            self.editBtn = QPushButton("✏️ 修改")
            self.editBtn.setCursor(Qt.PointingHandCursor)
            self.editBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 8px 16px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            self.editBtn.clicked.connect(self.toggleEdit)
            headerLayout.addWidget(self.editBtn)
            
            closeBtn = QPushButton("✕")
            closeBtn.setFixedSize(32, 32)
            closeBtn.setCursor(Qt.PointingHandCursor)
            closeBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F3F4F6;
                    border: none;
                    border-radius: 16px;
                    font-size: 16px;
                    color: #666666;
                }
                QPushButton:hover {
                    background-color: #E5E7EB;
                }
            """)
            closeBtn.clicked.connect(self.close)
            headerLayout.addWidget(closeBtn)
            cardLayout.addLayout(headerLayout)
            
            # 描述
            descLayout = QHBoxLayout()
            descLabel = QLabel(template_desc)
            descLabel.setStyleSheet("font-size: 13px; color: #666666; border: none; background: transparent;")
            descLabel.setWordWrap(True)
            descLayout.addWidget(descLabel)
            
            tag = QLabel("云端")
            tag.setStyleSheet("""
                QLabel {
                    background-color: #E3F7EA;
                    color: #16A34A;
                    padding: 2px 8px;
                    border-radius: 10px;
                    font-size: 10px;
                    font-weight: bold;
                    border: none;
                }
            """)
            descLayout.addWidget(tag)
            descLayout.addStretch()
            cardLayout.addLayout(descLayout)
            
            # 分隔线
            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setStyleSheet("background-color: #F3F4F6; border: none; height: 1px;")
            cardLayout.addWidget(separator)
            
            # 提示词内容（可编辑的 QTextEdit）
            self.promptEdit = QTextEdit()
            self.promptEdit.setPlainText(prompt_text)
            self.promptEdit.setReadOnly(True)
            self.promptEdit.setStyleSheet("""
                QTextEdit {
                    background-color: #F9FAFB;
                    padding: 16px;
                    border-radius: 8px;
                    font-size: 13px;
                    color: #374151;
                    font-family: Consolas, monospace;
                    border: 1px solid #E5E7EB;
                }
                QTextEdit:focus {
                    border: 1px solid #16A34A;
                }
            """)
            cardLayout.addWidget(self.promptEdit)
            
            # 底部保存按钮（默认隐藏）
            self.saveBtnLayout = QHBoxLayout()
            self.saveBtnLayout.addStretch()
            
            cancelSaveBtn = QPushButton("取消")
            cancelSaveBtn.setCursor(Qt.PointingHandCursor)
            cancelSaveBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 10px 20px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            cancelSaveBtn.clicked.connect(self.cancelEdit)
            self.saveBtnLayout.addWidget(cancelSaveBtn)
            
            saveBtn = QPushButton("保存修改")
            saveBtn.setCursor(Qt.PointingHandCursor)
            saveBtn.setStyleSheet("""
                QPushButton {
                    background-color: #16A34A;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #15803D;
                }
            """)
            saveBtn.clicked.connect(self.saveEdit)
            self.saveBtnLayout.addWidget(saveBtn)
            
            self.saveBtnWidget = QWidget()
            self.saveBtnWidget.setLayout(self.saveBtnLayout)
            self.saveBtnWidget.setVisible(False)
            cardLayout.addWidget(self.saveBtnWidget)
            
            overlayLayout.addWidget(card)
            mainLayout.addWidget(overlay)
        
        def toggleEdit(self):
            """切换编辑模式"""
            self.is_editing = not self.is_editing
            if self.is_editing:
                self.promptEdit.setReadOnly(False)
                self.promptEdit.setStyleSheet("""
                    QTextEdit {
                        background-color: white;
                        padding: 16px;
                        border-radius: 8px;
                        font-size: 13px;
                        color: #374151;
                        font-family: Consolas, monospace;
                        border: 2px solid #16A34A;
                    }
                """)
                self.editBtn.setText("👁 预览")
                self.saveBtnWidget.setVisible(True)
            else:
                self.promptEdit.setReadOnly(True)
                self.promptEdit.setStyleSheet("""
                    QTextEdit {
                        background-color: #F9FAFB;
                        padding: 16px;
                        border-radius: 8px;
                        font-size: 13px;
                        color: #374151;
                        font-family: Consolas, monospace;
                        border: 1px solid #E5E7EB;
                    }
                """)
                self.editBtn.setText("✏️ 修改")
                self.saveBtnWidget.setVisible(False)
        
        def cancelEdit(self):
            """取消编辑"""
            self.is_editing = False
            self.promptEdit.setReadOnly(True)
            self.promptEdit.setStyleSheet("""
                QTextEdit {
                    background-color: #F9FAFB;
                    padding: 16px;
                    border-radius: 8px;
                    font-size: 13px;
                    color: #374151;
                    font-family: Consolas, monospace;
                    border: 1px solid #E5E7EB;
                }
            """)
            self.editBtn.setText("✏️ 修改")
            self.saveBtnWidget.setVisible(False)
        
        def saveEdit(self):
            """保存修改"""
            new_content = self.promptEdit.toPlainText()
            self.prompt_updated.emit(self.template_index, new_content)
            
            # 切换回预览模式
            self.is_editing = False
            self.promptEdit.setReadOnly(True)
            self.promptEdit.setStyleSheet("""
                QTextEdit {
                    background-color: #F9FAFB;
                    padding: 16px;
                    border-radius: 8px;
                    font-size: 13px;
                    color: #374151;
                    font-family: Consolas, monospace;
                    border: 1px solid #E5E7EB;
                }
            """)
            self.editBtn.setText("✏️ 修改")
            self.saveBtnWidget.setVisible(False)
            
            InfoBar.success(
                title="保存成功",
                content="模板提示词已更新",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
        
        def mousePressEvent(self, event):
            # 点击遮罩关闭
            if event.pos().x() < 50 or event.pos().x() > self.width() - 50 or \
               event.pos().y() < 50 or event.pos().y() > self.height() - 50:
                self.close()
            super().mousePressEvent(event)
    
    class CreateTemplateDialog(QDialog):
        """创建模板弹窗（4.2）"""
        template_created = pyqtSignal(dict)  # 模板创建信号
        
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setModal(True)
            
            # 主布局
            mainLayout = QHBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            
            # 半透明遮罩
            overlay = QWidget()
            overlay.setStyleSheet("background-color: rgba(0, 0, 0, 100);")
            overlayLayout = QVBoxLayout(overlay)
            overlayLayout.setAlignment(Qt.AlignCenter)
            
            # 弹窗卡片
            card = QFrame()
            card.setFixedSize(750, 720)
            card.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 16px;
                    border: 1px solid #ECECEC;
                }
            """)
            cardLayout = QVBoxLayout(card)
            cardLayout.setContentsMargins(24, 24, 24, 24)
            cardLayout.setSpacing(16)
            
            # 头部
            headerLayout = QHBoxLayout()
            titleLabel = QLabel("创建模板")
            titleLabel.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            headerLayout.addWidget(titleLabel)
            headerLayout.addStretch()
            
            closeBtn = QPushButton("✕")
            closeBtn.setFixedSize(32, 32)
            closeBtn.setCursor(Qt.PointingHandCursor)
            closeBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F3F4F6;
                    border: none;
                    border-radius: 16px;
                    font-size: 16px;
                    color: #666666;
                }
                QPushButton:hover {
                    background-color: #E5E7EB;
                }
            """)
            closeBtn.clicked.connect(self.close)
            headerLayout.addWidget(closeBtn)
            cardLayout.addLayout(headerLayout)
            
            # 副标题
            subtitleLabel = QLabel("支持 Markdown 格式，AI 将参考此结构生成报告内容")
            subtitleLabel.setStyleSheet("font-size: 13px; color: #666666; border: none; background: transparent;")
            cardLayout.addWidget(subtitleLabel)
            
            # 模板名称
            nameLabel = QLabel("模板名称")
            nameLabel.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            cardLayout.addWidget(nameLabel)
            
            self.nameInput = QLineEdit()
            self.nameInput.setPlaceholderText("输入模板名称")
            self.nameInput.setStyleSheet("""
                QLineEdit {
                    padding: 10px 12px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    color: #1a1a1a;
                    background-color: white;
                }
                QLineEdit:focus {
                    border: 1px solid #16A34A;
                }
            """)
            cardLayout.addWidget(self.nameInput)
            
            # 模板简介
            introLabel = QLabel("模板简介")
            introLabel.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            cardLayout.addWidget(introLabel)
            
            self.introInput = QLineEdit()
            self.introInput.setPlaceholderText("简短描述模板用途，如：适合向领导汇报的简洁日报")
            self.introInput.setStyleSheet("""
                QLineEdit {
                    padding: 10px 12px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    color: #1a1a1a;
                    background-color: white;
                }
                QLineEdit:focus {
                    border: 1px solid #16A34A;
                }
            """)
            cardLayout.addWidget(self.introInput)
            
            # 模板正文
            bodyLabel = QLabel("模板正文")
            bodyLabel.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            cardLayout.addWidget(bodyLabel)
            
            self.bodyInput = QTextEdit()
            self.bodyInput.setPlaceholderText("## 周报 [日期]\n\n### 本周完成\n- [工作项]\n\n### 下周计划\n- [计划项]")
            self.bodyInput.setMinimumHeight(150)
            self.bodyInput.setStyleSheet("""
                QTextEdit {
                    padding: 10px 12px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    color: #1a1a1a;
                    background-color: white;
                    font-family: Consolas, monospace;
                }
                QTextEdit:focus {
                    border: 1px solid #16A34A;
                }
            """)
            cardLayout.addWidget(self.bodyInput)
            
            bodyHint = QLabel("支持 Markdown 格式，AI 将参考此结构生成报告内容")
            bodyHint.setStyleSheet("font-size: 11px; color: #9CA3AF; border: none; background: transparent;")
            cardLayout.addWidget(bodyHint)
            
            # 自定义指令
            instrLabel = QLabel("自定义指令")
            instrLabel.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            cardLayout.addWidget(instrLabel)
            
            self.instrInput = QTextEdit()
            self.instrInput.setPlaceholderText("例如：用表格输出耗时，结尾补充风险和明日计划，语气保持简洁")
            self.instrInput.setMinimumHeight(80)
            self.instrInput.setStyleSheet("""
                QTextEdit {
                    padding: 10px 12px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    color: #1a1a1a;
                    background-color: white;
                }
                QTextEdit:focus {
                    border: 1px solid #16A34A;
                }
            """)
            cardLayout.addWidget(self.instrInput)
            
            instrHint = QLabel("写模板固定的输出要求，会随该模板一起保存；生成时还可以叠加本次自定义指令")
            instrHint.setStyleSheet("font-size: 11px; color: #9CA3AF; border: none; background: transparent;")
            cardLayout.addWidget(instrHint)
            
            cardLayout.addStretch()
            
            # 底部按钮
            btnLayout = QHBoxLayout()
            btnLayout.addStretch()
            
            cancelBtn = QPushButton("取消")
            cancelBtn.setCursor(Qt.PointingHandCursor)
            cancelBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 10px 20px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            cancelBtn.clicked.connect(self.close)
            btnLayout.addWidget(cancelBtn)
            
            saveBtn = QPushButton("保存模板")
            saveBtn.setCursor(Qt.PointingHandCursor)
            saveBtn.setStyleSheet("""
                QPushButton {
                    background-color: #16A34A;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #15803D;
                }
            """)
            saveBtn.clicked.connect(self.saveTemplate)
            btnLayout.addWidget(saveBtn)
            
            cardLayout.addLayout(btnLayout)
            
            overlayLayout.addWidget(card)
            mainLayout.addWidget(overlay)
        
        def saveTemplate(self):
            name = self.nameInput.text().strip()
            intro = self.introInput.text().strip()
            body = self.bodyInput.toPlainText().strip()
            if not name:
                return
            template = {
                "name": name,
                "intro": intro if intro else body[:30] + "..." if len(body) > 30 else body,
                "desc": body[:50] + "..." if len(body) > 50 else body,
                "is_cloud": False,
                "prompt": body
            }
            self.template_created.emit(template)
            self.close()
    
    class GenerateConfirmDialog(QDialog):
        """生成报告确认弹窗（4.3）"""
        stay_here = pyqtSignal()  # 留在此页信号
        go_history = pyqtSignal()  # 查看历史报告信号
        
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setModal(True)
            
            # 主布局
            mainLayout = QHBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            
            # 半透明遮罩
            overlay = QWidget()
            overlay.setStyleSheet("background-color: rgba(0, 0, 0, 100);")
            overlayLayout = QVBoxLayout(overlay)
            overlayLayout.setAlignment(Qt.AlignCenter)
            
            # 弹窗卡片
            card = QFrame()
            card.setFixedSize(450, 220)
            card.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 16px;
                    border: 1px solid #ECECEC;
                }
            """)
            cardLayout = QVBoxLayout(card)
            cardLayout.setContentsMargins(24, 24, 24, 24)
            cardLayout.setSpacing(16)
            
            # 标题
            titleLabel = QLabel("报告正在生成中")
            titleLabel.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            cardLayout.addWidget(titleLabel)
            
            # 描述
            descLabel = QLabel("已提交报告生成任务，AI 正在为你撰写报告。\n是否跳转到历史报告页面查看进度？")
            descLabel.setStyleSheet("font-size: 13px; color: #666666; border: none; background: transparent;")
            descLabel.setWordWrap(True)
            cardLayout.addWidget(descLabel)
            
            cardLayout.addStretch()
            
            # 底部按钮
            btnLayout = QHBoxLayout()
            btnLayout.addStretch()
            
            stayBtn = QPushButton("留在此页")
            stayBtn.setCursor(Qt.PointingHandCursor)
            stayBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 10px 20px;
                    border: 1px solid #E5E7EB;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            stayBtn.clicked.connect(self.onStay)
            btnLayout.addWidget(stayBtn)
            
            goBtn = QPushButton("查看历史报告")
            goBtn.setCursor(Qt.PointingHandCursor)
            goBtn.setStyleSheet("""
                QPushButton {
                    background-color: #16A34A;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #15803D;
                }
            """)
            goBtn.clicked.connect(self.onGoHistory)
            btnLayout.addWidget(goBtn)
            
            cardLayout.addLayout(btnLayout)
            
            overlayLayout.addWidget(card)
            mainLayout.addWidget(overlay)
        
        def onStay(self):
            self.stay_here.emit()
            self.close()
        
        def onGoHistory(self):
            self.go_history.emit()
            self.close()
    
    class ReportResultDialog(QDialog):
        """报告生成结果弹窗（4.4）- 支持流式输出"""
        report_generated = pyqtSignal()  # 报告生成完成信号
        
        def __init__(self, report_type, date_range, template_name, template_prompt, parent=None):
            super().__init__(parent)
            self.report_type = report_type
            self.date_range = date_range
            self.template_name = template_name
            self.template_prompt = template_prompt
            self.is_generating = True
            self.full_content = ""
            
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setModal(True)
            
            # 主布局
            mainLayout = QHBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            
            # 半透明遮罩
            overlay = QWidget()
            overlay.setStyleSheet("background-color: rgba(0, 0, 0, 100);")
            overlayLayout = QVBoxLayout(overlay)
            overlayLayout.setAlignment(Qt.AlignCenter)
            
            # 弹窗卡片（接近全屏）
            card = QFrame()
            card.setMinimumSize(900, 700)
            card.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 16px;
                    border: 1px solid #ECECEC;
                }
            """)
            cardLayout = QVBoxLayout(card)
            cardLayout.setContentsMargins(24, 24, 24, 24)
            cardLayout.setSpacing(16)
            
            # 头部
            headerLayout = QHBoxLayout()
            
            self.titleLabel = QLabel(f"{report_type}报告正在生成中...")
            self.titleLabel.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            headerLayout.addWidget(self.titleLabel)
            
            self.statusTag = QLabel("生成中")
            self.statusTag.setStyleSheet("""
                QLabel {
                    color: #F59E0B;
                    font-size: 13px;
                    font-weight: bold;
                    border: none;
                    background: transparent;
                }
            """)
            headerLayout.addWidget(self.statusTag)
            headerLayout.addStretch()
            
            closeBtn = QPushButton("✕")
            closeBtn.setFixedSize(32, 32)
            closeBtn.setCursor(Qt.PointingHandCursor)
            closeBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F3F4F6;
                    border: none;
                    border-radius: 16px;
                    font-size: 16px;
                    color: #666666;
                }
                QPushButton:hover {
                    background-color: #E5E7EB;
                }
            """)
            closeBtn.clicked.connect(self.close)
            headerLayout.addWidget(closeBtn)
            cardLayout.addLayout(headerLayout)
            
            # 副标题
            self.subtitleLabel = QLabel(f"工作{report_type} — {date_range}")
            self.subtitleLabel.setStyleSheet("font-size: 13px; color: #666666; border: none; background: transparent;")
            cardLayout.addWidget(self.subtitleLabel)
            
            # 分隔线
            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setStyleSheet("background-color: #F3F4F6; border: none; height: 1px;")
            cardLayout.addWidget(separator)
            
            # 报告内容区域
            self.contentEdit = QTextEdit()
            self.contentEdit.setReadOnly(True)
            self.contentEdit.setStyleSheet("""
                QTextEdit {
                    font-size: 14px;
                    color: #374151;
                    border: none;
                    background: transparent;
                    padding: 16px;
                    font-family: "Microsoft YaHei", sans-serif;
                }
            """)
            cardLayout.addWidget(self.contentEdit)
            
            # 分隔线
            separator2 = QFrame()
            separator2.setFrameShape(QFrame.HLine)
            separator2.setStyleSheet("background-color: #F3F4F6; border: none; height: 1px;")
            cardLayout.addWidget(separator2)
            
            # 底部栏
            footerLayout = QHBoxLayout()
            
            self.infoLabel = QLabel(f"模板：{template_name} · 0 字")
            self.infoLabel.setStyleSheet("font-size: 12px; color: #9CA3AF; border: none; background: transparent;")
            footerLayout.addWidget(self.infoLabel)
            
            # token 用量显示（生成完成后填充）
            self.tokenLabel = QLabel("")
            self.tokenLabel.setStyleSheet("font-size: 12px; color: #9CA3AF; border: none; background: transparent;")
            self.tokenLabel.setToolTip("本次报告生成消耗的 Token")
            footerLayout.addWidget(self.tokenLabel)
            footerLayout.addStretch()
            
            # 操作按钮
            self.actionBtns = {}
            for icon, text, name in [("📋", "复制", "copy"), ("📥", "导出", "export"), ("🔄", "重新生成", "regenerate")]:
                btn = QPushButton(f"{icon} {text}")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setEnabled(False)  # 生成完成前禁用
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: white;
                        color: #374151;
                        padding: 8px 16px;
                        border: 1px solid #E5E7EB;
                        border-radius: 8px;
                        font-size: 12px;
                        border: none;
                    }
                    QPushButton:hover {
                        background-color: #F9FAFB;
                    }
                    QPushButton:disabled {
                        color: #9CA3AF;
                        background-color: #F3F4F6;
                    }
                """)
                if name == "copy":
                    btn.clicked.connect(self.copyContent)
                elif name == "export":
                    btn.clicked.connect(self.exportContent)
                elif name == "regenerate":
                    btn.clicked.connect(self.regenerate)
                self.actionBtns[name] = btn
                footerLayout.addWidget(btn)
            
            cardLayout.addLayout(footerLayout)
            
            overlayLayout.addWidget(card)
            mainLayout.addWidget(overlay)
            
            # 启动生成
            self.startGeneration()
        
        def startGeneration(self):
            """启动报告生成"""
            # 解析日期范围
            dates = self.date_range.split(" 至 ")
            start_date = dates[0]
            end_date = dates[1] if len(dates) > 1 else start_date
            
            # 创建工作线程
            self.worker = ReportGenerateWorker(
                self.template_prompt,
                start_date,
                end_date,
                self.report_type
            )
            
            # 连接信号
            self.worker.chunk_received.connect(self.onChunkReceived)
            self.worker.generation_finished.connect(self.onGenerationFinished)
            self.worker.generation_error.connect(self.onGenerationError)
            
            # 启动线程
            self.worker.start()
        
        def onChunkReceived(self, chunk):
            """接收到一块内容"""
            self.full_content += chunk
            self.contentEdit.setPlainText(self.full_content)
            # 滚动到底部
            self.contentEdit.verticalScrollBar().setValue(
                self.contentEdit.verticalScrollBar().maximum()
            )
            # 更新字数
            self.infoLabel.setText(f"模板：{self.template_name} · {len(self.full_content)} 字")
        
        def onGenerationFinished(self, content):
            """生成完成"""
            self.is_generating = False
            self.full_content = content
            self.contentEdit.setPlainText(content)
            
            # 读取本次生成消耗的 token 并显示
            usage = getattr(self.worker, 'usage', None) or {"input": 0, "output": 0, "total": 0}
            self.token_usage = usage
            from store import format_token_count
            self.tokenLabel.setText(f"·  Token：{format_token_count(usage.get('total', 0))}")
            self.tokenLabel.setToolTip(
                "总 Token：" + str(usage.get('total', 0)) + "\n输入：" + str(usage.get('input', 0)) + "\n输出：" + str(usage.get('output', 0))
            )
            
            # 保存报告到文件（含 token 用量）
            from store import save_report
            title = f"工作{self.report_type} — {self.date_range}"
            save_report(title, content, self.report_type, self.template_name, tokens=usage)
            
            # 同步报告到服务器
            try:
                from api_sync import sync_report_generated, upload_report
                # 记录报告生成事件
                sync_report_generated()
                # 上传报告内容（携带本次生成的 token 用量 input/output/total）
                upload_report(title, content, self.report_type, tokens=usage)
            except Exception as e:
                print(f"[同步] 报告同步失败: {e}")
            
            # 发送报告生成完成信号
            self.report_generated.emit()
            
            # 更新UI
            self.titleLabel.setText(f"{self.report_type}报告生成完成")
            self.statusTag.setText("已完成")
            self.statusTag.setStyleSheet("""
                QLabel {
                    color: #16A34A;
                    font-size: 13px;
                    font-weight: bold;
                    border: none;
                    background: transparent;
                }
            """)
            
            # 启用操作按钮
            for btn in self.actionBtns.values():
                btn.setEnabled(True)
            
            # 更新字数
            self.infoLabel.setText(f"模板：{self.template_name} · {len(content)} 字")
        
        def onGenerationError(self, error):
            """生成出错"""
            self.is_generating = False
            
            self.titleLabel.setText(f"{self.report_type}报告生成失败")
            self.statusTag.setText("失败")
            self.statusTag.setStyleSheet("""
                QLabel {
                    color: #EF4444;
                    font-size: 13px;
                    font-weight: bold;
                    border: none;
                    background: transparent;
                }
            """)
            
            # 清空内容区域
            self.contentEdit.setPlainText("")
            
            # 显示错误弹窗
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.critical(
                self,
                "报告生成失败",
                f"报告生成失败：\n\n{error}\n\n请稍后再试。",
                QMessageBox.Ok
            )
        
        def copyContent(self):
            """复制内容"""
            QApplication.clipboard().setText(self.full_content)
            InfoBar.success(
                title="复制成功",
                content="报告内容已复制到剪贴板",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
        
        def exportContent(self):
            """导出内容"""
            from PyQt5.QtWidgets import QFileDialog
            file_path, _ = QFileDialog.getSaveFileName(
                self, "导出报告", f"{self.report_type}报告_{self.date_range.replace('至', '-').strip()}.md",
                "Markdown Files (*.md);;Text Files (*.txt);;All Files (*)"
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.full_content)
                InfoBar.success(
                    title="导出成功",
                    content=f"报告已保存到: {file_path}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
        
        def regenerate(self):
            """重新生成"""
            self.full_content = ""
            self.contentEdit.clear()
            self.is_generating = True
            self.titleLabel.setText(f"{self.report_type}报告正在生成中...")
            self.statusTag.setText("生成中")
            self.statusTag.setStyleSheet("""
                QLabel {
                    color: #F59E0B;
                    font-size: 13px;
                    font-weight: bold;
                    border: none;
                    background: transparent;
                }
            """)
            for btn in self.actionBtns.values():
                btn.setEnabled(False)
            self.startGeneration()
    
    class FlowLayout(QLayout):
        """流式布局，根据宽度动态调整每行数量"""
        def __init__(self, parent=None, margin=0, hSpacing=12, vSpacing=12):
            super().__init__(parent)
            self._hSpacing = hSpacing
            self._vSpacing = vSpacing
            self.setContentsMargins(margin, margin, margin, margin)
            self._items = []
        
        def addItem(self, item):
            self._items.append(item)
        
        def count(self):
            return len(self._items)
        
        def itemAt(self, index):
            if 0 <= index < len(self._items):
                return self._items[index]
            return None
        
        def takeAt(self, index):
            if 0 <= index < len(self._items):
                return self._items.pop(index)
            return None
        
        def expandingDirections(self):
            return Qt.Orientations(0)
        
        def hasHeightForWidth(self):
            return True
        
        def heightForWidth(self, width):
            return self._doLayout(QRect(0, 0, width, 0), testOnly=True)
        
        def setGeometry(self, rect):
            super().setGeometry(rect)
            self._doLayout(rect, testOnly=False)
        
        def sizeHint(self):
            return self.minimumSize()
        
        def minimumSize(self):
            size = QSize()
            for item in self._items:
                size = size.expandedTo(item.minimumSize())
            margins = self.contentsMargins()
            size += QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
            return size
        
        def _doLayout(self, rect, testOnly):
            margins = self.contentsMargins()
            effectiveRect = rect.adjusted(margins.left(), margins.top(), -margins.right(), -margins.bottom())
            x = effectiveRect.x()
            y = effectiveRect.y()
            lineHeight = 0
            
            for item in self._items:
                wid = item.widget()
                
                # 获取卡片实际大小
                itemWidth = 180
                itemHeight = 90
                
                # 检查是否需要换行（超出可用宽度）
                if x + itemWidth > effectiveRect.right() and x > effectiveRect.x():
                    x = effectiveRect.x()
                    y = y + lineHeight + self._vSpacing
                    lineHeight = 0
                
                if not testOnly:
                    item.setGeometry(QRect(QPoint(x, y), QSize(itemWidth, itemHeight)))
                
                x = x + itemWidth + self._hSpacing
                lineHeight = max(lineHeight, itemHeight)
            
            return y + lineHeight - rect.y() + margins.bottom()
    
    class ReportPage(QWidget):
        """生成报告页面"""
        report_generated = pyqtSignal()  # 报告生成完成信号
        
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setObjectName("reportPage")
            self.selected_template_index = 2  # 默认选中AI工作轨迹日报
            self.template_cards = []
            
            # 主布局
            mainLayout = QVBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            mainLayout.setSpacing(0)
            
            # 滚动区域
            scrollArea = QScrollArea()
            scrollArea.setWidgetResizable(True)
            scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scrollArea.setStyleSheet("QScrollArea { border: none; background-color: #F7F8F7; }")
            
            contentWidget = QWidget()
            contentWidget.setStyleSheet("background-color: #F7F8F7; border: none;")
            layout = QVBoxLayout(contentWidget)
            layout.setContentsMargins(16, 12, 16, 12)
            layout.setSpacing(3)
            
            # ========== 顶部标题栏 ==========
            headerCard = QFrame()
            headerCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 10px;
                    border: 1px solid #ECECEC;
                }
            """)
            headerLayout = QHBoxLayout(headerCard)
            headerLayout.setContentsMargins(12, 8, 12, 8)
            headerLayout.setSpacing(10)
            
            # 左侧标题
            titleLeftLayout = QVBoxLayout()
            titleTopLayout = QHBoxLayout()
            
            iconLabel = QLabel("✨")
            iconLabel.setFixedSize(20, 20)
            iconLabel.setAlignment(Qt.AlignCenter)
            iconLabel.setStyleSheet("""
                QLabel {
                    background-color: #F0FBF4;
                    border-radius: 5px;
                    font-size: 12px;
                    border: none;
                }
            """)
            titleTopLayout.addWidget(iconLabel)
            
            titleLabel = QLabel("报告配置")
            titleLabel.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            titleTopLayout.addWidget(titleLabel)
            titleTopLayout.addStretch()
            titleLeftLayout.addLayout(titleTopLayout)
            
            subtitleLabel = QLabel("配置参数后点击生成，AI 将基于工作记录自动撰写报告")
            subtitleLabel.setStyleSheet("font-size: 10px; color: #666666; border: none; background: transparent;")
            titleLeftLayout.addWidget(subtitleLabel)
            
            headerLayout.addLayout(titleLeftLayout, 1)
            
            # 右侧按钮
            btnLayout = QHBoxLayout()
            btnLayout.setSpacing(6)
            
            # 自定义指令按钮（带角标）
            instrBtnContainer = QWidget()
            instrBtnContainer.setStyleSheet("border: none; background: transparent;")
            instrBtnLayout = QVBoxLayout(instrBtnContainer)
            instrBtnLayout.setContentsMargins(0, 0, 0, 0)
            
            instrBtn = QPushButton("💬 自定义指令")
            instrBtn.setCursor(Qt.PointingHandCursor)
            instrBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 4px 10px;
                    border: 1px solid #E5E7EB;
                    border-radius: 5px;
                    font-size: 10px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            instrBtnLayout.addWidget(instrBtn)
            
            # 建议填写角标
            badge = QLabel("建议填写")
            badge.setFixedSize(42, 14)
            badge.setAlignment(Qt.AlignCenter)
            badge.setStyleSheet("""
                QLabel {
                    background-color: #16A34A;
                    color: white;
                    padding: 1px 3px;
                    border-radius: 7px;
                    font-size: 8px;
                    font-weight: bold;
                    border: none;
                }
            """)
            badge.setParent(instrBtn)
            badge.move(instrBtn.width() - 35, -5)
            
            btnLayout.addWidget(instrBtnContainer)
            
            # 开始生成按钮
            generateBtn = QPushButton("✨ 开始生成报告")
            generateBtn.setCursor(Qt.PointingHandCursor)
            generateBtn.setStyleSheet("""
                QPushButton {
                    background-color: #16A34A;
                    color: white;
                    padding: 4px 12px;
                    border: none;
                    border-radius: 5px;
                    font-size: 10px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #15803D;
                }
            """)
            generateBtn.clicked.connect(self.onGenerate)
            btnLayout.addWidget(generateBtn)
            
            headerLayout.addLayout(btnLayout)
            layout.addWidget(headerCard)
            
            # ========== 主体两栏布局 ==========
            bodyLayout = QHBoxLayout()
            bodyLayout.setSpacing(12)
            
            # 左栏
            leftLayout = QVBoxLayout()
            leftLayout.setSpacing(3)
            
            # 左栏第一张卡片：报告类型 + 时间范围
            typeCard = QFrame()
            typeCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 10px;
                    border: 1px solid #ECECEC;
                }
            """)
            typeCardLayout = QHBoxLayout(typeCard)
            typeCardLayout.setContentsMargins(12, 8, 12, 8)
            typeCardLayout.setSpacing(16)
            
            # 左块：报告类型
            reportTypeLayout = QVBoxLayout()
            reportTypeLayout.setSpacing(6)
            
            reportTypeTitle = QLabel("报告类型")
            reportTypeTitle.setStyleSheet("font-size: 12px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            reportTypeLayout.addWidget(reportTypeTitle)
            
            reportTypeDesc = QLabel("选择要生成的报告周期")
            reportTypeDesc.setStyleSheet("font-size: 9px; color: #666666; border: none; background: transparent;")
            reportTypeLayout.addWidget(reportTypeDesc)
            
            # 分段选择按钮
            self.typeButtons = []
            typeBtnLayout = QHBoxLayout()
            typeBtnLayout.setSpacing(4)
            
            for i, (text, days) in enumerate([("日报", 0), ("周报", 7), ("月报", 30)]):
                btn = QPushButton(text)
                btn.setCheckable(True)
                btn.setCursor(Qt.PointingHandCursor)
                btn.setMinimumHeight(24)
                btn.setProperty("days", days)
                btn.clicked.connect(lambda checked, idx=i: self.selectReportType(idx))
                
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: white;
                        color: #374151;
                        padding: 3px 10px;
                        border: 1px solid #E5E7EB;
                        border-radius: 5px;
                        font-size: 10px;
                        font-weight: bold;
                    }
                    QPushButton:checked {
                        background-color: #F0FBF4;
                        color: #16A34A;
                        border: 1px solid #16A34A;
                    }
                    QPushButton:hover {
                        background-color: #F9FAFB;
                    }
                """)
                self.typeButtons.append(btn)
                typeBtnLayout.addWidget(btn)
            
            self.typeButtons[0].setChecked(True)
            reportTypeLayout.addLayout(typeBtnLayout)
            reportTypeLayout.addStretch()
            
            typeCardLayout.addLayout(reportTypeLayout, 1)
            
            # 分隔线
            typeSep = QFrame()
            typeSep.setFrameShape(QFrame.VLine)
            typeSep.setStyleSheet("background-color: #F3F4F6; border: none; width: 1px;")
            typeCardLayout.addWidget(typeSep)
            
            # 右块：时间范围
            timeRangeLayout = QVBoxLayout()
            timeRangeLayout.setSpacing(6)
            
            timeRangeTitle = QLabel("时间范围")
            timeRangeTitle.setStyleSheet("font-size: 12px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            timeRangeLayout.addWidget(timeRangeTitle)
            
            timeRangeDesc = QLabel("默认根据报告类型确定，可手动修改")
            timeRangeDesc.setStyleSheet("font-size: 9px; color: #666666; border: none; background: transparent;")
            timeRangeLayout.addWidget(timeRangeDesc)
            
            # 日期选择器（使用 Fluent CalendarPicker）
            dateLayout = QHBoxLayout()
            dateLayout.setSpacing(6)
            
            self.startDateEdit = CalendarPicker()
            self.startDateEdit.setDate(QDate.currentDate())
            self.startDateEdit.setDateFormat("yyyy/MM/dd")
            self.startDateEdit.setFixedWidth(105)
            self.startDateEdit.dateChanged.connect(self.updatePreviewDate)
            dateLayout.addWidget(self.startDateEdit)
            
            toLabel = QLabel("至")
            toLabel.setStyleSheet("font-size: 10px; color: #666666; border: none; background: transparent;")
            dateLayout.addWidget(toLabel)
            
            self.endDateEdit = CalendarPicker()
            self.endDateEdit.setDate(QDate.currentDate())
            self.endDateEdit.setDateFormat("yyyy/MM/dd")
            self.endDateEdit.setFixedWidth(105)
            self.endDateEdit.dateChanged.connect(self.updatePreviewDate)
            dateLayout.addWidget(self.endDateEdit)
            
            timeRangeLayout.addLayout(dateLayout)
            timeRangeLayout.addStretch()
            
            typeCardLayout.addLayout(timeRangeLayout, 1)
            
            leftLayout.addWidget(typeCard)
            
            # 左栏第二张卡片：选择报告模板
            templateCard = QFrame()
            templateCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 10px;
                    border: 1px solid #ECECEC;
                }
            """)
            templateCardLayout = QVBoxLayout(templateCard)
            templateCardLayout.setContentsMargins(12, 3, 12, 3)
            templateCardLayout.setSpacing(10)
            
            # 卡片头部
            templateHeaderLayout = QHBoxLayout()
            
            templateIcon = QLabel("📄")
            templateIcon.setFixedSize(18, 18)
            templateIcon.setAlignment(Qt.AlignCenter)
            templateIcon.setStyleSheet("""
                QLabel {
                    background-color: #F0FBF4;
                    border-radius: 4px;
                    font-size: 10px;
                    border: none;
                }
            """)
            templateHeaderLayout.addWidget(templateIcon)
            
            templateTitle = QLabel("选择报告模板")
            templateTitle.setStyleSheet("font-size: 12px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            templateHeaderLayout.addWidget(templateTitle)
            templateHeaderLayout.addStretch()
            
            # 导入模板按钮
            importBtn = QPushButton("📥 导入")
            importBtn.setCursor(Qt.PointingHandCursor)
            importBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 3px 8px;
                    border: 1px solid #E5E7EB;
                    border-radius: 5px;
                    font-size: 9px;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            importBtn.clicked.connect(self.onImportTemplate)
            templateHeaderLayout.addWidget(importBtn)
            
            # 导出模板按钮
            exportBtn = QPushButton("📤 导出")
            exportBtn.setCursor(Qt.PointingHandCursor)
            exportBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 3px 8px;
                    border: 1px solid #E5E7EB;
                    border-radius: 5px;
                    font-size: 9px;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            exportBtn.clicked.connect(self.onExportTemplate)
            templateHeaderLayout.addWidget(exportBtn)
            
            # 创建模板按钮
            createBtn = QPushButton("＋ 创建")
            createBtn.setCursor(Qt.PointingHandCursor)
            createBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 3px 8px;
                    border: 1px solid #E5E7EB;
                    border-radius: 5px;
                    font-size: 9px;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #F9FAFB;
                }
            """)
            createBtn.clicked.connect(self.onCreateTemplate)
            templateHeaderLayout.addWidget(createBtn)
            
            templateCardLayout.addLayout(templateHeaderLayout)
            
            # 副标题
            templateSubtitle = QLabel("选择合适的模板，AI 将为你生成更贴合需求的报告")
            templateSubtitle.setStyleSheet("font-size: 9px; color: #666666; border: none; background: transparent;")
            templateCardLayout.addWidget(templateSubtitle)
            
            # 模板网格（使用流式布局，根据宽度动态调整列数）
            templateGridWidget = QWidget()
            templateGridWidget.setStyleSheet("background: transparent; border: none;")
            self.templateFlowLayout = FlowLayout(templateGridWidget, margin=0, hSpacing=8, vSpacing=8)
            
            for i, template in enumerate(REPORT_TEMPLATES):
                card = TemplateCard(i, template["name"], template.get("intro", template["desc"]), template.get("is_cloud", True))
                card.clicked.connect(self.selectTemplate)
                card.preview_clicked.connect(self.showTemplatePreview)
                card.delete_clicked.connect(self.deleteTemplate)
                self.template_cards.append(card)
                self.templateFlowLayout.addWidget(card)
            
            # 将模板网格放在滚动区域中
            templateScrollArea = QScrollArea()
            templateScrollArea.setWidgetResizable(True)
            templateScrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            templateScrollArea.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            templateScrollArea.setStyleSheet("QScrollArea { border: none; background: transparent; } QScrollBar { width: 0px; height: 0px; }")
            templateScrollArea.setWidget(templateGridWidget)
            
            templateCardLayout.addWidget(templateScrollArea)
            
            # 设置模板卡片的下边距为3px
            templateCardLayout.setContentsMargins(12, 3, 12, 3)
            
            leftLayout.addWidget(templateCard)
            
            # 右栏：模板预览
            previewCard = QFrame()
            previewCard.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 10px;
                    border: 1px solid #ECECEC;
                }
            """)
            previewCardLayout = QVBoxLayout(previewCard)
            previewCardLayout.setContentsMargins(12, 8, 12, 8)
            previewCardLayout.setSpacing(10)
            
            # 预览标题
            previewTitle = QLabel("模板预览")
            previewTitle.setStyleSheet("font-size: 12px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            previewCardLayout.addWidget(previewTitle)
            
            # 预览头部
            previewHeaderLayout = QHBoxLayout()
            
            previewIcon = QLabel("📄")
            previewIcon.setFixedSize(20, 20)
            previewIcon.setAlignment(Qt.AlignCenter)
            previewIcon.setStyleSheet("""
                QLabel {
                    background-color: #F0FBF4;
                    border-radius: 5px;
                    font-size: 12px;
                    border: none;
                }
            """)
            previewHeaderLayout.addWidget(previewIcon)
            
            previewInfoLayout = QVBoxLayout()
            previewNameLayout = QHBoxLayout()
            self.previewNameLabel = QLabel(REPORT_TEMPLATES[2]["name"])
            self.previewNameLabel.setStyleSheet("font-size: 12px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            previewNameLayout.addWidget(self.previewNameLabel)
            
            previewCloudTag = QLabel("云端")
            previewCloudTag.setStyleSheet("""
                QLabel {
                    background-color: #E3F7EA;
                    color: #16A34A;
                    padding: 1px 5px;
                    border-radius: 7px;
                    font-size: 9px;
                    font-weight: bold;
                    border: none;
                }
            """)
            previewNameLayout.addWidget(previewCloudTag)
            previewNameLayout.addStretch()
            previewInfoLayout.addLayout(previewNameLayout)
            
            self.previewDateLabel = QLabel(f"时间范围：{QDate.currentDate().toString('yyyy-MM-dd')} 至 {QDate.currentDate().toString('yyyy-MM-dd')}")
            self.previewDateLabel.setStyleSheet("font-size: 9px; color: #666666; border: none; background: transparent;")
            previewInfoLayout.addWidget(self.previewDateLabel)
            
            previewHeaderLayout.addLayout(previewInfoLayout)
            previewCardLayout.addLayout(previewHeaderLayout)
            
            # 分隔线
            previewSep = QFrame()
            previewSep.setFrameShape(QFrame.HLine)
            previewSep.setStyleSheet("background-color: #F3F4F6; border: none; height: 1px;")
            previewCardLayout.addWidget(previewSep)
            
            # 预览内容区域
            self.previewContent = QLabel()
            self.previewContent.setWordWrap(True)
            self.previewContent.setStyleSheet("""
                QLabel {
                    background-color: #F9FAFB;
                    padding: 8px;
                    border-radius: 5px;
                    font-size: 10px;
                    color: #374151;
                    font-family: Consolas, monospace;
                    border: 1px solid #E5E7EB;
                }
            """)
            self.previewContent.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            self.previewContent.setMinimumHeight(150)
            
            previewScroll = QScrollArea()
            previewScroll.setWidget(self.previewContent)
            previewScroll.setWidgetResizable(True)
            previewScroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            previewScroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            previewScroll.setStyleSheet("QScrollArea { border: none; background: transparent; } QScrollBar { width: 0px; height: 0px; }")
            previewCardLayout.addWidget(previewScroll)
            
            # 底部提示
            previewHint = QLabel("实际内容将基于你的工作记录自动生成")
            previewHint.setStyleSheet("font-size: 9px; color: #9CA3AF; border: none; background: transparent;")
            previewHint.setAlignment(Qt.AlignCenter)
            previewCardLayout.addWidget(previewHint)
            
            # 使用 QSplitter 实现可调整的左右布局
            self.splitter = QSplitter(Qt.Horizontal)
            self.splitter.setHandleWidth(8)
            self.splitter.setStyleSheet("""
                QSplitter::handle {
                    background-color: #E5E7EB;
                    border-radius: 4px;
                    margin: 4px 0;
                }
                QSplitter::handle:hover {
                    background-color: #D1D5DB;
                }
            """)
            
            # 左栏容器
            leftWidget = QWidget()
            leftWidget.setLayout(leftLayout)
            self.splitter.addWidget(leftWidget)
            
            # 右栏容器
            self.splitter.addWidget(previewCard)
            
            # 设置初始大小比例
            self.splitter.setSizes([700, 300])
            
            bodyLayout.addWidget(self.splitter)
            
            layout.addLayout(bodyLayout)
            
            scrollArea.setWidget(contentWidget)
            mainLayout.addWidget(scrollArea)
            
            # 初始化
            self.selectReportType(0)
            self.selectTemplate(2)
        
        def selectReportType(self, index):
            """选择报告类型"""
            for i, btn in enumerate(self.typeButtons):
                btn.setChecked(i == index)
            
            today = QDate.currentDate()
            if index == 0:  # 日报
                self.startDateEdit.setDate(today)
                self.endDateEdit.setDate(today)
            elif index == 1:  # 周报
                monday = today.addDays(-(today.dayOfWeek() - 1))
                sunday = monday.addDays(6)
                self.startDateEdit.setDate(monday)
                self.endDateEdit.setDate(sunday)
            elif index == 2:  # 月报
                first = QDate(today.year(), today.month(), 1)
                last = first.addMonths(1).addDays(-1)
                self.startDateEdit.setDate(first)
                self.endDateEdit.setDate(last)
            
            self.updatePreviewDate()
        
        def selectTemplate(self, index):
            """选择模板"""
            self.selected_template_index = index
            for i, card in enumerate(self.template_cards):
                card.setSelected(i == index)
            
            template = REPORT_TEMPLATES[index]
            self.previewNameLabel.setText(template["name"])
            self.previewContent.setText(template["prompt"])
        
        def updatePreviewDate(self):
            """更新预览日期"""
            # 验证日期范围
            if self.startDateEdit.date > self.endDateEdit.date:
                InfoBar.warning(
                    title="日期范围错误",
                    content="开始日期不能晚于结束日期",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                return
            
            start = self.startDateEdit.date.toString("yyyy-MM-dd")
            end = self.endDateEdit.date.toString("yyyy-MM-dd")
            self.previewDateLabel.setText(f"时间范围：{start} 至 {end}")
        
        def showTemplatePreview(self, index):
            """显示模板预览弹窗"""
            template = REPORT_TEMPLATES[index]
            dialog = TemplatePreviewDialog(index, template["name"], template["desc"], template["prompt"], self)
            dialog.prompt_updated.connect(self.onPromptUpdated)
            dialog.exec_()
        
        def onPromptUpdated(self, index, new_content):
            """更新模板提示词"""
            REPORT_TEMPLATES[index]["prompt"] = new_content
            # 如果更新的是当前选中的模板，刷新预览
            if index == self.selected_template_index:
                self.previewContent.setText(new_content)
            # 保存到文件
            write_templates(REPORT_TEMPLATES)
        
        def onImportTemplate(self):
            """导入模板"""
            from PyQt5.QtWidgets import QFileDialog
            file_path, _ = QFileDialog.getOpenFileName(
                self, "导入模板", "",
                "CSV Files (*.csv);;All Files (*)"
            )
            if file_path:
                templates = import_templates(file_path)
                if templates:
                    # 添加导入的模板
                    for template in templates:
                        REPORT_TEMPLATES.append(template)
                        add_template(template)
                    
                    # 重建模板网格
                    self.rebuildTemplateGrid()
                    
                    InfoBar.success(
                        title="导入成功",
                        content=f"已导入 {len(templates)} 个模板",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=2000,
                        parent=self
                    )
                else:
                    InfoBar.error(
                        title="导入失败",
                        content="无法解析模板文件",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
        
        def onExportTemplate(self):
            """导出模板"""
            from PyQt5.QtWidgets import QFileDialog
            file_path, _ = QFileDialog.getSaveFileName(
                self, "导出模板", "report_templates.csv",
                "CSV Files (*.csv);;All Files (*)"
            )
            if file_path:
                if export_templates(file_path):
                    InfoBar.success(
                        title="导出成功",
                        content=f"模板已保存到: {file_path}",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                else:
                    InfoBar.error(
                        title="导出失败",
                        content="无法保存模板文件",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
        
        def onCreateTemplate(self):
            """打开创建模板弹窗"""
            dialog = CreateTemplateDialog(self)
            dialog.template_created.connect(self.addNewTemplate)
            dialog.exec_()
        
        def addNewTemplate(self, template):
            """添加新模板"""
            REPORT_TEMPLATES.append(template)
            add_template(template)  # 保存到文件
            
            # 重建模板网格
            self.rebuildTemplateGrid()
            
            InfoBar.success(
                title="创建成功",
                content=f"模板「{template['name']}」已添加",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
        
        def rebuildTemplateGrid(self):
            """重建模板网格"""
            # 清空现有卡片
            for card in self.template_cards:
                card.deleteLater()
            self.template_cards.clear()
            
            # 清空流式布局
            while self.templateFlowLayout.count():
                item = self.templateFlowLayout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            
            # 重新创建模板卡片
            for i, template in enumerate(REPORT_TEMPLATES):
                card = TemplateCard(i, template["name"], template.get("intro", template["desc"]), template.get("is_cloud", True))
                card.clicked.connect(self.selectTemplate)
                card.preview_clicked.connect(self.showTemplatePreview)
                card.delete_clicked.connect(self.deleteTemplate)
                self.template_cards.append(card)
                self.templateFlowLayout.addWidget(card)
        
        def deleteTemplate(self, index):
            """删除模板"""
            if len(REPORT_TEMPLATES) <= 1:
                InfoBar.warning(
                    title="无法删除",
                    content="至少需要保留一个模板",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
                return
            
            template_name = REPORT_TEMPLATES[index]["name"]
            
            # 确认删除
            reply = QMessageBox.question(
                self, "确认删除",
                f"确定要删除模板「{template_name}」吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                # 从数据中删除
                REPORT_TEMPLATES.pop(index)
                delete_template(index)  # 保存到文件
                
                # 重建模板网格
                self.rebuildTemplateGrid()
                
                # 如果删除的是当前选中的模板，选中第一个
                if self.selected_template_index >= len(REPORT_TEMPLATES):
                    self.selected_template_index = 0
                
                self.selectTemplate(self.selected_template_index)
                
                InfoBar.success(
                    title="删除成功",
                    content=f"模板「{template_name}」已删除",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
        
        def onGenerate(self):
            """开始生成报告"""
            # 验证日期范围
            if self.startDateEdit.date > self.endDateEdit.date:
                InfoBar.warning(
                    title="日期范围错误",
                    content="开始日期不能晚于结束日期",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                return
            
            # 获取报告参数
            report_type = "日报"
            for i, btn in enumerate(self.typeButtons):
                if btn.isChecked():
                    report_type = ["日报", "周报", "月报"][i]
                    break
            
            start = self.startDateEdit.date.toString("yyyy-MM-dd")
            end = self.endDateEdit.date.toString("yyyy-MM-dd")
            date_range = f"{start} 至 {end}"
            template_name = REPORT_TEMPLATES[self.selected_template_index]["name"]
            template_prompt = REPORT_TEMPLATES[self.selected_template_index]["prompt"]
            
            # 在后台开始生成报告
            self.resultDialog = ReportResultDialog(report_type, date_range, template_name, template_prompt, self)
            self.resultDialog.report_generated.connect(self.report_generated.emit)
            
            # 显示确认弹窗
            dialog = GenerateConfirmDialog(self)
            dialog.go_history.connect(self.goToHistory)
            dialog.stay_here.connect(self.resultDialog.show)
            dialog.exec_()
        
        def simulateGenerate(self):
            """生成报告（保留兼容性）"""
            pass
        
        def goToHistory(self):
            """跳转到历史报告页面"""
            # 获取主窗口
            main_window = self.window()
            # 使用 MainWindow 的 switchToPage 方法
            if hasattr(main_window, 'switchToPage'):
                main_window.switchToPage('historyReportPage')
            # 刷新历史报告列表
            if hasattr(main_window, 'historyReportPage'):
                main_window.historyReportPage.refreshList()

    # ==================== 设置页面 ====================
    
    class ConnectionTestWorker(QThread):
        """连接测试工作线程"""
        finished = pyqtSignal(bool, str)  # (success, message)
        
        def __init__(self, test_type, host=None, model=None):
            super().__init__()
            self.test_type = test_type
            self.host = host
            self.model = model
        
        def run(self):
            try:
                if self.test_type == "glm":
                    success, message = test_glm_connection()
                elif self.test_type == "ollama":
                    if self.host:
                        set_ollama_config(self.host, self.model)
                    success, message = test_ollama_connection()
                else:
                    success, message = False, "未知的测试类型"
                self.finished.emit(success, message)
            except Exception as e:
                self.finished.emit(False, str(e))

    # ==================== 历史报告页面 ====================
    
    # 示例报告数据
    SAMPLE_REPORTS = [
        {
            "title": "工作周报 — 2026-07-20 至 2026-07-26",
            "type": "周报",
            "status": "已完成",
            "time": "今日 10:38",
            "word_count": 2070,
            "output_method": "直接输出",
            "model": "默认模型"
        },
        {
            "title": "工作日报 — 2026-07-04",
            "type": "日报",
            "status": "已完成",
            "time": "07月04日 17:42",
            "word_count": 1136,
            "output_method": "直接输出",
            "model": "默认模型"
        },
        {
            "title": "工作周报 — 2026-06-29 至 2026-07-05",
            "type": "周报",
            "status": "已完成",
            "time": "06月29日 12:26",
            "word_count": 2388,
            "output_method": "直接输出",
            "model": "默认模型"
        },
        {
            "title": "工作周报 — 2026-06-15 至 2026-06-22",
            "type": "周报",
            "status": "已完成",
            "time": "06月22日 14:59",
            "word_count": 3579,
            "output_method": "直接输出",
            "model": "默认模型"
        },
        {
            "title": "工作日报 — 2026-06-18",
            "type": "日报",
            "status": "已完成",
            "time": "06月18日 23:33",
            "word_count": 2374,
            "output_method": "直接输出",
            "model": "默认模型"
        },
        {
            "title": "工作日报 — 2026-06-17",
            "type": "日报",
            "status": "已完成",
            "time": "06月17日 18:20",
            "word_count": 1856,
            "output_method": "直接输出",
            "model": "默认模型"
        }
    ]
    
    class ReportItemWidget(QFrame):
        """报告条目控件"""
        view_clicked = pyqtSignal(dict)  # 查看信号
        copy_clicked = pyqtSignal(dict)  # 复制信号
        export_clicked = pyqtSignal(dict)  # 导出信号
        delete_clicked = pyqtSignal(dict)  # 删除信号
        
        def __init__(self, report_data, parent=None):
            super().__init__(parent)
            self.report_data = report_data
            self.setStyleSheet("QFrame { border: none; background: transparent; }")
            
            layout = QHBoxLayout(self)
            layout.setContentsMargins(20, 16, 20, 16)
            layout.setSpacing(16)
            
            # 左侧信息区
            infoLayout = QVBoxLayout()
            infoLayout.setSpacing(8)
            
            # 标题行
            titleLayout = QHBoxLayout()
            titleLayout.setSpacing(8)
            
            titleLabel = QLabel(report_data["title"])
            titleLabel.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            titleLayout.addWidget(titleLabel)
            
            # 类型标签
            report_type = report_data["type"]
            if report_type == "周报":
                typeTag = QLabel(report_type)
                typeTag.setStyleSheet("""
                    QLabel {
                        background-color: #e0f2fe;
                        color: #0284c7;
                        padding: 2px 8px;
                        border-radius: 4px;
                        font-size: 11px;
                        font-weight: bold;
                        border: none;
                    }
                """)
            else:
                typeTag = QLabel(report_type)
                typeTag.setStyleSheet("""
                    QLabel {
                        background-color: #dcfce7;
                        color: #16a34a;
                        padding: 2px 8px;
                        border-radius: 4px;
                        font-size: 11px;
                        font-weight: bold;
                        border: none;
                    }
                """)
            titleLayout.addWidget(typeTag)
            
            # 状态标签
            statusTag = QLabel(report_data["status"])
            statusTag.setStyleSheet("""
                QLabel {
                    background-color: #dcfce7;
                    color: #16a34a;
                    padding: 2px 8px;
                    border-radius: 4px;
                    font-size: 11px;
                    font-weight: bold;
                    border: none;
                }
            """)
            titleLayout.addWidget(statusTag)
            
            # Token 用量标签（默认只显示总数，鼠标悬停显示 总/输入/输出）
            token_total = report_data.get("token_total", 0)
            if token_total:
                from store import format_token_count
                tokenTag = QLabel(f"🪙 {format_token_count(token_total)}")
                tokenTag.setStyleSheet("""
                    QLabel {
                        background-color: #f3f4f6;
                        color: #6b7280;
                        padding: 2px 8px;
                        border-radius: 4px;
                        font-size: 11px;
                        border: none;
                    }
                """)
                tokenTag.setToolTip(
                    f"总 Token：{token_total}\n输入：{report_data.get('token_input', 0)}\n输出：{report_data.get('token_output', 0)}"
                )
                titleLayout.addWidget(tokenTag)
            
            titleLayout.addStretch()
            
            infoLayout.addLayout(titleLayout)
            
            # 元信息行
            metaText = f"{report_data['time']} · {report_data['word_count']} 字 · {report_data['output_method']} · {report_data['model']}"
            metaLabel = QLabel(metaText)
            metaLabel.setStyleSheet("font-size: 12px; color: #9ca3af; border: none; background: transparent;")
            infoLayout.addWidget(metaLabel)
            
            layout.addLayout(infoLayout, 1)
            
            # 右侧操作区
            actionLayout = QHBoxLayout()
            actionLayout.setSpacing(16)
            
            # 查看按钮
            viewBtn = QPushButton("👁 查看")
            viewBtn.setCursor(Qt.PointingHandCursor)
            viewBtn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #6b7280;
                    border: none;
                    font-size: 12px;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    color: #16a34a;
                }
            """)
            viewBtn.clicked.connect(lambda: self.view_clicked.emit(self.report_data))
            actionLayout.addWidget(viewBtn)
            
            # 复制按钮
            copyBtn = QPushButton("📋 复制")
            copyBtn.setCursor(Qt.PointingHandCursor)
            copyBtn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #6b7280;
                    border: none;
                    font-size: 12px;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    color: #16a34a;
                }
            """)
            copyBtn.clicked.connect(lambda: self.copy_clicked.emit(self.report_data))
            actionLayout.addWidget(copyBtn)
            
            # 导出按钮
            exportBtn = QPushButton("📥 导出")
            exportBtn.setCursor(Qt.PointingHandCursor)
            exportBtn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #6b7280;
                    border: none;
                    font-size: 12px;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    color: #16a34a;
                }
            """)
            exportBtn.clicked.connect(lambda: self.export_clicked.emit(self.report_data))
            actionLayout.addWidget(exportBtn)
            
            # 删除按钮
            deleteBtn = QPushButton("🗑 删除")
            deleteBtn.setCursor(Qt.PointingHandCursor)
            deleteBtn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #6b7280;
                    border: none;
                    font-size: 12px;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    color: #ef4444;
                }
            """)
            deleteBtn.clicked.connect(lambda: self.delete_clicked.emit(self.report_data))
            actionLayout.addWidget(deleteBtn)
            
            layout.addLayout(actionLayout)
    
    class ReportDrawer(QWidget):
        """报告详情抽屉面板"""
        closed = pyqtSignal()  # 关闭信号
        
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setVisible(False)
            self.report_data = None
            self.is_editing = False
            self.raw_content = ""  # 原始Markdown内容
            
            # 遮罩层
            self.overlay = QWidget(parent)
            self.overlay.setStyleSheet("background-color: rgba(0, 0, 0, 0.35);")
            self.overlay.setVisible(False)
            self.overlay.installEventFilter(self)
            
            # 抽屉主体
            self.drawer = QWidget(self)
            self.drawer.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-top-left-radius: 16px;
                    border-bottom-left-radius: 16px;
                }
            """)
            
            # 添加阴影效果
            shadow = QGraphicsDropShadowEffect(self.drawer)
            shadow.setBlurRadius(30)
            shadow.setXOffset(-10)
            shadow.setYOffset(0)
            shadow.setColor(QColor(0, 0, 0, 80))
            self.drawer.setGraphicsEffect(shadow)
            
            # 抽屉内部布局
            drawerLayout = QVBoxLayout(self.drawer)
            drawerLayout.setContentsMargins(0, 0, 0, 0)
            drawerLayout.setSpacing(0)
            
            # ========== 头部区 ==========
            headerWidget = QWidget()
            headerWidget.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-top-left-radius: 16px;
                    border-bottom: 1px solid #e5e7eb;
                }
            """)
            headerLayout = QVBoxLayout(headerWidget)
            headerLayout.setContentsMargins(32, 24, 32, 24)
            headerLayout.setSpacing(12)
            
            # 标题行
            titleRow = QHBoxLayout()
            titleRow.setSpacing(8)
            
            self.titleLabel = QLabel("")
            self.titleLabel.setStyleSheet("""
                QLabel {
                    font-size: 18px;
                    font-weight: bold;
                    color: #1a1a1a;
                    border: none;
                    background: transparent;
                }
            """)
            titleRow.addWidget(self.titleLabel, 1)
            
            # 编辑按钮
            editTitleBtn = QPushButton("✏️")
            editTitleBtn.setFixedSize(28, 28)
            editTitleBtn.setCursor(Qt.PointingHandCursor)
            editTitleBtn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    font-size: 14px;
                    color: #9ca3af;
                }
                QPushButton:hover {
                    color: #374151;
                }
            """)
            titleRow.addWidget(editTitleBtn)
            
            # 关闭按钮
            closeBtn = QPushButton("✕")
            closeBtn.setFixedSize(32, 32)
            closeBtn.setCursor(Qt.PointingHandCursor)
            closeBtn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    font-size: 18px;
                    color: #9ca3af;
                }
                QPushButton:hover {
                    color: #374151;
                }
            """)
            closeBtn.clicked.connect(self.close)
            titleRow.addWidget(closeBtn)
            
            headerLayout.addLayout(titleRow)
            
            # 标签 + 元信息行
            metaRow = QHBoxLayout()
            metaRow.setSpacing(8)
            
            self.typeTag = QLabel("")
            self.typeTag.setStyleSheet("""
                QLabel {
                    background-color: #e0f2fe;
                    color: #0284c7;
                    padding: 2px 8px;
                    border-radius: 4px;
                    font-size: 11px;
                    font-weight: bold;
                    border: none;
                }
            """)
            metaRow.addWidget(self.typeTag)
            
            self.metaLabel = QLabel("")
            self.metaLabel.setStyleSheet("""
                QLabel {
                    font-size: 12px;
                    color: #9ca3af;
                    border: none;
                    background: transparent;
                }
            """)
            metaRow.addWidget(self.metaLabel)
            metaRow.addStretch()
            
            headerLayout.addLayout(metaRow)
            
            drawerLayout.addWidget(headerWidget)
            
            # ========== 内容滚动区 ==========
            # 预览模式（QTextBrowser）
            self.contentBrowser = QTextBrowser()
            self.contentBrowser.setStyleSheet("""
                QTextBrowser {
                    border: none;
                    background: white;
                    font-size: 14px;
                    color: #1f2937;
                    padding: 32px 40px;
                }
            """)
            self.contentBrowser.setOpenExternalLinks(False)
            self.contentBrowser.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.contentBrowser.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            drawerLayout.addWidget(self.contentBrowser, 1)
            
            # 编辑模式（QTextEdit）- 默认隐藏
            self.contentEdit = QTextEdit()
            self.contentEdit.setStyleSheet("""
                QTextEdit {
                    border: none;
                    background: white;
                    font-size: 14px;
                    color: #1f2937;
                    padding: 32px 40px;
                    font-family: Consolas, monospace;
                }
            """)
            self.contentEdit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.contentEdit.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.contentEdit.setVisible(False)
            drawerLayout.addWidget(self.contentEdit, 1)
            
            # ========== 底部操作栏 ==========
            footerWidget = QWidget()
            footerWidget.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-bottom-left-radius: 16px;
                    border-top: 1px solid #e5e7eb;
                }
            """)
            footerLayout = QHBoxLayout(footerWidget)
            footerLayout.setContentsMargins(32, 16, 32, 16)
            footerLayout.setSpacing(12)
            footerLayout.addStretch()
            
            # 网页报告按钮（带角标）
            webBtnContainer = QWidget()
            webBtnContainer.setStyleSheet("background: transparent; border: none;")
            webBtnLayout = QHBoxLayout(webBtnContainer)
            webBtnLayout.setContentsMargins(0, 0, 0, 0)
            
            webBtn = QPushButton("🌐 网页报告")
            webBtn.setCursor(Qt.PointingHandCursor)
            webBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 8px 16px;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #f9fafb;
                }
            """)
            webBtnLayout.addWidget(webBtn)
            
            # 增值角标
            badge = QLabel("增值")
            badge.setStyleSheet("""
                QLabel {
                    background-color: #dcfce7;
                    color: #16a34a;
                    padding: 1px 6px;
                    border-radius: 8px;
                    font-size: 10px;
                    font-weight: bold;
                    border: none;
                }
            """)
            badge.setParent(webBtn)
            badge.move(webBtn.width() - 15, -5)
            
            footerLayout.addWidget(webBtnContainer)
            
            # 编辑按钮
            self.editBtn = QPushButton("✏️ 编辑")
            self.editBtn.setCursor(Qt.PointingHandCursor)
            self.editBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 8px 16px;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #f9fafb;
                }
            """)
            self.editBtn.clicked.connect(self.toggleEdit)
            footerLayout.addWidget(self.editBtn)
            
            # 复制全文按钮
            copyBtn = QPushButton("📋 复制全文")
            copyBtn.setCursor(Qt.PointingHandCursor)
            copyBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 8px 16px;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #f9fafb;
                }
            """)
            copyBtn.clicked.connect(self.copyContent)
            footerLayout.addWidget(copyBtn)
            
            # 导出按钮
            exportBtn = QPushButton("📥 导出")
            exportBtn.setCursor(Qt.PointingHandCursor)
            exportBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #374151;
                    padding: 8px 16px;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #f9fafb;
                }
            """)
            exportBtn.clicked.connect(self.exportContent)
            footerLayout.addWidget(exportBtn)
            
            drawerLayout.addWidget(footerWidget)
        
        def eventFilter(self, obj, event):
            """事件过滤器，点击遮罩关闭抽屉"""
            if obj == self.overlay and event.type() == event.MouseButtonPress:
                self.close()
                return True
            return super().eventFilter(obj, event)
        
        def open(self, report_data):
            """打开抽屉"""
            self.report_data = report_data
            self.is_editing = False
            
            # 更新头部信息
            self.titleLabel.setText(report_data.get('title', ''))
            
            # 更新类型标签
            report_type = report_data.get('type', '日报')
            if report_type == "周报":
                self.typeTag.setText("周报")
                self.typeTag.setStyleSheet("""
                    QLabel {
                        background-color: #e0f2fe;
                        color: #0284c7;
                        padding: 2px 8px;
                        border-radius: 4px;
                        font-size: 11px;
                        font-weight: bold;
                        border: none;
                    }
                """)
            else:
                self.typeTag.setText("日报")
                self.typeTag.setStyleSheet("""
                    QLabel {
                        background-color: #dcfce7;
                        color: #16a34a;
                        padding: 2px 8px;
                        border-radius: 4px;
                        font-size: 11px;
                        font-weight: bold;
                        border: none;
                    }
                """)
            
            # 更新元信息
            meta_text = f"{report_data.get('time', '')} · {report_data.get('word_count', 0)} 字 · {report_data.get('output_method', '')} · {report_data.get('model', '')}"
            self.metaLabel.setText(meta_text)
            
            # 读取报告内容
            from store import read_report
            self.raw_content = read_report(report_data.get('filepath', ''))
            
            # 将 Markdown 转换为 HTML
            html_content = self.markdown_to_html(self.raw_content)
            self.contentBrowser.setHtml(html_content)
            self.contentEdit.setPlainText(self.raw_content)
            
            # 确保显示预览模式
            self.contentBrowser.setVisible(True)
            self.contentEdit.setVisible(False)
            self.editBtn.setText("✏️ 编辑")
            
            # 调整大小和位置
            parent = self.parent()
            if parent:
                self.setGeometry(parent.rect())
                self.overlay.setGeometry(parent.rect())
                
                drawer_width = int(parent.width() * 0.55)
                drawer_width = min(max(drawer_width, 600), 900)
                self.drawer.setGeometry(
                    parent.width() - drawer_width,
                    0,
                    drawer_width,
                    parent.height()
                )
            
            # 显示
            self.overlay.show()
            self.overlay.raise_()
            self.show()
            self.raise_()
            self.drawer.raise_()
        
        def toggleEdit(self):
            """切换编辑/预览模式"""
            if self.is_editing:
                # 从编辑模式切换到预览模式
                self.raw_content = self.contentEdit.toPlainText()
                html_content = self.markdown_to_html(self.raw_content)
                self.contentBrowser.setHtml(html_content)
                
                self.contentEdit.setVisible(False)
                self.contentBrowser.setVisible(True)
                self.editBtn.setText("✏️ 编辑")
                self.is_editing = False
            else:
                # 从预览模式切换到编辑模式
                self.contentEdit.setPlainText(self.raw_content)
                
                self.contentBrowser.setVisible(False)
                self.contentEdit.setVisible(True)
                self.editBtn.setText("👁 预览")
                self.is_editing = True
        
        def close(self):
            """关闭抽屉"""
            self.is_editing = False
            self.contentBrowser.setVisible(True)
            self.contentEdit.setVisible(False)
            self.hide()
            self.overlay.hide()
            self.closed.emit()
        
        def markdown_to_html(self, markdown_text):
            """将 Markdown 转换为 HTML"""
            import re
            
            html = markdown_text
            
            # 转义 HTML 特殊字符
            html = html.replace('&', '&amp;')
            html = html.replace('<', '&lt;')
            html = html.replace('>', '&gt;')
            
            # 代码块
            html = re.sub(r'```(\w*)\n(.*?)```', r'<pre><code>\2</code></pre>', html, flags=re.DOTALL)
            
            # 行内代码
            html = re.sub(r'`([^`]+)`', r'<code style="background-color: #f3f4f6; padding: 2px 6px; border-radius: 4px; font-family: Consolas, monospace; font-size: 13px;">\1</code>', html)
            
            # 标题
            html = re.sub(r'^### (.+)$', r'<h3 style="font-size: 16px; font-weight: bold; color: #1a1a1a; margin-top: 24px; margin-bottom: 12px;">\1</h3>', html, flags=re.MULTILINE)
            html = re.sub(r'^## (.+)$', r'<h2 style="font-size: 18px; font-weight: bold; color: #1a1a1a; margin-top: 28px; margin-bottom: 14px;">\1</h2>', html, flags=re.MULTILINE)
            html = re.sub(r'^# (.+)$', r'<h1 style="font-size: 22px; font-weight: bold; color: #1a1a1a; margin-top: 32px; margin-bottom: 16px;">\1</h1>', html, flags=re.MULTILINE)
            
            # 粗体
            html = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', html)
            
            # 无序列表
            html = re.sub(r'^- (.+)$', r'<li style="margin-left: 20px; margin-bottom: 8px;">\1</li>', html, flags=re.MULTILINE)
            
            # 有序列表
            html = re.sub(r'^(\d+)\. (.+)$', r'<li style="margin-left: 20px; margin-bottom: 8px;">\2</li>', html, flags=re.MULTILINE)
            
            # 水平线
            html = re.sub(r'^---$', '<hr style="border: none; border-top: 1px solid #e5e7eb; margin: 24px 0;">', html, flags=re.MULTILINE)
            
            # 段落
            html = re.sub(r'\n\n', '</p><p style="margin-bottom: 16px; line-height: 1.8;">', html)
            
            # 包装在 div 中
            html = f'''
            <div style="font-family: "Microsoft YaHei", sans-serif; color: #1f2937; line-height: 1.8;">
                <p style="margin-bottom: 16px;">{html}</p>
            </div>
            '''
            
            return html
        
        def copyContent(self):
            """复制内容"""
            if self.report_data:
                from store import read_report
                content = read_report(self.report_data.get('filepath', ''))
                QApplication.clipboard().setText(content)
                
                InfoBar.success(
                    title="复制成功",
                    content="报告内容已复制到剪贴板",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )
        
        def exportContent(self):
            """导出内容"""
            if self.report_data:
                from PyQt5.QtWidgets import QFileDialog
                file_path, _ = QFileDialog.getSaveFileName(
                    self, "导出报告",
                    f"{self.report_data.get('title', '报告')}.md",
                    "Markdown Files (*.md);;All Files (*)"
                )
                if file_path:
                    from store import read_report
                    content = read_report(self.report_data.get('filepath', ''))
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    
                    InfoBar.success(
                        title="导出成功",
                        content=f"报告已保存到: {file_path}",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
        
        def resizeEvent(self, event):
            """窗口大小变化时更新布局"""
            super().resizeEvent(event)
            parent = self.parent()
            if parent and self.isVisible():
                self.setGeometry(parent.rect())
                self.overlay.setGeometry(parent.rect())
                
                drawer_width = int(parent.width() * 0.55)
                drawer_width = min(max(drawer_width, 600), 900)
                self.drawer.setGeometry(
                    parent.width() - drawer_width,
                    0,
                    drawer_width,
                    parent.height()
                )
    
    class HistoryReportPage(QWidget):
        """历史报告页面"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setObjectName("historyReportPage")
            self.current_page = 1
            self.items_per_page = 5
            self.reports = self.loadReports()
            self.total_pages = max(1, (len(self.reports) + self.items_per_page - 1) // self.items_per_page)
            self.drawer = None
            
            # 主布局
            mainLayout = QVBoxLayout(self)
            mainLayout.setContentsMargins(16, 12, 16, 12)
            mainLayout.setSpacing(16)
            
            # ========== 顶部说明文字 ==========
            descLabel = QLabel("查看和管理所有 AI 生成的报告，支持按类型筛选和关键词搜索")
            descLabel.setStyleSheet("font-size: 11px; color: #9ca3af; border: none; background: transparent;")
            mainLayout.addWidget(descLabel)
            
            # ========== 筛选面板 ==========
            filterPanel = QFrame()
            filterPanel.setStyleSheet("""
                QFrame {
                    background-color: #f7f8fa;
                    border-radius: 10px;
                    border: none;
                }
            """)
            filterLayout = QVBoxLayout(filterPanel)
            filterLayout.setContentsMargins(16, 12, 16, 12)
            filterLayout.setSpacing(12)
            
            # 第一行：类型选择 + 搜索
            firstRowLayout = QHBoxLayout()
            firstRowLayout.setSpacing(12)
            
            # 报告类型分段按钮
            typeGroupLayout = QHBoxLayout()
            typeGroupLayout.setSpacing(0)
            typeGroupLayout.setContentsMargins(0, 0, 0, 0)
            
            self.typeButtons = []
            for i, text in enumerate(["全部", "日报", "周报", "月报"]):
                btn = QPushButton(text)
                btn.setCheckable(True)
                btn.setCursor(Qt.PointingHandCursor)
                btn.setFixedHeight(28)
                btn.setProperty("type", text)
                btn.clicked.connect(lambda checked, idx=i: self.selectType(idx))
                
                if i == 0:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: #16a34a;
                            color: white;
                            border: 1px solid #16a34a;
                            padding: 0 12px;
                            font-size: 11px;
                            font-weight: bold;
                            border-top-left-radius: 6px;
                            border-bottom-left-radius: 6px;
                            border-top-right-radius: 0px;
                            border-bottom-right-radius: 0px;
                        }
                    """)
                elif i == 3:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: white;
                            color: #374151;
                            border: 1px solid #e5e7eb;
                            padding: 0 12px;
                            font-size: 11px;
                            border-top-left-radius: 0px;
                            border-bottom-left-radius: 0px;
                            border-top-right-radius: 6px;
                            border-bottom-right-radius: 6px;
                        }
                        QPushButton:hover {
                            background-color: #f9fafb;
                        }
                    """)
                else:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: white;
                            color: #374151;
                            border: 1px solid #e5e7eb;
                            border-left: none;
                            padding: 0 12px;
                            font-size: 11px;
                            border-radius: 0px;
                        }
                        QPushButton:hover {
                            background-color: #f9fafb;
                        }
                    """)
                
                self.typeButtons.append(btn)
                typeGroupLayout.addWidget(btn)
            
            firstRowLayout.addLayout(typeGroupLayout)
            firstRowLayout.addStretch()
            
            # 搜索框 + 刷新按钮
            searchLayout = QHBoxLayout()
            searchLayout.setSpacing(6)
            
            # 使用 Fluent SearchLineEdit
            self.searchInput = SearchLineEdit()
            self.searchInput.setPlaceholderText("搜索报告...")
            self.searchInput.setFixedWidth(220)
            self.searchInput.textChanged.connect(self.onSearch)
            searchLayout.addWidget(self.searchInput)
            
            # 刷新按钮
            refreshBtn = QPushButton("🔄")
            refreshBtn.setFixedSize(28, 28)
            refreshBtn.setCursor(Qt.PointingHandCursor)
            refreshBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    border: 1px solid #e5e7eb;
                    border-radius: 6px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #f9fafb;
                }
            """)
            refreshBtn.clicked.connect(self.refreshList)
            searchLayout.addWidget(refreshBtn)
            
            firstRowLayout.addLayout(searchLayout)
            filterLayout.addLayout(firstRowLayout)
            
            # 第二行：快捷日期 + 日期范围
            secondRowLayout = QHBoxLayout()
            secondRowLayout.setSpacing(12)
            
            # 日期标签
            dateLabel = QLabel("日期")
            dateLabel.setStyleSheet("font-size: 11px; color: #6b7280; border: none; background: transparent;")
            secondRowLayout.addWidget(dateLabel)
            
            # 快捷日期分段按钮
            quickDateLayout = QHBoxLayout()
            quickDateLayout.setSpacing(0)
            
            self.quickDateButtons = []
            for i, text in enumerate(["本周", "本月", "最近7天", "最近30天"]):
                btn = QPushButton(text)
                btn.setCursor(Qt.PointingHandCursor)
                btn.setFixedHeight(28)
                btn.setProperty("range", text)
                btn.clicked.connect(lambda checked, t=text: self.selectQuickDate(t))
                
                if i == 0:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: white;
                            color: #374151;
                            border: 1px solid #e5e7eb;
                            padding: 0 10px;
                            font-size: 10px;
                            border-top-left-radius: 6px;
                            border-bottom-left-radius: 6px;
                            border-top-right-radius: 0px;
                            border-bottom-right-radius: 0px;
                        }
                        QPushButton:hover {
                            background-color: #f9fafb;
                        }
                    """)
                elif i == 3:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: white;
                            color: #374151;
                            border: 1px solid #e5e7eb;
                            border-left: none;
                            padding: 0 10px;
                            font-size: 10px;
                            border-top-left-radius: 0px;
                            border-bottom-left-radius: 0px;
                            border-top-right-radius: 6px;
                            border-bottom-right-radius: 6px;
                        }
                        QPushButton:hover {
                            background-color: #f9fafb;
                        }
                    """)
                else:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: white;
                            color: #374151;
                            border: 1px solid #e5e7eb;
                            border-left: none;
                            padding: 0 10px;
                            font-size: 10px;
                            border-radius: 0px;
                        }
                        QPushButton:hover {
                            background-color: #f9fafb;
                        }
                    """)
                
                self.quickDateButtons.append(btn)
                quickDateLayout.addWidget(btn)
            
            secondRowLayout.addLayout(quickDateLayout)
            secondRowLayout.addStretch()
            
            # 日期范围选择器（使用 Fluent CalendarPicker）
            dateRangeLayout = QHBoxLayout()
            dateRangeLayout.setSpacing(6)
            
            self.startDateEdit = CalendarPicker()
            self.startDateEdit.setDate(QDate.currentDate())
            self.startDateEdit.setDateFormat("yyyy/MM/dd")
            self.startDateEdit.setFixedWidth(110)
            self.startDateEdit.dateChanged.connect(self.validateDateRange)
            dateRangeLayout.addWidget(self.startDateEdit)
            
            toLabel = QLabel("至")
            toLabel.setStyleSheet("font-size: 10px; color: #6b7280; border: none; background: transparent;")
            dateRangeLayout.addWidget(toLabel)
            
            self.endDateEdit = CalendarPicker()
            self.endDateEdit.setDate(QDate.currentDate())
            self.endDateEdit.setDateFormat("yyyy/MM/dd")
            self.endDateEdit.setFixedWidth(110)
            self.endDateEdit.dateChanged.connect(self.validateDateRange)
            dateRangeLayout.addWidget(self.endDateEdit)
            
            secondRowLayout.addLayout(dateRangeLayout)
            filterLayout.addLayout(secondRowLayout)
            
            mainLayout.addWidget(filterPanel)
            
            # ========== 报告列表面板 ==========
            listPanel = QFrame()
            listPanel.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 10px;
                    border: 1px solid #f3f4f6;
                }
            """)
            listPanelLayout = QVBoxLayout(listPanel)
            listPanelLayout.setContentsMargins(0, 0, 0, 0)
            listPanelLayout.setSpacing(0)
            
            # 表头行
            headerLayout = QHBoxLayout()
            headerLayout.setContentsMargins(16, 12, 16, 12)
            
            headerLeftLayout = QHBoxLayout()
            headerLeftLayout.setSpacing(6)
            
            docIcon = QLabel("📄")
            docIcon.setStyleSheet("border: none; background: transparent;")
            headerLeftLayout.addWidget(docIcon)
            
            headerTitle = QLabel("报告列表")
            headerTitle.setStyleSheet("font-size: 13px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            headerLeftLayout.addWidget(headerTitle)
            
            headerLayout.addLayout(headerLeftLayout)
            headerLayout.addStretch()
            
            self.totalCountLabel = QLabel(f"共 {len(self.reports)} 份")
            self.totalCountLabel.setStyleSheet("font-size: 11px; color: #9ca3af; border: none; background: transparent;")
            headerLayout.addWidget(self.totalCountLabel)
            
            listPanelLayout.addLayout(headerLayout)
            
            # 分隔线
            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setStyleSheet("background-color: #f3f4f6; border: none; height: 1px;")
            listPanelLayout.addWidget(separator)
            
            # 报告列表（可滚动）
            self.listScrollArea = QScrollArea()
            self.listScrollArea.setWidgetResizable(True)
            self.listScrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.listScrollArea.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.listScrollArea.setStyleSheet("QScrollArea { border: none; background: transparent; } QScrollBar { width: 0px; height: 0px; }")
            
            self.listWidget = QWidget()
            self.listWidget.setStyleSheet("background: transparent; border: none;")
            self.listLayout = QVBoxLayout(self.listWidget)
            self.listLayout.setContentsMargins(0, 0, 0, 0)
            self.listLayout.setSpacing(0)
            
            self.listScrollArea.setWidget(self.listWidget)
            listPanelLayout.addWidget(self.listScrollArea, 1)
            
            # 底部分页控件
            paginationLayout = QHBoxLayout()
            paginationLayout.setContentsMargins(20, 16, 20, 16)
            paginationLayout.setSpacing(8)
            
            # 上一页按钮
            self.prevBtn = QPushButton("‹")
            self.prevBtn.setFixedSize(36, 36)
            self.prevBtn.setCursor(Qt.PointingHandCursor)
            self.prevBtn.setEnabled(False)
            self.prevBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    font-size: 18px;
                    color: #9ca3af;
                }
                QPushButton:hover:enabled {
                    background-color: #f9fafb;
                    color: #374151;
                }
                QPushButton:disabled {
                    background-color: #f9fafb;
                    color: #d1d5db;
                }
            """)
            self.prevBtn.clicked.connect(self.prevPage)
            paginationLayout.addWidget(self.prevBtn)
            
            # 页码按钮容器
            self.pageButtonsContainer = QWidget()
            self.pageButtonsContainer.setStyleSheet("background: transparent; border: none;")
            self.pageButtonsLayout = QHBoxLayout(self.pageButtonsContainer)
            self.pageButtonsLayout.setContentsMargins(0, 0, 0, 0)
            self.pageButtonsLayout.setSpacing(8)
            paginationLayout.addWidget(self.pageButtonsContainer)
            
            # 页码按钮列表
            self.pageButtons = []
            
            # 下一页按钮
            self.nextBtn = QPushButton("›")
            self.nextBtn.setFixedSize(36, 36)
            self.nextBtn.setCursor(Qt.PointingHandCursor)
            self.nextBtn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    border: 1px solid #e5e7eb;
                    border-radius: 8px;
                    font-size: 18px;
                    color: #374151;
                }
                QPushButton:hover {
                    background-color: #f9fafb;
                }
            """)
            self.nextBtn.clicked.connect(self.nextPage)
            paginationLayout.addWidget(self.nextBtn)
            
            paginationLayout.addStretch()
            
            listPanelLayout.addLayout(paginationLayout)
            
            # 初始化列表（在分页按钮创建之后）
            self.refreshList()
            
            mainLayout.addWidget(listPanel, 1)
        
        def setDrawer(self, drawer):
            """设置报告抽屉引用"""
            self.drawer = drawer
        
        def loadReports(self):
            """从文件夹加载报告列表"""
            from store import get_report_list
            return get_report_list()
        
        def selectType(self, index):
            """选择报告类型"""
            self.current_type_index = index
            for i, btn in enumerate(self.typeButtons):
                if i == index:
                    btn.setChecked(True)
                    if i == 0:
                        btn.setStyleSheet("""
                            QPushButton {
                                background-color: #16a34a;
                                color: white;
                                border: 1px solid #16a34a;
                                padding: 0 16px;
                                font-size: 13px;
                                font-weight: bold;
                                border-top-left-radius: 8px;
                                border-bottom-left-radius: 8px;
                                border-top-right-radius: 0px;
                                border-bottom-right-radius: 0px;
                            }
                        """)
                    elif i == len(self.typeButtons) - 1:
                        btn.setStyleSheet("""
                            QPushButton {
                                background-color: #16a34a;
                                color: white;
                                border: 1px solid #16a34a;
                                padding: 0 16px;
                                font-size: 13px;
                                font-weight: bold;
                                border-top-left-radius: 0px;
                                border-bottom-left-radius: 0px;
                                border-top-right-radius: 8px;
                                border-bottom-right-radius: 8px;
                            }
                        """)
                    else:
                        btn.setStyleSheet("""
                            QPushButton {
                                background-color: #16a34a;
                                color: white;
                                border: 1px solid #16a34a;
                                padding: 0 16px;
                                font-size: 13px;
                                font-weight: bold;
                                border-radius: 0px;
                            }
                        """)
                else:
                    btn.setChecked(False)
                    if i == 0:
                        btn.setStyleSheet("""
                            QPushButton {
                                background-color: white;
                                color: #374151;
                                border: 1px solid #e5e7eb;
                                padding: 0 16px;
                                font-size: 13px;
                                border-top-left-radius: 8px;
                                border-bottom-left-radius: 8px;
                                border-top-right-radius: 0px;
                                border-bottom-right-radius: 0px;
                            }
                            QPushButton:hover {
                                background-color: #f9fafb;
                            }
                        """)
                    elif i == 3:
                        btn.setStyleSheet("""
                            QPushButton {
                                background-color: white;
                                color: #374151;
                                border: 1px solid #e5e7eb;
                                padding: 0 16px;
                                font-size: 13px;
                                border-top-left-radius: 0px;
                                border-bottom-left-radius: 0px;
                                border-top-right-radius: 8px;
                                border-bottom-right-radius: 8px;
                            }
                            QPushButton:hover {
                                background-color: #f9fafb;
                            }
                        """)
                    else:
                        btn.setStyleSheet("""
                            QPushButton {
                                background-color: white;
                                color: #374151;
                                border: 1px solid #e5e7eb;
                                border-left: none;
                                padding: 0 16px;
                                font-size: 13px;
                                border-radius: 0px;
                            }
                            QPushButton:hover {
                                background-color: #f9fafb;
                            }
                        """)
            
            self.refreshList()
        
        def validateDateRange(self):
            """验证日期范围"""
            if self.startDateEdit.date > self.endDateEdit.date:
                InfoBar.warning(
                    title="日期范围错误",
                    content="开始日期不能晚于结束日期",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
        
        def selectQuickDate(self, text):
            """选择快捷日期"""
            today = QDate.currentDate()
            if text == "本周":
                monday = today.addDays(-(today.dayOfWeek() - 1))
                self.startDateEdit.setDate(monday)
                self.endDateEdit.setDate(today)
            elif text == "本月":
                first = QDate(today.year(), today.month(), 1)
                self.startDateEdit.setDate(first)
                self.endDateEdit.setDate(today)
            elif text == "最近7天":
                self.startDateEdit.setDate(today.addDays(-6))
                self.endDateEdit.setDate(today)
            elif text == "最近30天":
                self.startDateEdit.setDate(today.addDays(-29))
                self.endDateEdit.setDate(today)
        
        def onSearch(self, text):
            """搜索报告"""
            self.refreshList()
        
        def refreshList(self):
            """刷新报告列表"""
            # 重新加载报告数据
            all_reports = self.loadReports()
            
            # 按类型筛选
            type_filter = ["全部", "日报", "周报", "月报"]
            current_type = type_filter[self.current_type_index] if hasattr(self, 'current_type_index') else "全部"
            
            if current_type == "全部":
                self.reports = all_reports
            else:
                self.reports = [r for r in all_reports if r.get('type') == current_type]
            
            self.total_pages = max(1, (len(self.reports) + self.items_per_page - 1) // self.items_per_page)
            
            # 如果当前页超过总页数，重置为第一页
            if self.current_page > self.total_pages:
                self.current_page = 1
            
            # 清空现有条目
            while self.listLayout.count():
                item = self.listLayout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            
            # 计算当前页的报告范围
            start_idx = (self.current_page - 1) * self.items_per_page
            end_idx = min(start_idx + self.items_per_page, len(self.reports))
            current_reports = self.reports[start_idx:end_idx]
            
            # 添加报告条目
            for report in current_reports:
                itemWidget = ReportItemWidget(report)
                itemWidget.view_clicked.connect(self.onViewReport)
                itemWidget.copy_clicked.connect(self.onCopyReport)
                itemWidget.export_clicked.connect(self.onExportReport)
                itemWidget.delete_clicked.connect(self.onDeleteReport)
                self.listLayout.addWidget(itemWidget)
                
                # 分隔线
                separator = QFrame()
                separator.setFrameShape(QFrame.HLine)
                separator.setStyleSheet("background-color: #f3f4f6; border: none; height: 1px;")
                self.listLayout.addWidget(separator)
            
            self.listLayout.addStretch()
            
            # 更新总数
            self.totalCountLabel.setText(f"共 {len(self.reports)} 份")
            
            # 重建页码按钮
            self.rebuildPageButtons()
            
            # 更新分页状态
            self.updatePagination()
        
        def rebuildPageButtons(self):
            """重建页码按钮"""
            # 清空现有页码按钮
            for btn in self.pageButtons:
                btn.deleteLater()
            self.pageButtons.clear()
            
            # 清空容器布局
            while self.pageButtonsLayout.count():
                item = self.pageButtonsLayout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            
            # 创建新的页码按钮
            for i in range(1, self.total_pages + 1):
                btn = QPushButton(str(i))
                btn.setFixedSize(36, 36)
                btn.setCursor(Qt.PointingHandCursor)
                btn.setProperty("page", i)
                btn.clicked.connect(lambda checked, p=i: self.goToPage(p))
                self.pageButtons.append(btn)
                self.pageButtonsLayout.addWidget(btn)
        
        def goToPage(self, page):
            """跳转到指定页"""
            self.current_page = page
            self.updatePagination()
        
        def prevPage(self):
            """上一页"""
            if self.current_page > 1:
                self.current_page -= 1
                self.updatePagination()
        
        def nextPage(self):
            """下一页"""
            if self.current_page < self.total_pages:
                self.current_page += 1
                self.updatePagination()
        
        def updatePagination(self):
            """更新分页状态"""
            # 更新页码按钮样式
            for i, btn in enumerate(self.pageButtons):
                page = i + 1
                if page == self.current_page:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: #16a34a;
                            color: white;
                            border: 1px solid #16a34a;
                            border-radius: 8px;
                            font-size: 14px;
                            font-weight: bold;
                        }
                    """)
                else:
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: white;
                            color: #374151;
                            border: 1px solid #e5e7eb;
                            border-radius: 8px;
                            font-size: 14px;
                        }
                        QPushButton:hover {
                            background-color: #f9fafb;
                        }
                    """)
            
            # 更新箭头按钮状态
            self.prevBtn.setEnabled(self.current_page > 1)
            self.nextBtn.setEnabled(self.current_page < self.total_pages)
        
        def onViewReport(self, report):
            """查看报告"""
            if self.drawer:
                self.drawer.open(report)
        
        def onCopyReport(self, report):
            """复制报告"""
            QApplication.clipboard().setText(report['title'])
            InfoBar.success(
                title="复制成功",
                content="报告内容已复制到剪贴板",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
        
        def onExportReport(self, report):
            """导出报告"""
            from PyQt5.QtWidgets import QFileDialog
            from store import read_report
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, "导出报告",
                f"{report['title']}.md",
                "Markdown Files (*.md);;All Files (*)"
            )
            if file_path:
                content = read_report(report.get('filepath', ''))
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                InfoBar.success(
                    title="导出成功",
                    content=f"报告已保存到: {file_path}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
        
        def onDeleteReport(self, report):
            """删除报告"""
            reply = QMessageBox.question(
                self, "确认删除",
                f"确定要删除报告「{report['title']}」吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.reports.remove(report)
                self.refreshList()
                InfoBar.success(
                    title="删除成功",
                    content=f"报告已删除",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=2000,
                    parent=self
                )

    class LatestVersionDialog(QDialog):
        """最新版本弹窗"""
        def __init__(self, current_version, update_log, parent=None):
            super().__init__(parent)
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setModal(True)
            
            mainLayout = QHBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            
            overlay = QWidget()
            overlay.setStyleSheet("background-color: rgba(0, 0, 0, 100);")
            overlayLayout = QVBoxLayout(overlay)
            overlayLayout.setAlignment(Qt.AlignCenter)
            
            card = QFrame()
            card.setFixedSize(400, 350)
            card.setStyleSheet("QFrame { background-color: white; border-radius: 16px; border: 1px solid #ECECEC; }")
            cardLayout = QVBoxLayout(card)
            cardLayout.setContentsMargins(24, 24, 24, 24)
            cardLayout.setSpacing(16)
            
            # 图标
            iconLabel = QLabel("✅")
            iconLabel.setFixedSize(48, 48)
            iconLabel.setAlignment(Qt.AlignCenter)
            iconLabel.setStyleSheet("font-size: 32px; background: transparent; border: none;")
            cardLayout.addWidget(iconLabel, 0, Qt.AlignCenter)
            
            # 标题
            titleLabel = QLabel("已是最新版本")
            titleLabel.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            titleLabel.setAlignment(Qt.AlignCenter)
            cardLayout.addWidget(titleLabel)
            
            # 版本号
            versionLabel = QLabel(f"当前版本：{current_version}")
            versionLabel.setStyleSheet("font-size: 14px; color: #16A34A; font-weight: bold; border: none; background: transparent;")
            versionLabel.setAlignment(Qt.AlignCenter)
            cardLayout.addWidget(versionLabel)
            
            # 更新日志（支持 Markdown 渲染）
            logLabel = QLabel('')
            logLabel.setTextFormat(Qt.MarkdownText)
            logLabel.setText(update_log or '')
            logLabel.setWordWrap(True)
            logLabel.setStyleSheet("font-size: 12px; color: #666666; border: none; background: transparent;")
            logLabel.setAlignment(Qt.AlignCenter)
            cardLayout.addWidget(logLabel)
            
            cardLayout.addStretch()
            
            # 关闭按钮
            closeBtn = QPushButton("关闭")
            closeBtn.setFixedHeight(40)
            closeBtn.setCursor(Qt.PointingHandCursor)
            closeBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F44336;
                    color: white;
                    border: none;
                    border-radius: 8px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover { background-color: #E53935; }
            """)
            closeBtn.clicked.connect(self.close)
            cardLayout.addWidget(closeBtn)
            
            overlayLayout.addWidget(card)
            mainLayout.addWidget(overlay)

    class UpdateDialog(QDialog):
        """更新弹窗"""
        def __init__(self, current_version, latest_version, update_log, download_url, parent=None, force_update=False):
            super().__init__(parent)
            self.download_url = download_url
            self.force_update = force_update
            self._update_downloaded = False
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setModal(True)
            
            mainLayout = QHBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            
            overlay = QWidget()
            overlay.setStyleSheet("background-color: rgba(0, 0, 0, 100);")
            overlayLayout = QVBoxLayout(overlay)
            overlayLayout.setAlignment(Qt.AlignCenter)
            
            card = QFrame()
            card.setFixedSize(450, 400)
            card.setStyleSheet("QFrame { background-color: white; border-radius: 16px; border: 1px solid #ECECEC; }")
            cardLayout = QVBoxLayout(card)
            cardLayout.setContentsMargins(24, 24, 24, 24)
            cardLayout.setSpacing(16)
            
            # 强制更新标识
            if self.force_update:
                forceBadge = QLabel("⚡ 强制更新 · 此版本必须更新后才能继续使用")
                forceBadge.setStyleSheet("""
                    QLabel {
                        background-color: #FEF2F2;
                        color: #DC2626;
                        padding: 8px 12px;
                        border-radius: 8px;
                        font-size: 12px;
                        font-weight: bold;
                        border: 1px solid #FECACA;
                    }
                """)
                forceBadge.setAlignment(Qt.AlignCenter)
                cardLayout.addWidget(forceBadge)
            
            # 版本号标题
            versionTitle = QLabel(f"{current_version} → {latest_version}")
            versionTitle.setStyleSheet("font-size: 20px; font-weight: bold; color: #16A34A; border: none; background: transparent;")
            versionTitle.setAlignment(Qt.AlignCenter)
            cardLayout.addWidget(versionTitle)
            
            # 发现新版本
            newVersionLabel = QLabel("发现新版本！")
            newVersionLabel.setStyleSheet("font-size: 14px; color: #666666; border: none; background: transparent;")
            newVersionLabel.setAlignment(Qt.AlignCenter)
            cardLayout.addWidget(newVersionLabel)
            
            # 更新日志
            logTitle = QLabel("更新日志")
            logTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            cardLayout.addWidget(logTitle)
            
            logContent = QTextBrowser()
            logContent.setOpenExternalLinks(True)
            logContent.setMarkdown(update_log or '暂无更新内容')
            logContent.setStyleSheet("""
                QTextBrowser {
                    background-color: #F9FAFB;
                    padding: 12px;
                    border-radius: 8px;
                    font-size: 12px;
                    color: #374151;
                    border: 1px solid #E5E7EB;
                }
            """)
            cardLayout.addWidget(logContent)
            
            cardLayout.addStretch()
            
            # 按钮
            btnLayout = QHBoxLayout()
            btnLayout.setSpacing(12)
            
            # 强制更新时不显示关闭按钮，改为显示退出软件按钮
            if not self.force_update:
                closeBtn = QPushButton("关闭")
                closeBtn.setFixedHeight(40)
                closeBtn.setCursor(Qt.PointingHandCursor)
                closeBtn.setStyleSheet("""
                    QPushButton {
                        background-color: #F44336;
                        color: white;
                        border: none;
                        border-radius: 8px;
                        font-size: 14px;
                        font-weight: bold;
                    }
                    QPushButton:hover { background-color: #E53935; }
                """)
                closeBtn.clicked.connect(self.close)
                btnLayout.addWidget(closeBtn)
            else:
                exitBtn = QPushButton("退出软件")
                exitBtn.setFixedHeight(40)
                exitBtn.setCursor(Qt.PointingHandCursor)
                exitBtn.setStyleSheet("""
                    QPushButton {
                        background-color: #9CA3AF;
                        color: white;
                        border: none;
                        border-radius: 8px;
                        font-size: 14px;
                        font-weight: bold;
                    }
                    QPushButton:hover { background-color: #6B7280; }
                """)
                exitBtn.clicked.connect(self.exitApp)
                btnLayout.addWidget(exitBtn)
            
            updateBtn = QPushButton("立即更新" if self.force_update else "更新")
            updateBtn.setFixedHeight(40)
            updateBtn.setCursor(Qt.PointingHandCursor)
            updateBtn.setStyleSheet("""
                QPushButton {
                    background-color: #16A34A;
                    color: white;
                    border: none;
                    border-radius: 8px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover { background-color: #15803D; }
            """)
            updateBtn.clicked.connect(self.startUpdate)
            btnLayout.addWidget(updateBtn)
            
            cardLayout.addLayout(btnLayout)
            
            overlayLayout.addWidget(card)
            mainLayout.addWidget(overlay)
        
        def startUpdate(self):
            """开始更新"""
            import requests
            from PyQt5.QtWidgets import QFileDialog
            
            if not self.download_url:
                InfoBar.error(
                    title="下载失败",
                    content="没有可用的下载链接",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                return
            
            # 选择保存位置
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存更新文件",
                "WorkDiary_Update.exe",
                "Executable Files (*.exe);;All Files (*)"
            )
            
            if not file_path:
                return
            
            try:
                InfoBar.info(
                    title="下载中",
                    content="正在下载更新文件...",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                
                # 下载文件
                full_url = f"{API_BASE_URL}{self.download_url}"
                response = requests.get(full_url, timeout=300)
                
                with open(file_path, 'wb') as f:
                    f.write(response.content)
                
                InfoBar.success(
                    title="下载完成",
                    content=f"更新文件已保存到: {file_path}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=5000,
                    parent=self
                )
                
                # 标记已下载，允许关闭（强制更新时）
                self._update_downloaded = True
                self.close()
                
            except Exception as e:
                InfoBar.error(
                    title="下载失败",
                    content=f"下载失败: {str(e)}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
        
        def reject(self):
            """强制更新时阻止 ESC / 拒绝关闭"""
            if self.force_update:
                return
            super().reject()
        
        def exitApp(self):
            """退出整个软件"""
            from PyQt5.QtWidgets import QApplication
            self._update_downloaded = True  # 放行弹窗关闭
            self.accept()  # 结束弹窗模态循环
            app = QApplication.instance()
            if app:
                app.quit()  # 退出主程序
        
        def keyPressEvent(self, event):
            """强制更新时拦截 ESC 键"""
            if self.force_update and event.key() == Qt.Key_Escape:
                return
            super().keyPressEvent(event)
        
        def closeEvent(self, event):
            """强制更新时阻止关闭（下载完成后除外）"""
            if self.force_update and not self._update_downloaded:
                event.ignore()
                return
            super().closeEvent(event)

    # ==================== 热力图画布（QPainter 自绘，渲染干净） ====================
    
    class HeatmapCanvas(QWidget):
        """用 QPainter 精确绘制热力图，格子大小按可用宽度计算，渲染整齐"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.matrix = []        # 二维列表 [row][col]
            self.row_labels = []    # 左侧行标签 [(main, sub)] 或 [str]
            self.col_labels = []    # [(col_index, text)]
            self.col_label_pos = "bottom"  # top / bottom
            self.left_gutter = 70
            self.edge_gutter = 20
            self.spacing = 4
            self.max_cell = 26
            self.min_cell = 8
            self.setMinimumHeight(100)
        
        def setData(self, matrix, row_labels, col_labels, col_label_pos="bottom",
                    left_gutter=70, spacing=4, max_cell=26):
            self.matrix = matrix
            self.row_labels = row_labels
            self.col_labels = col_labels
            self.col_label_pos = col_label_pos
            self.left_gutter = left_gutter
            self.spacing = spacing
            self.max_cell = max_cell
            self._recalc_height()
            self.update()
        
        def _cols(self):
            return max([len(r) for r in self.matrix]) if self.matrix else 0
        
        def _rows(self):
            return len(self.matrix)
        
        def _cell_size(self):
            cols = self._cols()
            if cols <= 0:
                return self.min_cell
            avail = self.width() - self.left_gutter - 8
            cell = (avail - (cols - 1) * self.spacing) / cols
            cell = int(max(self.min_cell, min(self.max_cell, cell)))
            return cell
        
        def _recalc_height(self):
            rows = self._rows()
            cell = self._cell_size() or self.min_cell
            h = rows * cell + (rows - 1) * self.spacing
            h += self.edge_gutter + 8
            self.setMinimumHeight(max(80, h))
        
        def sizeHint(self):
            rows = self._rows()
            cell = self._cell_size() or self.min_cell
            h = rows * cell + (rows - 1) * self.spacing
            h += self.edge_gutter + 8
            return QSize(self.width() or 600, max(80, h))
        
        def resizeEvent(self, event):
            self._recalc_height()
            super().resizeEvent(event)
        
        def _green(self, intensity):
            if intensity <= 0:
                return QColor("#EBEBEB")
            levels = [QColor("#D6F5DE"), QColor("#A8EBC0"), QColor("#6FDB9B"),
                      QColor("#34C777"), QColor("#16A34A")]
            idx = min(len(levels) - 1, int(intensity * len(levels)))
            return levels[idx]
        
        def paintEvent(self, event):
            painter = QPainter(self)
            painter.setRenderHint(QPainter.Antialiasing)
            
            rows = self._rows()
            cols = self._cols()
            if rows == 0 or cols == 0:
                painter.end()
                return
            
            cell = self._cell_size()
            sp = self.spacing
            x0 = self.left_gutter
            y0 = self.edge_gutter if self.col_label_pos == "top" else 4
            
            # 最大值
            max_v = 1
            for r in self.matrix:
                for v in r:
                    max_v = max(max_v, v)
            
            # 画格子
            for ri, row in enumerate(self.matrix):
                for ci in range(cols):
                    v = row[ci] if ci < len(row) else 0
                    intensity = v / max_v if max_v > 0 else 0
                    color = self._green(intensity) if v > 0 else QColor("#EBEBEB")
                    x = x0 + ci * (cell + sp)
                    y = y0 + ri * (cell + sp)
                    painter.setPen(Qt.NoPen)
                    painter.setBrush(color)
                    painter.drawRoundedRect(x, y, cell, cell, 3, 3)
                    # 数值文字（时段模式格子较大时显示）
                    if v > 0 and cell >= 22:
                        painter.setPen(QColor("white") if intensity > 0.5 else QColor("#333333"))
                        painter.setFont(QFont("Microsoft YaHei", 8, QFont.Bold))
                        painter.drawText(QRect(x, y, cell, cell), Qt.AlignCenter, str(v))
            
            # 行标签（左侧）
            painter.setPen(QColor("#888888"))
            painter.setFont(QFont("Microsoft YaHei", 9))
            for ri in range(rows):
                y = y0 + ri * (cell + sp)
                label = self.row_labels[ri] if ri < len(self.row_labels) else ""
                if isinstance(label, tuple):
                    main, sub = label
                    painter.drawText(QRect(0, y - 2, self.left_gutter - 6, cell), Qt.AlignRight | Qt.AlignTop, main)
                    painter.setPen(QColor("#BBBBBB"))
                    painter.setFont(QFont("Microsoft YaHei", 8))
                    painter.drawText(QRect(0, y + cell // 2, self.left_gutter - 6, cell), Qt.AlignRight | Qt.AlignTop, sub)
                    painter.setPen(QColor("#888888"))
                    painter.setFont(QFont("Microsoft YaHei", 9))
                else:
                    painter.drawText(QRect(0, y, self.left_gutter - 6, cell), Qt.AlignRight | Qt.AlignVCenter, label)
            
            # 列标签
            painter.setPen(QColor("#999999"))
            painter.setFont(QFont("Microsoft YaHei", 9))
            for ci, text in self.col_labels:
                x = x0 + ci * (cell + sp)
                if self.col_label_pos == "top":
                    painter.drawText(QRect(x, 0, cell * 3 + sp * 2, self.edge_gutter - 4), Qt.AlignLeft | Qt.AlignVCenter, text)
                else:
                    y = y0 + rows * (cell + sp) + 2
                    painter.drawText(QRect(x, y, cell * 3 + sp * 2, 16), Qt.AlignLeft | Qt.AlignVCenter, text)
            
            painter.end()

    # ==================== 热力图页面 ====================
    
    class HeatmapPage(QWidget):
        """热力图页面 - 时段/年度热力图 + 导出数据表格"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setObjectName("heatmapPage")
            self.mode = "period"  # period=时段, year=年度
            self._records_cache = []
            self._build_ui()
            # 先设置模式按钮可见性，再延迟渲染（等待布局完成，保证画布宽度正确）
            self._update_mode_buttons()
            QTimer.singleShot(0, self.updateData)
        
        def _build_ui(self):
            scrollLayout = QVBoxLayout(self)
            scrollLayout.setContentsMargins(0, 0, 0, 0)
            scrollLayout.setSpacing(0)
            
            scrollArea = QScrollArea()
            scrollArea.setWidgetResizable(True)
            scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scrollArea.setStyleSheet("QScrollArea { border: none; background-color: #F5F6F7; }")
            
            contentWidget = QWidget()
            contentWidget.setStyleSheet("background-color: #F5F6F7; border: none;")
            layout = QVBoxLayout(contentWidget)
            layout.setContentsMargins(24, 20, 24, 20)
            layout.setSpacing(18)
            
            # ========== 顶部工具栏 ==========
            topBar = QHBoxLayout()
            topBar.setSpacing(10)
            
            # 热力图标签按钮（激活态）
            heatTag = QPushButton("🔥 热力图")
            heatTag.setFixedHeight(36)
            heatTag.setStyleSheet("""
                QPushButton {
                    background-color: white; color: #1a1a1a;
                    border: 1px solid #E5E7EB; border-radius: 8px;
                    font-size: 13px; font-weight: bold; padding: 0 16px;
                }
            """)
            topBar.addWidget(heatTag)
            
            # 时段/年度 切换
            self.periodBtn = QPushButton("📅 时段")
            self.yearBtn = QPushButton("📆 年度")
            for b in (self.periodBtn, self.yearBtn):
                b.setFixedHeight(36)
                b.setCursor(Qt.PointingHandCursor)
            self.periodBtn.clicked.connect(lambda: self.set_mode("period"))
            self.yearBtn.clicked.connect(lambda: self.set_mode("year"))
            topBar.addWidget(self.periodBtn)
            topBar.addWidget(self.yearBtn)
            
            topBar.addStretch()
            
            # 生成热力图按钮（导出对应的数据表格 .xlsx）
            self.exportBtn = QPushButton("🖼 生成热力图")
            self.exportBtn.setFixedHeight(36)
            self.exportBtn.setCursor(Qt.PointingHandCursor)
            self.exportBtn.setStyleSheet("""
                QPushButton {
                    background-color: #16A34A; color: white;
                    border: none; border-radius: 8px;
                    font-size: 13px; font-weight: bold; padding: 0 18px;
                }
                QPushButton:hover { background-color: #15803D; }
            """)
            self.exportBtn.clicked.connect(self.export_heatmap_table)
            topBar.addWidget(self.exportBtn)
            
            # 年度选择（年度模式，使用 Fluent CalendarPicker，取所选日期的年份）
            self.yearLabel = QLabel("年份")
            self.yearLabel.setStyleSheet("font-size: 12px; color: #666666; border: none; background: transparent;")
            self.yearEdit = CalendarPicker()
            self.yearEdit.setDateFormat("yyyy/MM/dd")
            self.yearEdit.setFixedWidth(140)
            self.yearEdit.setDate(QDate.currentDate())
            self.yearEdit.dateChanged.connect(lambda: self.updateData())
            topBar.addWidget(self.yearLabel)
            topBar.addWidget(self.yearEdit)
            
            # 日期范围（时段模式，使用 Fluent CalendarPicker）
            self.startLabel = QLabel("从")
            self.startLabel.setStyleSheet("font-size: 12px; color: #666666; border: none; background: transparent;")
            self.startEdit = CalendarPicker()
            self.startEdit.setDateFormat("yyyy/MM/dd")
            self.startEdit.setFixedWidth(140)
            self.startEdit.setDate(QDate.currentDate().addDays(-6))
            self.startEdit.dateChanged.connect(lambda: self.updateData())
            topBar.addWidget(self.startLabel)
            topBar.addWidget(self.startEdit)
            
            self.endLabel = QLabel("至")
            self.endLabel.setStyleSheet("font-size: 12px; color: #666666; border: none; background: transparent;")
            self.endEdit = CalendarPicker()
            self.endEdit.setDateFormat("yyyy/MM/dd")
            self.endEdit.setFixedWidth(140)
            self.endEdit.setDate(QDate.currentDate())
            self.endEdit.dateChanged.connect(lambda: self.updateData())
            topBar.addWidget(self.endLabel)
            topBar.addWidget(self.endEdit)
            
            layout.addLayout(topBar)
            
            # 副标题
            subtitle = QLabel("查看多时段工作热力分布，直观了解工作节奏")
            subtitle.setStyleSheet("font-size: 12px; color: #888888; border: none; background: transparent;")
            layout.addWidget(subtitle)
            
            # ========== 统计卡片 ==========
            statsCard = QFrame()
            statsCard.setStyleSheet("QFrame { background-color: white; border-radius: 16px; border: 1px solid #F0F0F0; }")
            statsLayout = QHBoxLayout(statsCard)
            statsLayout.setContentsMargins(24, 18, 24, 18)
            statsLayout.setSpacing(30)
            
            self.statCount = QLabel("0")
            self.statDuration = QLabel("0h")
            self.statActiveDays = QLabel("0")
            self.statDailyAvg = QLabel("0")
            stat_configs = [
                (self.statCount, "记录条数"),
                (self.statDuration, "专注时长"),
                (self.statActiveDays, "活跃天数"),
                (self.statDailyAvg, "日均记录"),
            ]
            for label, sub in stat_configs:
                w = QWidget()
                w.setStyleSheet("border: none; background: transparent;")
                vl = QVBoxLayout(w)
                vl.setSpacing(4)
                label.setStyleSheet("font-size: 22px; font-weight: 800; color: #1a1a1a; border: none; background: transparent;")
                vl.addWidget(label)
                sl = QLabel(sub)
                sl.setStyleSheet("font-size: 11px; color: #999999; border: none; background: transparent;")
                vl.addWidget(sl)
                statsLayout.addWidget(w)
            
            statsLayout.addStretch()
            self.statSlogan = QLabel("工作轨迹，清晰可见")
            self.statSlogan.setStyleSheet("font-size: 12px; color: #BBBBBB; border: none; background: transparent;")
            statsLayout.addWidget(self.statSlogan)
            layout.addWidget(statsCard)
            
            # ========== 热力图区域 ==========
            self.heatCard = QFrame()
            self.heatCard.setStyleSheet("QFrame { background-color: white; border-radius: 16px; border: 1px solid #F0F0F0; }")
            self.heatCardLayout = QVBoxLayout(self.heatCard)
            self.heatCardLayout.setContentsMargins(20, 18, 20, 18)
            self.heatCardLayout.setSpacing(12)
            layout.addWidget(self.heatCard)
            
            # ========== 年度概览（年度模式）==========
            self.yearOverviewCard = QFrame()
            self.yearOverviewCard.setStyleSheet("QFrame { background-color: white; border-radius: 16px; border: 1px solid #F0F0F0; }")
            self.yearOverviewLayout = QVBoxLayout(self.yearOverviewCard)
            self.yearOverviewLayout.setContentsMargins(24, 18, 24, 18)
            self.yearOverviewLayout.setSpacing(12)
            layout.addWidget(self.yearOverviewCard)
            
            layout.addStretch()
            
            scrollArea.setWidget(contentWidget)
            scrollLayout.addWidget(scrollArea)
        
        def set_mode(self, mode):
            self.mode = mode
            self._update_mode_buttons()
            self.updateData()
        
        def _update_mode_buttons(self):
            active = """
                QPushButton { background-color: #1a1a1a; color: white; border: none; border-radius: 8px; font-size: 13px; font-weight: bold; padding: 0 16px; }
            """
            inactive = """
                QPushButton { background-color: white; color: #374151; border: 1px solid #E5E7EB; border-radius: 8px; font-size: 13px; padding: 0 16px; }
                QPushButton:hover { background-color: #F9FAFB; }
            """
            if self.mode == "period":
                self.periodBtn.setStyleSheet(active)
                self.yearBtn.setStyleSheet(inactive)
                self.yearLabel.setVisible(False)
                self.yearEdit.setVisible(False)
                self.startLabel.setVisible(True)
                self.startEdit.setVisible(True)
                self.endLabel.setVisible(True)
                self.endEdit.setVisible(True)
                self.statSlogan.setText("工作轨迹，清晰可见")
                self.yearOverviewCard.setVisible(False)
            else:
                self.yearBtn.setStyleSheet(active)
                self.periodBtn.setStyleSheet(inactive)
                self.yearLabel.setVisible(True)
                self.yearEdit.setVisible(True)
                self.startLabel.setVisible(False)
                self.startEdit.setVisible(False)
                self.endLabel.setVisible(False)
                self.endEdit.setVisible(False)
                self.statSlogan.setText("让时间说话，让努力有迹可循")
                self.yearOverviewCard.setVisible(True)
        
        def _load_records(self):
            from store import read_records
            return read_records()
        
        def _green_color(self, intensity):
            # 从浅到深的绿色
            if intensity <= 0:
                return "#EBEBEB"
            levels = ["#D6F5DE", "#A8EBC0", "#6FDB9B", "#34C777", "#16A34A"]
            idx = min(len(levels) - 1, int(intensity * len(levels)))
            return levels[idx]
        
        def updateData(self):
            records = self._load_records()
            self._records_cache = records
            
            # 清空热力图区域
            while self.heatCardLayout.count():
                item = self.heatCardLayout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            
            if self.mode == "period":
                self._render_period(records)
            else:
                self._render_year(records)
        
        def _render_period(self, records):
            start = self.startEdit.date
            end = self.endEdit.date
            if end < start:
                start, end = end, start
            
            # 展示所选范围内所有天
            n_days = start.daysTo(end) + 1
            days = [start.addDays(i) for i in range(n_days)]
            
            # 统计
            day_hour = {}
            day_count = {}
            day_min = {}
            total_count = 0
            total_min = 0.0
            for r in records:
                d = QDate.fromString(r.get('日期', ''), "yyyy-MM-dd")
                if not d.isValid() or d < start or d > end:
                    continue
                t = r.get('时间', '00:00:00')
                try:
                    hour = int(t.split(':')[0])
                except:
                    hour = 0
                key = d.toString("yyyy-MM-dd")
                day_hour.setdefault(key, [0] * 24)
                day_hour[key][hour] += 1
                day_count[key] = day_count.get(key, 0) + 1
                try:
                    day_min[key] = day_min.get(key, 0) + float(r.get('持续时长(分钟)', 0))
                except:
                    pass
                total_count += 1
                try:
                    total_min += float(r.get('持续时长(分钟)', 0))
                except:
                    pass
            
            active_days = len([k for k in day_count if day_count[k] > 0])
            daily_avg = round(total_count / n_days) if n_days > 0 else 0
            
            self.statCount.setText(str(total_count))
            self.statDuration.setText(f"{total_min / 60:.1f}h")
            self.statActiveDays.setText(str(active_days))
            self.statDailyAvg.setText(str(daily_avg))
            
            # 标题
            title = QLabel("时段记录")
            title.setStyleSheet("font-size: 15px; font-weight: 800; color: #1a1a1a; border: none; background: transparent;")
            self.heatCardLayout.addWidget(title)
            
            # 构建矩阵与标签
            matrix = []
            row_labels = []
            for d in days:
                key = d.toString("yyyy-MM-dd")
                matrix.append(day_hour.get(key, [0] * 24))
                if d == QDate.currentDate():
                    dname = "今天"
                elif d == QDate.currentDate().addDays(-1):
                    dname = "昨天"
                else:
                    dname = d.toString("MM-dd")
                row_labels.append((dname, f"{day_count.get(key, 0)}条 · {day_min.get(key, 0):.0f}min"))
            
            col_labels = [(h, f"{h}:00") for h in range(0, 24, 3)]
            
            canvas = HeatmapCanvas(self)
            canvas.setData(matrix, row_labels, col_labels,
                           col_label_pos="bottom", left_gutter=70, spacing=5, max_cell=30)
            
            # 直接加入布局，画布按行数自然撑高；
            # 超过页面高度时由页面外层滚动条整体往下浏览
            self.heatCardLayout.addWidget(canvas)
        
        def _render_year(self, records):
            year = self.yearEdit.date.year()
            
            # 统计 date -> count
            date_count = {}
            total_count = 0
            total_min = 0.0
            for r in records:
                dstr = r.get('日期', '')
                if not dstr.startswith(str(year)):
                    continue
                date_count[dstr] = date_count.get(dstr, 0) + 1
                total_count += 1
                try:
                    total_min += float(r.get('持续时长(分钟)', 0))
                except:
                    pass
            
            active_days = len([k for k in date_count if date_count[k] > 0])
            daily_avg = round(total_count / active_days) if active_days > 0 else 0
            
            self.statCount.setText(str(total_count))
            self.statDuration.setText(f"{total_min / 60:.1f}h")
            self.statActiveDays.setText(str(active_days))
            self.statDailyAvg.setText(str(daily_avg))
            
            title = QLabel(f"{year} 年度记录")
            title.setStyleSheet("font-size: 15px; font-weight: 800; color: #1a1a1a; border: none; background: transparent;")
            self.heatCardLayout.addWidget(title)
            
            # 构建周x7网格
            import datetime as dt
            jan1 = dt.date(year, 1, 1)
            dec31 = dt.date(year, 12, 31)
            # 从1月1日所在周的周一开始
            start_monday = jan1 - dt.timedelta(days=jan1.weekday())
            
            # 构建周x7矩阵
            matrix = []
            d = start_monday
            week = 0
            month_labels = []
            last_month = None
            while d <= dec31:
                col = []
                for row in range(7):
                    cur = d + dt.timedelta(days=row)
                    col.append(date_count.get(cur.strftime("%Y-%m-%d"), 0) if cur.year == year else 0)
                matrix.append(col)
                if d.month != last_month and d.year == year:
                    month_labels.append((week, f"{d.month}月"))
                    last_month = d.month
                d += dt.timedelta(days=7)
                week += 1
            
            # 转置为 [7行][week列]
            matrix_t = [[matrix[w][r] for w in range(len(matrix))] for r in range(7)]
            row_labels = ["周一", "", "周三", "", "周五", "", ""]
            
            canvas = HeatmapCanvas(self)
            canvas.setData(matrix_t, row_labels, month_labels,
                           col_label_pos="top", left_gutter=40, spacing=3, max_cell=14)
            self.heatCardLayout.addWidget(canvas)
            
            # 年度概览
            self._render_year_overview(records, year, date_count, total_min)
        
        def _render_year_overview(self, records, year, date_count, total_min):
            # 清空
            while self.yearOverviewLayout.count():
                item = self.yearOverviewLayout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            
            import datetime as dt
            title = QLabel("年度概览")
            title.setStyleSheet("font-size: 15px; font-weight: 800; color: #1a1a1a; border: none; background: transparent;")
            self.yearOverviewLayout.addWidget(title)
            
            # 最活跃月份
            month_count = {}
            weekday_count = {}
            for dstr, c in date_count.items():
                try:
                    d = dt.datetime.strptime(dstr, "%Y-%m-%d").date()
                    month_count[d.month] = month_count.get(d.month, 0) + c
                    weekday_count[d.weekday()] = weekday_count.get(d.weekday(), 0) + c
                except:
                    pass
            
            if month_count:
                best_month = max(month_count, key=month_count.get)
                best_month_str = f"{best_month}月"
                best_month_count = month_count[best_month]
            else:
                best_month_str, best_month_count = "-", 0
            
            wd_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
            if weekday_count:
                best_wd = max(weekday_count, key=weekday_count.get)
                best_wd_str = wd_names[best_wd]
                best_wd_count = weekday_count[best_wd]
            else:
                best_wd_str, best_wd_count = "-", 0
            
            # 连续活跃天数
            sorted_dates = sorted([dt.datetime.strptime(k, "%Y-%m-%d").date() for k in date_count if date_count[k] > 0])
            max_streak = 0
            cur_streak = 0
            prev = None
            for d in sorted_dates:
                if prev and (d - prev).days == 1:
                    cur_streak += 1
                else:
                    cur_streak = 1
                max_streak = max(max_streak, cur_streak)
                prev = d
            
            # 峰值日期 & 记录最多一天
            if date_count:
                peak_date = max(date_count, key=date_count.get)
                peak_count = date_count[peak_date]
            else:
                peak_date, peak_count = "-", 0
            
            row = QHBoxLayout()
            row.setSpacing(30)
            configs = [
                ("最活跃月份", best_month_str, f"{best_month_count} 条记录"),
                ("最活跃星期", best_wd_str, f"{best_wd_count} 条记录"),
                ("连续活跃天数", f"{max_streak} 天", "最长连续活跃"),
                ("年度峰值日期", peak_date, f"{total_min / 60:.1f}h"),
                ("记录最多一天", peak_date, f"{peak_count} 条记录"),
            ]
            for sub, main, sub2 in configs:
                w = QWidget()
                w.setStyleSheet("border: none; background: transparent;")
                vl = QVBoxLayout(w)
                vl.setSpacing(4)
                s1 = QLabel(sub)
                s1.setStyleSheet("font-size: 11px; color: #999999; border: none; background: transparent;")
                m = QLabel(main)
                m.setStyleSheet("font-size: 16px; font-weight: 800; color: #1a1a1a; border: none; background: transparent;")
                s2 = QLabel(sub2)
                s2.setStyleSheet("font-size: 11px; color: #999999; border: none; background: transparent;")
                vl.addWidget(s1)
                vl.addWidget(m)
                vl.addWidget(s2)
                row.addWidget(w)
            row.addStretch()
            self.yearOverviewLayout.addLayout(row)
        
        # ==================== 导出表格（热力图数据 + 详细记录） ====================
        
        def export_heatmap_table(self):
            """导出热力图对应的数据表格（.xlsx，无第三方依赖，Excel/WPS 可直接打开）"""
            from PyQt5.QtWidgets import QFileDialog
            from qfluentwidgets import InfoBar, InfoBarPosition
            
            mode_name = "时段" if self.mode == "period" else "年度"
            default_name = f"热力图数据_{mode_name}_{QDate.currentDate().toString('yyyyMMdd')}.xlsx"
            file_path, _ = QFileDialog.getSaveFileName(
                self, "导出热力图数据表格", default_name, "Excel 文件 (*.xlsx)"
            )
            if not file_path:
                return
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"
            
            records = self._records_cache or self._load_records()
            try:
                if self.mode == "period":
                    sheets = self._build_period_table_sheets(records)
                else:
                    sheets = self._build_year_table_sheets(records)
                _write_xlsx(file_path, sheets)
                InfoBar.success(
                    title="导出成功",
                    content=f"数据表格已保存: {file_path}",
                    orient=Qt.Horizontal, isClosable=True,
                    position=InfoBarPosition.TOP, duration=3000, parent=self
                )
            except Exception as e:
                InfoBar.error(
                    title="导出失败",
                    content=str(e),
                    orient=Qt.Horizontal, isClosable=True,
                    position=InfoBarPosition.TOP, duration=4000, parent=self
                )
        
        def _build_period_table_sheets(self, records):
            """构建时段模式的数据表：汇总统计 / 时段分布表 / 类型统计 / 详细记录"""
            from store import WORK_TYPES
            
            start = self.startEdit.date
            end = self.endEdit.date
            if end < start:
                start, end = end, start
            n_days = start.daysTo(end) + 1
            days = [start.addDays(i) for i in range(n_days)]
            
            day_hour = {}
            day_count = {}
            day_min = {}
            day_first = {}
            day_last = {}
            type_count = {}
            type_min = {}
            total_count = 0
            total_min = 0.0
            detail_rows = []
            for r in records:
                d = QDate.fromString(r.get('日期', ''), "yyyy-MM-dd")
                if not d.isValid() or d < start or d > end:
                    continue
                t = r.get('时间', '00:00:00')
                try:
                    hour = int(t.split(':')[0])
                except Exception:
                    hour = 0
                wt = r.get('工作类型', '其他')
                try:
                    mins = float(r.get('持续时长(分钟)', 0) or 0)
                except Exception:
                    mins = 0.0
                
                key = d.toString("yyyy-MM-dd")
                day_hour.setdefault(key, [0] * 24)[hour] += 1
                day_count[key] = day_count.get(key, 0) + 1
                day_min[key] = day_min.get(key, 0) + mins
                day_first.setdefault(key, t)
                day_last[key] = t
                type_count[wt] = type_count.get(wt, 0) + 1
                type_min[wt] = type_min.get(wt, 0) + mins
                total_count += 1
                total_min += mins
                detail_rows.append([r.get('日期', ''), t, wt, r.get('工作描述', ''), round(mins, 1)])
            
            active_days = len(day_count)
            
            # 1) 汇总统计
            stats = [
                ("日期范围", f"{start.toString('yyyy-MM-dd')} 至 {end.toString('yyyy-MM-dd')}"),
                ("记录条数", total_count),
                ("专注时长(小时)", round(total_min / 60, 2)),
                ("活跃天数", active_days),
                ("日均记录", round(total_count / n_days, 1) if n_days else 0),
            ]
            if day_first:
                stats.append(("最早使用时间", sorted(day_first.values())[0]))
                stats.append(("最晚使用时间", sorted(day_last.values())[-1]))
            
            # 2) 时段分布表（日期 × 24小时）
            weekday_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
            matrix_rows = []
            for d in days:
                key = d.toString("yyyy-MM-dd")
                counts = day_hour.get(key, [0] * 24)
                if d == QDate.currentDate():
                    dname = "今天"
                elif d == QDate.currentDate().addDays(-1):
                    dname = "昨天"
                else:
                    dname = weekday_names[d.dayOfWeek() - 1]
                matrix_rows.append([key, dname] + counts + [day_count.get(key, 0), round(day_min.get(key, 0), 1)])
            
            # 3) 类型统计
            type_rows = []
            for wt in WORK_TYPES:
                if type_count.get(wt, 0) > 0 or type_min.get(wt, 0) > 0:
                    type_rows.append([wt, type_count.get(wt, 0), round(type_min.get(wt, 0), 1)])
            for wt in type_count:
                if wt not in WORK_TYPES:
                    type_rows.append([wt, type_count[wt], round(type_min.get(wt, 0), 1)])
            type_rows.sort(key=lambda x: -x[1])
            
            # 4) 详细记录（按日期、时间排序）
            detail_rows.sort(key=lambda x: (x[0], x[1]))
            
            hour_headers = [f"{h:02d}:00" for h in range(24)]
            return [
                {"name": "汇总统计", "headers": ["统计项", "数值"], "rows": stats, "widths": [22, 40]},
                {"name": "时段分布表",
                 "headers": ["日期", "星期"] + hour_headers + ["合计", "持续时长(分钟)"],
                 "rows": matrix_rows,
                 "widths": [12, 10] + [6] * 24 + [8, 14]},
                {"name": "类型统计",
                 "headers": ["工作类型", "记录条数", "持续时长(分钟)"],
                 "rows": type_rows, "widths": [14, 10, 16]},
                {"name": "详细记录",
                 "headers": ["日期", "时间", "工作类型", "工作描述", "持续时长(分钟)"],
                 "rows": detail_rows, "widths": [12, 10, 12, 60, 16]},
            ]
        
        def _build_year_table_sheets(self, records):
            """构建年度模式的数据表：汇总统计 / 年度分布表 / 月度统计 / 星期统计 / 类型统计 / 详细记录"""
            import datetime as dt
            from store import WORK_TYPES
            
            year = self.yearEdit.date.year()
            weekday_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
            
            date_count = {}
            date_min = {}
            type_count = {}
            type_min = {}
            total_count = 0
            total_min = 0.0
            detail_rows = []
            for r in records:
                dstr = r.get('日期', '')
                if not dstr.startswith(str(year)):
                    continue
                wt = r.get('工作类型', '其他')
                try:
                    mins = float(r.get('持续时长(分钟)', 0) or 0)
                except Exception:
                    mins = 0.0
                date_count[dstr] = date_count.get(dstr, 0) + 1
                date_min[dstr] = date_min.get(dstr, 0) + mins
                type_count[wt] = type_count.get(wt, 0) + 1
                type_min[wt] = type_min.get(wt, 0) + mins
                total_count += 1
                total_min += mins
                detail_rows.append([dstr, r.get('时间', ''), wt, r.get('工作描述', ''), round(mins, 1)])
            
            active_days = len(date_count)
            
            # 月度 / 星期统计
            month_count = {}
            month_min = {}
            weekday_count = {}
            weekday_min = {}
            for dstr in date_count:
                d = dt.datetime.strptime(dstr, "%Y-%m-%d").date()
                month_count[d.month] = month_count.get(d.month, 0) + date_count[dstr]
                month_min[d.month] = month_min.get(d.month, 0) + date_min[dstr]
                weekday_count[d.weekday()] = weekday_count.get(d.weekday(), 0) + date_count[dstr]
                weekday_min[d.weekday()] = weekday_min.get(d.weekday(), 0) + date_min[dstr]
            
            # 汇总统计
            best_month = max(month_count, key=month_count.get) if month_count else None
            best_wd = max(weekday_count, key=weekday_count.get) if weekday_count else None
            sorted_dates = sorted(
                dt.datetime.strptime(k, "%Y-%m-%d").date()
                for k in date_count if date_count[k] > 0
            )
            max_streak = 0
            cur = 0
            prev = None
            for d in sorted_dates:
                if prev and (d - prev).days == 1:
                    cur += 1
                else:
                    cur = 1
                max_streak = max(max_streak, cur)
                prev = d
            peak_date = max(date_count, key=date_count.get) if date_count else "-"
            
            stats = [
                ("年份", year),
                ("记录条数", total_count),
                ("专注时长(小时)", round(total_min / 60, 2)),
                ("活跃天数", active_days),
                ("日均记录(按活跃日)", round(total_count / active_days, 1) if active_days else 0),
                ("最活跃月份", f"{best_month}月" if best_month else "-"),
                ("最活跃月份记录数", month_count[best_month] if best_month else 0),
                ("最活跃星期", weekday_names[best_wd] if best_wd is not None else "-"),
                ("最活跃星期记录数", weekday_count[best_wd] if best_wd is not None else 0),
                ("连续活跃天数", max_streak),
                ("年度峰值日期", peak_date),
                ("峰值日期记录数", date_count.get(peak_date, 0)),
            ]
            
            # 年度分布表：全年每天（与年度热力图对应）
            daily_rows = []
            d = dt.date(year, 1, 1)
            while d <= dt.date(year, 12, 31):
                dstr = d.strftime("%Y-%m-%d")
                daily_rows.append([dstr, weekday_names[d.weekday()], date_count.get(dstr, 0), round(date_min.get(dstr, 0), 1)])
                d += dt.timedelta(days=1)
            
            month_rows = [[f"{m}月", month_count.get(m, 0), round(month_min.get(m, 0), 1)] for m in range(1, 13)]
            weekday_rows = [[weekday_names[i], weekday_count.get(i, 0), round(weekday_min.get(i, 0), 1)] for i in range(7)]
            
            type_rows = []
            for wt in WORK_TYPES:
                if type_count.get(wt, 0) > 0 or type_min.get(wt, 0) > 0:
                    type_rows.append([wt, type_count.get(wt, 0), round(type_min.get(wt, 0), 1)])
            for wt in type_count:
                if wt not in WORK_TYPES:
                    type_rows.append([wt, type_count[wt], round(type_min.get(wt, 0), 1)])
            type_rows.sort(key=lambda x: -x[1])
            
            detail_rows.sort(key=lambda x: (x[0], x[1]))
            
            return [
                {"name": "汇总统计", "headers": ["统计项", "数值"], "rows": stats, "widths": [22, 40]},
                {"name": "年度分布表",
                 "headers": ["日期", "星期", "记录数", "持续时长(分钟)"],
                 "rows": daily_rows, "widths": [12, 10, 10, 16]},
                {"name": "月度统计",
                 "headers": ["月份", "记录数", "持续时长(分钟)"],
                 "rows": month_rows, "widths": [10, 10, 16]},
                {"name": "星期统计",
                 "headers": ["星期", "记录数", "持续时长(分钟)"],
                 "rows": weekday_rows, "widths": [10, 10, 16]},
                {"name": "类型统计",
                 "headers": ["工作类型", "记录条数", "持续时长(分钟)"],
                 "rows": type_rows, "widths": [14, 10, 16]},
                {"name": "详细记录",
                 "headers": ["日期", "时间", "工作类型", "工作描述", "持续时长(分钟)"],
                 "rows": detail_rows, "widths": [12, 10, 12, 60, 16]},
            ]

    class SettingsPage(QWidget):
        """设置页面"""
        logout_signal = pyqtSignal()  # 退出登录信号
        
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setObjectName("settingsPage")
            self._test_worker = None
            
            scrollLayout = QVBoxLayout(self)
            scrollLayout.setContentsMargins(0, 0, 0, 0)
            
            scrollArea = QScrollArea()
            scrollArea.setWidgetResizable(True)
            scrollArea.setStyleSheet("QScrollArea { border: none; background-color: #F5F5F5; }")
            
            contentWidget = QWidget()
            contentWidget.setStyleSheet("background-color: #F5F5F5; border: none;")
            layout = QVBoxLayout(contentWidget)
            layout.setSpacing(15)
            layout.setContentsMargins(20, 15, 20, 15)
            
            # 页面标题
            title = QLabel("⚙️ 设置")
            title.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            layout.addWidget(title)
            
            # ========== 账号信息 ==========
            accountCard = QFrame()
            accountCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            accountLayout = QVBoxLayout(accountCard)
            accountLayout.setContentsMargins(20, 18, 20, 18)
            accountLayout.setSpacing(12)
            
            accountTitle = QLabel("👤 账号信息")
            accountTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            accountLayout.addWidget(accountTitle)
            
            # 邮箱信息
            emailLayout = QHBoxLayout()
            emailLabel = QLabel("邮箱:")
            emailLabel.setStyleSheet("font-size: 12px; color: #666666; border: none; background: transparent;")
            emailLayout.addWidget(emailLabel)
            
            self.emailValue = QLabel("未登录")
            self.emailValue.setStyleSheet("font-size: 12px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            emailLayout.addWidget(self.emailValue)
            emailLayout.addStretch()
            accountLayout.addLayout(emailLayout)
            
            # 分隔线
            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setStyleSheet("background-color: #F0F0F0; border: none; height: 1px;")
            accountLayout.addWidget(separator)
            
            # 退出登录按钮
            logoutBtn = QPushButton("🚪 退出登录")
            logoutBtn.setCursor(Qt.PointingHandCursor)
            logoutBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F44336;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #E53935;
                }
                QPushButton:pressed {
                    background-color: #C62828;
                }
            """)
            logoutBtn.clicked.connect(self.onLogout)
            accountLayout.addWidget(logoutBtn)
            
            layout.addWidget(accountCard)
            
            # 更新账号信息
            self.updateAccountInfo()
            
            # ========== 显示缩放设置 ==========
            scaleCard = QFrame()
            scaleCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            scaleLayout = QVBoxLayout(scaleCard)
            scaleLayout.setContentsMargins(20, 18, 20, 18)
            scaleLayout.setSpacing(10)
            
            scaleTitle = QLabel("🖥️ 显示缩放")
            scaleTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            scaleLayout.addWidget(scaleTitle)
            
            scaleInfo = QLabel("界面缩放比例（修改后需重启程序生效）")
            scaleInfo.setStyleSheet("color: #888888; font-size: 11px; border: none; background: transparent;")
            scaleLayout.addWidget(scaleInfo)
            
            scaleComboLayout = QHBoxLayout()
            scaleLabel = QLabel("缩放比例:")
            scaleLabel.setStyleSheet("font-size: 12px; color: #333333; border: none; background: transparent;")
            scaleComboLayout.addWidget(scaleLabel)
            
            self.scaleCombo = ComboBox()
            self.scaleCombo.addItems(["25%", "50%", "75%", "100%", "125%", "150%", "175%", "200%"])
            self.scaleCombo.setCurrentText(f"{SCALE_FACTOR * 100:.0f}%")
            self.scaleCombo.currentTextChanged.connect(self.onScaleChanged)
            scaleComboLayout.addWidget(self.scaleCombo)
            scaleComboLayout.addStretch()
            
            scaleLayout.addLayout(scaleComboLayout)
            
            systemScaleLabel = QLabel(f"系统缩放: {get_system_dpi_scale() * 100:.0f}%")
            systemScaleLabel.setStyleSheet("color: #999999; font-size: 10px; border: none; background: transparent;")
            scaleLayout.addWidget(systemScaleLabel)
            
            layout.addWidget(scaleCard)
            
            # ========== 识别模型设置 ==========
            modelCard = QFrame()
            modelCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            modelLayout = QVBoxLayout(modelCard)
            modelLayout.setContentsMargins(20, 18, 20, 18)
            modelLayout.setSpacing(12)
            
            modelTitle = QLabel("🧠 识别模型")
            modelTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            modelLayout.addWidget(modelTitle)
            
            # 模型选择
            modelSelectLayout = QHBoxLayout()
            modelSelectLabel = QLabel("选择模型:")
            modelSelectLabel.setStyleSheet("font-size: 12px; color: #333333; border: none; background: transparent;")
            modelSelectLayout.addWidget(modelSelectLabel)
            
            self.modelCombo = ComboBox()
            self.modelCombo.addItems(["GLM 通用模型", "自定义 Ollama"])
            self.modelCombo.currentTextChanged.connect(self.onModelTypeChanged)
            modelSelectLayout.addWidget(self.modelCombo)
            modelSelectLayout.addStretch()
            
            modelLayout.addLayout(modelSelectLayout)
            
            # GLM设置区域
            self.glmWidget = QWidget()
            self.glmWidget.setStyleSheet("border: none; background: transparent;")
            glmLayout = QVBoxLayout(self.glmWidget)
            glmLayout.setContentsMargins(0, 5, 0, 0)
            glmLayout.setSpacing(8)
            
            glmInfo = QLabel("使用 GLM 通用模型（无需本地部署）")
            glmInfo.setStyleSheet("color: #666666; font-size: 11px; border: none; background: transparent;")
            glmLayout.addWidget(glmInfo)
            
            glmTestBtnLayout = QHBoxLayout()
            self.glmTestBtn = QPushButton("🔗 测试 GLM 连接")
            self.glmTestBtn.setCursor(Qt.PointingHandCursor)
            self.glmTestBtn.setStyleSheet("""
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 12px;
                    border: none;
                }
                QPushButton:hover { background-color: #1976D2; }
                QPushButton:pressed { background-color: #1565C0; }
            """)
            self.glmTestBtn.clicked.connect(self.testGlmConnection)
            glmTestBtnLayout.addWidget(self.glmTestBtn)
            glmTestBtnLayout.addStretch()
            glmLayout.addLayout(glmTestBtnLayout)
            
            self.glmStatusLabel = QLabel("")
            self.glmStatusLabel.setStyleSheet("font-size: 11px; border: none; background: transparent;")
            glmLayout.addWidget(self.glmStatusLabel)
            
            modelLayout.addWidget(self.glmWidget)
            
            # Ollama设置区域
            self.ollamaWidget = QWidget()
            self.ollamaWidget.setStyleSheet("border: none; background: transparent;")
            ollamaLayout = QVBoxLayout(self.ollamaWidget)
            ollamaLayout.setContentsMargins(0, 5, 0, 0)
            ollamaLayout.setSpacing(10)
            
            ollamaHostLayout = QHBoxLayout()
            ollamaHostLabel = QLabel("服务器地址:")
            ollamaHostLabel.setStyleSheet("font-size: 12px; color: #333333; border: none; background: transparent;")
            ollamaHostLayout.addWidget(ollamaHostLabel)
            
            self.ollamaHostInput = QLineEdit()
            self.ollamaHostInput.setText("http://192.168.31.23:11434")
            self.ollamaHostInput.setPlaceholderText("http://192.168.31.23:11434")
            self.ollamaHostInput.setStyleSheet("""
                QLineEdit {
                    padding: 6px 10px;
                    border: 1px solid #E0E0E0;
                    border-radius: 6px;
                    font-size: 12px;
                    color: #333333;
                    background-color: white;
                }
                QLineEdit:focus { border: 1px solid #1976D2; }
            """)
            ollamaHostLayout.addWidget(self.ollamaHostInput)
            ollamaLayout.addLayout(ollamaHostLayout)
            
            ollamaModelLayout = QHBoxLayout()
            ollamaModelLabel = QLabel("模型名称:")
            ollamaModelLabel.setStyleSheet("font-size: 12px; color: #333333; border: none; background: transparent;")
            ollamaModelLayout.addWidget(ollamaModelLabel)
            
            self.ollamaModelInput = QLineEdit()
            self.ollamaModelInput.setText("minicpm-v4.6")
            self.ollamaModelInput.setPlaceholderText("minicpm-v4.6")
            self.ollamaModelInput.setStyleSheet("""
                QLineEdit {
                    padding: 6px 10px;
                    border: 1px solid #E0E0E0;
                    border-radius: 6px;
                    font-size: 12px;
                    color: #333333;
                    background-color: white;
                }
                QLineEdit:focus { border: 1px solid #1976D2; }
            """)
            ollamaModelLayout.addWidget(self.ollamaModelInput)
            ollamaLayout.addLayout(ollamaModelLayout)
            
            ollamaBtnLayout = QHBoxLayout()
            
            applyBtn = QPushButton("✅ 应用设置")
            applyBtn.setCursor(Qt.PointingHandCursor)
            applyBtn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 12px;
                    border: none;
                }
                QPushButton:hover { background-color: #43A047; }
                QPushButton:pressed { background-color: #388E3C; }
            """)
            applyBtn.clicked.connect(self.applyOllamaSettings)
            ollamaBtnLayout.addWidget(applyBtn)
            
            self.ollamaTestBtn = QPushButton("🔗 测试 Ollama 连接")
            self.ollamaTestBtn.setCursor(Qt.PointingHandCursor)
            self.ollamaTestBtn.setStyleSheet("""
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 12px;
                    border: none;
                }
                QPushButton:hover { background-color: #1976D2; }
                QPushButton:pressed { background-color: #1565C0; }
            """)
            self.ollamaTestBtn.clicked.connect(self.testOllamaConnection)
            ollamaBtnLayout.addWidget(self.ollamaTestBtn)
            ollamaBtnLayout.addStretch()
            ollamaLayout.addLayout(ollamaBtnLayout)
            
            self.ollamaStatusLabel = QLabel("")
            self.ollamaStatusLabel.setStyleSheet("font-size: 11px; border: none; background: transparent;")
            ollamaLayout.addWidget(self.ollamaStatusLabel)
            
            modelLayout.addWidget(self.ollamaWidget)
            
            layout.addWidget(modelCard)
            
            # ========== 测试模式 ==========
            testCard = QFrame()
            testCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            testLayout = QVBoxLayout(testCard)
            testLayout.setContentsMargins(20, 18, 20, 18)
            testLayout.setSpacing(10)
            
            testTitle = QLabel("🧪 测试模式")
            testTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            testLayout.addWidget(testTitle)
            
            testInfo = QLabel("启用后将保存每次截图分析的图片到 data/photo 文件夹")
            testInfo.setStyleSheet("color: #888888; font-size: 11px; border: none; background: transparent;")
            testInfo.setWordWrap(True)
            testLayout.addWidget(testInfo)
            
            # 测试模式开关
            testSwitchLayout = QHBoxLayout()
            testSwitchLabel = QLabel("启用测试模式:")
            testSwitchLabel.setStyleSheet("font-size: 12px; color: #333333; border: none; background: transparent;")
            testSwitchLayout.addWidget(testSwitchLabel)
            
            self.testSwitch = QCheckBox()
            self.testSwitch.setChecked(is_test_mode())
            self.testSwitch.stateChanged.connect(self.onTestModeChanged)
            testSwitchLayout.addWidget(self.testSwitch)
            testSwitchLayout.addStretch()
            
            testLayout.addLayout(testSwitchLayout)
            
            # 测试模式状态
            self.testStatusLabel = QLabel("")
            self.testStatusLabel.setStyleSheet("font-size: 11px; border: none; background: transparent;")
            testLayout.addWidget(self.testStatusLabel)
            self.updateTestStatus()
            
            layout.addWidget(testCard)
            
            # ========== 更新 ==========
            updateCard = QFrame()
            updateCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            updateLayout = QVBoxLayout(updateCard)
            updateLayout.setContentsMargins(20, 18, 20, 18)
            updateLayout.setSpacing(12)
            
            updateTitle = QLabel("🔄 更新")
            updateTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            updateLayout.addWidget(updateTitle)
            
            # 检查更新按钮
            checkUpdateBtn = QPushButton("检查更新")
            checkUpdateBtn.setCursor(Qt.PointingHandCursor)
            checkUpdateBtn.setStyleSheet("""
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 12px;
                    border: none;
                }
                QPushButton:hover { background-color: #1976D2; }
                QPushButton:pressed { background-color: #1565C0; }
            """)
            checkUpdateBtn.clicked.connect(self.checkUpdate)
            updateLayout.addWidget(checkUpdateBtn)
            
            layout.addWidget(updateCard)
            
            # ========== 关于 ==========
            aboutCard = QFrame()
            aboutCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            aboutLayout = QVBoxLayout(aboutCard)
            aboutLayout.setContentsMargins(20, 18, 20, 18)
            aboutLayout.setSpacing(8)
            
            aboutTitle = QLabel("ℹ️ 关于")
            aboutTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            aboutLayout.addWidget(aboutTitle)
            
            aboutText = QLabel(
                "工作日报助手 v1.2\n"
                "自动截图分析工作内容，生成工作日报。"
            )
            aboutText.setWordWrap(True)
            aboutText.setStyleSheet("color: #666666; font-size: 12px; line-height: 1.5; border: none; background: transparent;")
            aboutLayout.addWidget(aboutText)
            
            layout.addWidget(aboutCard)
            layout.addStretch()
            
            scrollArea.setWidget(contentWidget)
            scrollLayout.addWidget(scrollArea)
            
            # 初始化模型选择状态
            self.onModelTypeChanged(self.modelCombo.currentText())
        
        def checkUpdate(self, silent=False):
            """
            检查更新
            
            参数:
                silent: 是否静默检查（不显示"已是最新"提示）
            
            返回值:
                dict: {'has_update': bool, 'latest_version': str, 'update_log': str, 'download_url': str}
            """
            import requests
            
            try:
                print(f"[checkUpdate] 开始检查, silent={silent}")
                response = requests.get(
                    f"{API_BASE_URL}/api/check-update",
                    params={"current_version": "v1.2"},
                    timeout=5
                )
                
                print(f"[checkUpdate] 响应状态: {response.status_code}")
                
                if response.status_code != 200:
                    if not silent:
                        InfoBar.error(
                            title="检查失败",
                            content=f"服务器返回错误: {response.status_code}",
                            orient=Qt.Horizontal,
                            isClosable=True,
                            position=InfoBarPosition.TOP,
                            duration=3000,
                            parent=self
                        )
                    return None
                
                try:
                    result = response.json()
                    print(f"[checkUpdate] 解析结果: {result}")
                except:
                    if not silent:
                        InfoBar.error(
                            title="检查失败",
                            content="服务器返回数据格式错误",
                            orient=Qt.Horizontal,
                            isClosable=True,
                            position=InfoBarPosition.TOP,
                            duration=3000,
                            parent=self
                        )
                    return None
                
                if result.get('success'):
                    has_update = result.get('has_update', False)
                    current_version = result.get('current_version', 'v1.2')
                    latest_version = result.get('latest_version', 'v1.2')
                    update_log = result.get('update_log', '')
                    download_url = result.get('download_url', '')
                    force_update = result.get('force_update', False)
                    
                    print(f"[checkUpdate] has_update={has_update}, latest={latest_version}, force={force_update}")
                    
                    if has_update:
                        if not silent:
                            dialog = UpdateDialog(
                                current_version,
                                latest_version,
                                update_log,
                                download_url,
                                self,
                                force_update=force_update
                            )
                            dialog.exec_()
                        return result
                    else:
                        if not silent:
                            dialog = LatestVersionDialog(current_version, update_log, self)
                            dialog.exec_()
                        return result
                else:
                    if not silent:
                        InfoBar.error(
                            title="检查失败",
                            content=result.get('message', '无法获取版本信息'),
                            orient=Qt.Horizontal,
                            isClosable=True,
                            position=InfoBarPosition.TOP,
                            duration=3000,
                            parent=self
                        )
                    return None
            except Exception as e:
                print(f"[checkUpdate] 异常: {e}")
                if not silent:
                    InfoBar.error(
                        title="检查失败",
                        content=f"发生错误: {str(e)}",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                return None
        
        def onModelTypeChanged(self, text):
            """模型类型改变时的处理"""
            if text == "GLM 通用模型":
                self.glmWidget.setVisible(True)
                self.ollamaWidget.setVisible(False)
                set_use_glm(True)
            else:
                self.glmWidget.setVisible(False)
                self.ollamaWidget.setVisible(True)
                set_use_glm(False)
        
        def testGlmConnection(self):
            """测试GLM连接（多线程）"""
            self.glmTestBtn.setEnabled(False)
            self.glmTestBtn.setText("测试中...")
            self.glmStatusLabel.setText("⏳ 正在测试连接...")
            self.glmStatusLabel.setStyleSheet("color: #FF9800; font-size: 11px; border: none; background: transparent;")
            
            self._test_worker = ConnectionTestWorker("glm")
            self._test_worker.finished.connect(self.onGlmTestFinished)
            self._test_worker.start()
        
        def onGlmTestFinished(self, success, message):
            """GLM测试完成回调"""
            self.glmTestBtn.setEnabled(True)
            self.glmTestBtn.setText("🔗 测试 GLM 连接")
            if success:
                self.glmStatusLabel.setText(f"✅ {message}")
                self.glmStatusLabel.setStyleSheet("color: #4CAF50; font-size: 11px; border: none; background: transparent;")
            else:
                self.glmStatusLabel.setText(f"❌ {message}")
                self.glmStatusLabel.setStyleSheet("color: #F44336; font-size: 11px; border: none; background: transparent;")
        
        def applyOllamaSettings(self):
            """应用Ollama设置"""
            host = self.ollamaHostInput.text().strip()
            model = self.ollamaModelInput.text().strip()
            if not host:
                host = "http://192.168.31.23:11434"
            if not model:
                model = "minicpm-v4.6"
            set_ollama_config(host, model)
            InfoBar.success(
                title="设置已保存",
                content=f"Ollama 服务器: {host}\n模型: {model}",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
        
        def testOllamaConnection(self):
            """测试Ollama连接（多线程）"""
            host = self.ollamaHostInput.text().strip()
            model = self.ollamaModelInput.text().strip()
            if not host:
                host = "http://192.168.31.23:11434"
            if not model:
                model = "minicpm-v4.6"
            
            self.ollamaTestBtn.setEnabled(False)
            self.ollamaTestBtn.setText("测试中...")
            self.ollamaStatusLabel.setText("⏳ 正在测试连接...")
            self.ollamaStatusLabel.setStyleSheet("color: #FF9800; font-size: 11px; border: none; background: transparent;")
            
            self._test_worker = ConnectionTestWorker("ollama", host, model)
            self._test_worker.finished.connect(self.onOllamaTestFinished)
            self._test_worker.start()
        
        def onOllamaTestFinished(self, success, message):
            """Ollama测试完成回调"""
            self.ollamaTestBtn.setEnabled(True)
            self.ollamaTestBtn.setText("🔗 测试 Ollama 连接")
            if success:
                self.ollamaStatusLabel.setText(f"✅ {message}")
                self.ollamaStatusLabel.setStyleSheet("color: #4CAF50; font-size: 11px; border: none; background: transparent;")
            else:
                self.ollamaStatusLabel.setText(f"❌ {message}")
                self.ollamaStatusLabel.setStyleSheet("color: #F44336; font-size: 11px; border: none; background: transparent;")
        
        def onTestModeChanged(self, state):
            """测试模式开关变化"""
            enabled = state == Qt.Checked
            set_test_mode(enabled)
            self.updateTestStatus()
            
            if enabled:
                InfoBar.success(
                    title="测试模式已启用",
                    content="截图将保存到 data/photo 文件夹",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
            else:
                InfoBar.info(
                    title="测试模式已关闭",
                    content="截图将不再保存",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
        
        def updateTestStatus(self):
            """更新测试模式状态显示"""
            if is_test_mode():
                import sys
                if getattr(sys, 'frozen', False):
                    base_dir = os.path.dirname(sys.executable)
                else:
                    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                photo_dir = os.path.join(base_dir, 'data', 'photo')
                file_count = 0
                if os.path.exists(photo_dir):
                    file_count = len([f for f in os.listdir(photo_dir) if f.endswith('.png')])
                self.testStatusLabel.setText(f"✅ 已启用 - 已保存 {file_count} 张截图")
                self.testStatusLabel.setStyleSheet("color: #4CAF50; font-size: 11px; border: none; background: transparent;")
            else:
                self.testStatusLabel.setText("⏸️ 未启用")
                self.testStatusLabel.setStyleSheet("color: #999999; font-size: 11px; border: none; background: transparent;")
        
        def updateAccountInfo(self):
            """更新账号信息"""
            import json
            config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
            config_file = os.path.join(config_dir, 'login_state.json')
            
            if os.path.exists(config_file):
                try:
                    with open(config_file, 'r', encoding='utf-8') as f:
                        state = json.load(f)
                        email = state.get('email', '未登录')
                        self.emailValue.setText(email)
                except:
                    self.emailValue.setText("未登录")
            else:
                self.emailValue.setText("未登录")
        
        def onLogout(self):
            """退出登录"""
            reply = QMessageBox.question(
                self, "确认退出",
                "确定要退出登录吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                # 删除登录状态文件
                import json
                config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
                config_file = os.path.join(config_dir, 'login_state.json')
                
                if os.path.exists(config_file):
                    os.remove(config_file)
                
                # 发送退出登录信号
                self.logout_signal.emit()
        
        def onScaleChanged(self, text):
            global SCALE_FACTOR
            try:
                new_scale = int(text.replace('%', '')) / 100.0
                if new_scale != SCALE_FACTOR:
                    SCALE_FACTOR = new_scale
                    self.saveScaleSetting(new_scale)
                    InfoBar.info(
                        title="缩放设置已更改",
                        content=f"新缩放比例: {text}，请重启程序使设置生效",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=5000,
                        parent=self
                    )
            except Exception as e:
                print(f"设置缩放失败: {e}")
        
        def saveScaleSetting(self, scale):
            try:
                config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
                os.makedirs(config_dir, exist_ok=True)
                config_file = os.path.join(config_dir, 'config.txt')
                
                config = {}
                if os.path.exists(config_file):
                    with open(config_file, 'r', encoding='utf-8') as f:
                        for line in f:
                            if '=' in line:
                                key, value = line.strip().split('=', 1)
                                config[key] = value
                
                config['scale_factor'] = str(scale)
                
                with open(config_file, 'w', encoding='utf-8') as f:
                    for key, value in config.items():
                        f.write(f"{key}={value}\n")
            except Exception as e:
                print(f"保存配置失败: {e}")

    # ==================== 管理监控页面 ====================
    
    class MonitorPage(QWidget):
        """管理监控页面 - 定时自动截图分析"""
        def __init__(self, main_window, parent=None):
            super().__init__(parent)
            from datetime import datetime  # 导入datetime模块
            self.datetime = datetime  # 保存到实例变量
            self.main_window = main_window
            self.setObjectName("monitorPage")
            self.is_monitoring = False  # 监控状态标志
            
            # 主布局
            mainLayout = QVBoxLayout(self)
            mainLayout.setContentsMargins(0, 0, 0, 0)
            mainLayout.setSpacing(0)
            
            # 滚动区域
            scrollArea = QScrollArea()
            scrollArea.setWidgetResizable(True)
            scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scrollArea.setStyleSheet("QScrollArea { border: none; background-color: #F5F5F5; }")
            
            contentWidget = QWidget()
            contentWidget.setStyleSheet("background-color: #F5F5F5; border: none;")
            layout = QVBoxLayout(contentWidget)
            layout.setContentsMargins(20, 20, 20, 20)
            layout.setSpacing(20)
            
            # ========== 页面标题 ==========
            headerCard = QFrame()
            headerCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            headerLayout = QHBoxLayout(headerCard)
            headerLayout.setContentsMargins(20, 15, 20, 15)
            
            title = QLabel("⚙️ 管理监控")
            title.setStyleSheet("font-size: 18px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            headerLayout.addWidget(title)
            headerLayout.addStretch()
            
            # 监控状态标签
            self.statusBadge = QLabel("⏸️ 未监控")
            self.statusBadge.setStyleSheet("""
                QLabel {
                    background-color: #E0E0E0;
                    color: #666666;
                    padding: 6px 14px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: bold;
                    border: none;
                }
            """)
            headerLayout.addWidget(self.statusBadge)
            
            layout.addWidget(headerCard)
            
            # ========== 监控间隔设置 ==========
            intervalCard = QFrame()
            intervalCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            intervalLayout = QVBoxLayout(intervalCard)
            intervalLayout.setContentsMargins(20, 18, 20, 18)
            intervalLayout.setSpacing(12)
            
            # 标题
            intervalTitle = QLabel("⏱️ 监控间隔时长")
            intervalTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            intervalLayout.addWidget(intervalTitle)
            
            # 说明
            intervalDesc = QLabel("选择自动截图分析的时间间隔，点击开始后将在第一个间隔结束后进行首次分析")
            intervalDesc.setStyleSheet("font-size: 11px; color: #888888; border: none; background: transparent;")
            intervalDesc.setWordWrap(True)
            intervalLayout.addWidget(intervalDesc)
            
            # 间隔选择网格
            intervalGrid = QGridLayout()
            intervalGrid.setSpacing(10)
            
            # 间隔选项（分钟数，显示文本）
            self.interval_options = [
                (1, "1 分钟"), (2, "2 分钟"), (5, "5 分钟"),
                (10, "10 分钟"), (15, "15 分钟"), (20, "20 分钟"),
                (30, "30 分钟"), (60, "1 小时"), (120, "2 小时")
            ]
            
            self.interval_buttons = []
            self.selected_interval = 10  # 默认选中10分钟
            
            for i, (minutes, text) in enumerate(self.interval_options):
                btn = QPushButton(text)
                btn.setCheckable(True)
                btn.setCursor(Qt.PointingHandCursor)
                btn.setMinimumHeight(40)
                
                # 默认选中10分钟
                if minutes == 10:
                    btn.setChecked(True)
                
                # 样式
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #F5F5F5;
                        color: #333333;
                        border: 2px solid #E0E0E0;
                        border-radius: 8px;
                        font-size: 12px;
                        font-weight: bold;
                        padding: 8px 16px;
                    }
                    QPushButton:checked {
                        background-color: #E3F2FD;
                        color: #1976D2;
                        border: 2px solid #1976D2;
                    }
                    QPushButton:hover {
                        background-color: #E8F5E9;
                        border: 2px solid #4CAF50;
                    }
                """)
                
                # 点击事件
                btn.clicked.connect(lambda checked, m=minutes, b=btn: self.selectInterval(m, b))
                
                self.interval_buttons.append(btn)
                intervalGrid.addWidget(btn, i // 3, i % 3)
            
            intervalLayout.addLayout(intervalGrid)
            layout.addWidget(intervalCard)
            
            # ========== 监控日志 ==========
            logCard = QFrame()
            logCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            logLayout = QVBoxLayout(logCard)
            logLayout.setContentsMargins(20, 18, 20, 18)
            logLayout.setSpacing(10)
            
            logTitle = QLabel("📋 监控日志")
            logTitle.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333; border: none; background: transparent;")
            logLayout.addWidget(logTitle)
            
            self.logText = QLabel("等待开始监控...")
            self.logText.setStyleSheet("""
                QLabel {
                    background-color: #F5F5F5;
                    color: #666666;
                    padding: 12px;
                    border-radius: 8px;
                    font-size: 11px;
                    font-family: Consolas, monospace;
                    border: none;
                }
            """)
            self.logText.setWordWrap(True)
            self.logText.setAlignment(Qt.AlignTop)
            self.logText.setMinimumHeight(100)
            logLayout.addWidget(self.logText)
            
            layout.addWidget(logCard)
            
            # ========== 开始/结束监控按钮 ==========
            btnCard = QFrame()
            btnCard.setStyleSheet("QFrame { background-color: white; border-radius: 12px; border: none; }")
            btnLayout = QHBoxLayout(btnCard)
            btnLayout.setContentsMargins(20, 18, 20, 18)
            btnLayout.setSpacing(15)
            
            # 开始监控按钮
            self.startBtn = QPushButton("▶️ 开始监控")
            self.startBtn.setCursor(Qt.PointingHandCursor)
            self.startBtn.setMinimumHeight(50)
            self.startBtn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 12px 30px;
                    border-radius: 10px;
                    font-size: 15px;
                    font-weight: bold;
                    border: none;
                }
                QPushButton:hover { background-color: #43A047; }
                QPushButton:pressed { background-color: #388E3C; }
                QPushButton:disabled { background-color: #C8E6C9; color: #A5D6A7; }
            """)
            self.startBtn.clicked.connect(self.startMonitoring)
            btnLayout.addWidget(self.startBtn)
            
            # 结束监控按钮
            self.stopBtn = QPushButton("⏹️ 结束监控")
            self.stopBtn.setCursor(Qt.PointingHandCursor)
            self.stopBtn.setMinimumHeight(50)
            self.stopBtn.setEnabled(False)
            self.stopBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F44336;
                    color: white;
                    padding: 12px 30px;
                    border-radius: 10px;
                    font-size: 15px;
                    font-weight: bold;
                    border: none;
                }
                QPushButton:hover { background-color: #E53935; }
                QPushButton:pressed { background-color: #C62828; }
                QPushButton:disabled { background-color: #FFCDD2; color: #EF9A9A; }
            """)
            self.stopBtn.clicked.connect(self.stopMonitoring)
            btnLayout.addWidget(self.stopBtn)
            
            layout.addWidget(btnCard)
            
            # 添加弹性空间
            layout.addStretch()
            
            scrollArea.setWidget(contentWidget)
            mainLayout.addWidget(scrollArea)
        
        def selectInterval(self, minutes, btn):
            """选择监控间隔"""
            self.selected_interval = minutes
            
            # 更新按钮状态
            for b in self.interval_buttons:
                b.setChecked(b == btn)
            
            print(f"[监控设置] 间隔已选择: {minutes} 分钟")
        
        def startMonitoring(self):
            """开始监控"""
            if self.is_monitoring:
                return
            
            self.is_monitoring = True
            
            # 更新按钮状态
            self.startBtn.setEnabled(False)
            self.stopBtn.setEnabled(True)
            
            # 更新状态标签
            self.statusBadge.setText(f"🔴 监控中 (每{self.selected_interval}分钟)")
            self.statusBadge.setStyleSheet("""
                QLabel {
                    background-color: #E8F5E9;
                    color: #2E7D32;
                    padding: 6px 14px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: bold;
                    border: none;
                }
            """)
            
            # 更新日志
            self.logText.setText(f"[{get_now().strftime('%H:%M:%S')}] 监控已启动，间隔 {self.selected_interval} 分钟\n等待第一个间隔结束后开始首次分析...")
            
            # 启动定时监控
            start_monitor(
                interval_minutes=self.selected_interval,
                callback=self.onMonitorCallback
            )
            
            InfoBar.success(
                title="监控已启动",
                content=f"每 {self.selected_interval} 分钟自动截图分析",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
        
        def stopMonitoring(self):
            """停止监控"""
            if not self.is_monitoring:
                return
            
            self.is_monitoring = False
            
            # 停止定时监控
            stop_monitor()
            
            # 更新按钮状态
            self.startBtn.setEnabled(True)
            self.stopBtn.setEnabled(False)
            
            # 更新状态标签
            self.statusBadge.setText("⏸️ 未监控")
            self.statusBadge.setStyleSheet("""
                QLabel {
                    background-color: #E0E0E0;
                    color: #666666;
                    padding: 6px 14px;
                    border-radius: 12px;
                    font-size: 12px;
                    font-weight: bold;
                    border: none;
                }
            """)
            
            # 更新日志
            current_text = self.logText.text()
            self.logText.setText(current_text + f"\n[{get_now().strftime('%H:%M:%S')}] 监控已停止")
            
            InfoBar.info(
                title="监控已停止",
                content="定时截图分析已结束",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
        
        def onMonitorCallback(self, result, error):
            """监控回调函数 - 每次截图分析完成后调用"""
            if error:
                # 分析失败
                current_text = self.logText.text()
                self.logText.setText(current_text + f"\n[{get_now().strftime('%H:%M:%S')}] ❌ 分析失败: {str(error)}")
            elif result:
                # 分析成功
                work_type = result.get('type', '未知')
                description = result.get('description', '')
                current_text = self.logText.text()
                # 截断描述，避免日志过长
                short_desc = description[:50] + "..." if len(description) > 50 else description
                self.logText.setText(current_text + f"\n[{get_now().strftime('%H:%M:%S')}] ✅ [{work_type}] {short_desc}")
                
                # 更新其他页面数据
                self.main_window.todayPage.updateData()
                if hasattr(self.main_window, 'recordsPage'):
                    self.main_window.recordsPage.updateData()
                if hasattr(self.main_window, 'timelinePage'):
                    self.main_window.timelinePage.updateData()

    # ==================== 主窗口 ====================
    
    class MainWindow(FluentWindow):
        """主窗口"""
        def __init__(self):
            super().__init__()
            init_db()
            
            # 创建报告抽屉
            self.reportDrawer = ReportDrawer(self)
            
            # 创建所有页面
            self.todayPage = TodayWorkPage(self)
            self.screenshotPage = ScreenshotPage(self, self)
            self.recordsPage = RecordsPage(self)
            self.timelinePage = TimelinePage(self)
            self.monitorPage = MonitorPage(self)
            self.reportPage = ReportPage(self)
            self.historyReportPage = HistoryReportPage(self)
            self.heatmapPage = HeatmapPage(self)
            self.settingsPage = SettingsPage(self)
            
            # 设置历史报告页面的抽屉引用
            self.historyReportPage.setDrawer(self.reportDrawer)
            
            # 连接报告生成完成信号到历史报告页面刷新
            self.reportPage.report_generated.connect(self.historyReportPage.refreshList)
            
            # 连接退出登录信号
            self.settingsPage.logout_signal.connect(self.onLogout)
            
            # 添加导航项
            self.addSubInterface(self.todayPage, FluentIcon.HOME, "今日工作")
            self.addSubInterface(self.timelinePage, FluentIcon.PIE_SINGLE, "工作时间线")
            self.addSubInterface(self.reportPage, FluentIcon.DOCUMENT, "生成报告")
            self.addSubInterface(self.historyReportPage, FluentIcon.HISTORY, "历史报告")
            self.addSubInterface(self.heatmapPage, FluentIcon.CALENDAR, "热力图")
            self.addSubInterface(self.monitorPage, FluentIcon.PLAY, "管理监控")
            self.addSubInterface(self.recordsPage, FluentIcon.DOCUMENT, "工作记录（内测）")
            self.addSubInterface(self.screenshotPage, FluentIcon.CAMERA, "截图分析（内测）")
            self.addSubInterface(self.settingsPage, FluentIcon.SETTING, "设置",
                                NavigationItemPosition.BOTTOM)
            
            self.setWindowTitle("工作日报助手")
            self.resize(1000, 700)
            self.setMinimumSize(800, 600)
            
            # 初始化时加载数据
            self.todayPage.updateData()
            self.recordsPage.updateData()
            self.timelinePage.updateData()  # 更新时间线页面
            
            # 设置系统托盘
            self.setupSystemTray()
            
            # 标志是否从托盘恢复
            self._was_minimized = False
            
            # 更新检查相关
            self._update_available = False
            self._update_timer = QTimer()
            self._update_timer.timeout.connect(self.autoCheckUpdate)
            self._update_badge = None
            
            # 启动自动检查更新定时器（每10分钟）
            self._update_timer.start(10 * 60 * 1000)
            
            # 标记是否已检查过更新
            self._has_checked_update = False
            
            # 保存页面引用映射
            self._page_map = {
                'todayPage': self.todayPage,
                'timelinePage': self.timelinePage,
                'reportPage': self.reportPage,
                'historyReportPage': self.historyReportPage,
                'monitorPage': self.monitorPage,
                'recordsPage': self.recordsPage,
                'screenshotPage': self.screenshotPage,
                'settingsPage': self.settingsPage
            }
        
        def switchToPage(self, page_name):
            """切换到指定页面"""
            if page_name in self._page_map:
                page = self._page_map[page_name]
                # 使用 FluentWindow 的切换页面方法
                self.stackedWidget.setCurrentWidget(page)
        
        def onLogout(self):
            """退出登录"""
            self.hide()  # 隐藏主窗口
            # 显示登录窗口
            self.login_window = LoginWindow()
            self.login_window.login_success.connect(self.onLoginSuccess)
            self.login_window.show()
        
        def onLoginSuccess(self):
            """重新登录成功"""
            self.show()
            self.settingsPage.updateAccountInfo()
            
            # 登录成功后延迟检查更新
            if not self._has_checked_update:
                self._has_checked_update = True
                QTimer.singleShot(3000, self.checkUpdateOnStartup)
        
        def autoCheckUpdate(self):
            """自动检查更新（静默）"""
            try:
                result = self.settingsPage.checkUpdate(silent=True)
                if result and result.get('has_update'):
                    self._update_available = True
                    self.showUpdateBadge(True)
                else:
                    self._update_available = False
                    self.showUpdateBadge(False)
            except:
                pass
        
        def checkUpdateOnStartup(self):
            """启动时检查更新"""
            try:
                print("[更新检查] 开始检查更新...")
                result = self.settingsPage.checkUpdate(silent=True)
                print(f"[更新检查] 结果: {result}")
                if result and result.get('has_update'):
                    print("[更新检查] 发现新版本，弹窗提示")
                    self._update_available = True
                    self.showUpdateBadge(True)
                    # 弹窗提示
                    dialog = UpdateDialog(
                        result.get('current_version', 'v1.2'),
                        result.get('latest_version', ''),
                        result.get('update_log', ''),
                        result.get('download_url', ''),
                        self,
                        force_update=result.get('force_update', False)
                    )
                    dialog.exec_()
                else:
                    print("[更新检查] 已是最新版本")
            except Exception as e:
                print(f"[更新检查] 发生错误: {e}")
        
        def showUpdateBadge(self, show):
            """显示/隐藏发现新版本标签"""
            if show:
                if not self._update_badge:
                    self._update_badge = QLabel("发现新版本", self)
                    self._update_badge.setStyleSheet("""
                        QLabel {
                            background-color: #16A34A;
                            color: white;
                            padding: 4px 12px;
                            border-radius: 10px;
                            font-size: 12px;
                            font-weight: bold;
                        }
                    """)
                    self._update_badge.setCursor(Qt.PointingHandCursor)
                    self._update_badge.mousePressEvent = lambda e: self.settingsPage.checkUpdate()
                
                # 定位到标题栏右侧
                self._update_badge.show()
                self._update_badge.raise_()
                QTimer.singleShot(100, self._positionUpdateBadge)
            else:
                if self._update_badge:
                    self._update_badge.hide()
        
        def _positionUpdateBadge(self):
            """定位发现新版本标签到标题右侧"""
            if self._update_badge and self._update_badge.isVisible():
                try:
                    self._update_badge.adjustSize()
                    # 与标题文字垂直居中对齐
                    self._update_badge.move(175, 12)
                except:
                    pass
        
        def setupSystemTray(self):
            """设置系统托盘图标和菜单"""
            # 创建系统托盘图标
            self.trayIcon = QSystemTrayIcon(self)
            
            # 使用指定的图片作为图标
            icon_path = r"C:\Users\20057\Desktop\frog.jpg"
            if os.path.exists(icon_path):
                # 加载图片并缩放为图标大小
                pixmap = QPixmap(icon_path)
                pixmap = pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.trayIcon.setIcon(QIcon(pixmap))
            else:
                # 如果图片不存在，使用默认图标
                pixmap = QPixmap(32, 32)
                pixmap.fill(Qt.transparent)
                painter = QPainter(pixmap)
                painter.setRenderHint(QPainter.Antialiasing)
                painter.setBrush(QBrush(QColor("#4CAF50")))
                painter.setPen(QPen(QColor("#388E3C"), 2))
                painter.drawEllipse(2, 2, 28, 28)
                painter.setPen(QColor(Qt.white))
                painter.setFont(QFont("Arial", 14, QFont.Bold))
                painter.drawText(pixmap.rect(), Qt.AlignCenter, "W")
                painter.end()
                self.trayIcon.setIcon(QIcon(pixmap))
            
            # 创建托盘菜单
            trayMenu = QMenu()
            
            # 显示主窗口动作
            showAction = QAction("显示主窗口", self)
            showAction.triggered.connect(self.showMainWindow)
            trayMenu.addAction(showAction)
            
            trayMenu.addSeparator()
            
            # 退出动作
            quitAction = QAction("退出", self)
            quitAction.triggered.connect(self.quitApplication)
            trayMenu.addAction(quitAction)
            
            # 设置托盘菜单
            self.trayIcon.setContextMenu(trayMenu)
            
            # 双击托盘图标显示主窗口
            self.trayIcon.activated.connect(self.trayIconActivated)
            
            # 立即显示托盘图标（程序启动时就在任务栏常驻）
            self.trayIcon.show()
        
        def trayIconActivated(self, reason):
            """处理托盘图标激活事件"""
            if reason == QSystemTrayIcon.DoubleClick:
                self.showMainWindow()
        
        def showMainWindow(self):
            """显示主窗口"""
            self.showNormal()
            self.activateWindow()
            self.raise_()
        
        def quitApplication(self):
            """退出应用程序"""
            # 停止监控
            stop_monitor()
            # 退出应用
            QApplication.quit()
        
        def closeEvent(self, event):
            """处理关闭事件 - 弹窗提示用户选择最小化或关闭"""
            # 忽略默认关闭事件
            event.ignore()
            
            # 创建自定义 Fluent 风格对话框
            dialog = QDialog(self)
            dialog.setWindowTitle(" ")
            dialog.setFixedSize(380, 220)
            dialog.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            dialog.setAttribute(Qt.WA_TranslucentBackground)
            
            # 主容器
            mainWidget = QWidget(dialog)
            mainWidget.setGeometry(10, 10, 360, 200)
            mainWidget.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border-radius: 12px;
                    border: 1px solid #E0E0E0;
                }
            """)
            
            # 添加阴影效果
            shadow = QGraphicsDropShadowEffect(dialog)
            shadow.setBlurRadius(20)
            shadow.setXOffset(0)
            shadow.setYOffset(4)
            shadow.setColor(QColor(0, 0, 0, 50))
            mainWidget.setGraphicsEffect(shadow)
            
            # 布局
            layout = QVBoxLayout(mainWidget)
            layout.setContentsMargins(24, 20, 24, 20)
            layout.setSpacing(16)
            
            # 标题
            titleLabel = QLabel("👋 确认操作", mainWidget)
            titleLabel.setStyleSheet("font-size: 16px; font-weight: bold; color: #1a1a1a; border: none; background: transparent;")
            layout.addWidget(titleLabel)
            
            # 描述文本
            descLabel = QLabel("请选择您要执行的操作：", mainWidget)
            descLabel.setStyleSheet("font-size: 12px; color: #666666; border: none; background: transparent;")
            layout.addWidget(descLabel)
            
            # 按钮容器
            btnLayout = QHBoxLayout()
            btnLayout.setSpacing(12)
            
            # 最小化到任务栏按钮
            minimizeBtn = QPushButton("🌙 最小化到任务栏", mainWidget)
            minimizeBtn.setCursor(Qt.PointingHandCursor)
            minimizeBtn.setMinimumHeight(40)
            minimizeBtn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #43A047;
                }
                QPushButton:pressed {
                    background-color: #388E3C;
                }
            """)
            minimizeBtn.clicked.connect(lambda: dialog.done(1))
            btnLayout.addWidget(minimizeBtn)
            
            # 直接退出按钮
            closeBtn = QPushButton("🚪 直接退出", mainWidget)
            closeBtn.setCursor(Qt.PointingHandCursor)
            closeBtn.setMinimumHeight(40)
            closeBtn.setStyleSheet("""
                QPushButton {
                    background-color: #F44336;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 8px;
                    font-size: 13px;
                    font-weight: bold;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #E53935;
                }
                QPushButton:pressed {
                    background-color: #C62828;
                }
            """)
            closeBtn.clicked.connect(lambda: dialog.done(2))
            btnLayout.addWidget(closeBtn)
            
            layout.addLayout(btnLayout)
            
            # 取消按钮（文字按钮）
            cancelBtn = QPushButton("取消", mainWidget)
            cancelBtn.setCursor(Qt.PointingHandCursor)
            cancelBtn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #999999;
                    padding: 4px 8px;
                    border: none;
                    font-size: 11px;
                }
                QPushButton:hover {
                    color: #666666;
                }
            """)
            cancelBtn.clicked.connect(lambda: dialog.done(0))
            layout.addWidget(cancelBtn, 0, Qt.AlignCenter)
            
            # 显示对话框并获取结果
            result = dialog.exec_()
            
            # 根据用户选择执行操作
            if result == 1:
                # 最小化到系统托盘
                self.hide()
                self.trayIcon.show()
                self.trayIcon.showMessage(
                    "工作日报助手",
                    "程序已最小化到任务栏，双击图标可恢复窗口",
                    QSystemTrayIcon.Information,
                    2000
                )
            elif result == 2:
                # 直接退出
                self.quitApplication()
            # 如果点击取消，不做任何操作
        
        def changeEvent(self, event):
            """处理窗口状态变化事件"""
            if event.type() == event.WindowStateChange:
                # 如果窗口从最小化恢复
                if self.windowState() == Qt.WindowNoState:
                    self.show()
                    self.activateWindow()
                    self.raise_()
            super().changeEvent(event)
        
        def event(self, event):
            """处理事件，确保窗口能正常恢复"""
            if event.type() == event.WindowStateChange:
                if self.isMinimized():
                    # 最小化时记录状态
                    self._was_minimized = True
                elif self._was_minimized:
                    # 从最小化恢复时
                    self._was_minimized = False
                    self.showNormal()
                    self.activateWindow()
                    self.raise_()
            return super().event(event)

    # ==================== 启动应用 ====================
    
    # 设置应用程序在关闭最后一个窗口时不退出（用于最小化到托盘）
    app.setQuitOnLastWindowClosed(False)
    
    # 创建主窗口
    window = MainWindow()
    
    # 显示登录窗口
    login_window = LoginWindow()
    
    def on_login_success():
        """登录成功后显示主窗口"""
        login_window.close()
        window.show()
        window.onLoginSuccess()
    
    login_window.login_success.connect(on_login_success)
    login_window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
