"""
工作日报助手 - 主程序

功能：
1. 加载UI界面
2. 应用QSS样式
3. 实现页面切换
4. 导航栏折叠动画
"""

import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QListWidgetItem, 
                             QGraphicsDropShadowEffect, QMessageBox,
                             QTableWidgetItem)
from PyQt6.QtCore import Qt, QSize, QPropertyAnimation, QEasingCurve, QTimer, QParallelAnimationGroup
from PyQt6.QtGui import QColor, QIcon, QFont
from PyQt6 import uic

# 导入自定义模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from store import init_db, get_daily_summary, get_daily_records, print_today_summary
from screenshot import run_and_store


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        super().__init__()
        
        # 加载UI文件
        ui_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'main_window.ui')
        uic.loadUi(ui_path, self)
        
        # 加载样式表
        self.load_stylesheet()
        
        # 初始化界面
        self.init_ui()
        
        # 初始化数据库
        init_db()
        
        # 更新仪表盘数据
        self.update_dashboard()
        
        # 导航栏状态
        self.nav_visible = True
        self.nav_expanded_width = 300  # 展开宽度
        self.nav_collapsed_width = 40  # 收起宽度（只显示按钮）
    
    def load_stylesheet(self):
        """加载QSS样式表"""
        style_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'styles.qss')
        with open(style_path, 'r', encoding='utf-8') as f:
            self.setStyleSheet(f.read())
    
    def init_ui(self):
        """初始化界面"""
        # 设置窗口标题和大小
        self.setWindowTitle("工作日报助手")
        self.setMinimumSize(1000, 700)
        
        # 设置导航列表样式
        self.navList.setIconSize(QSize(24, 24))
        
        # 连接信号和槽
        self.navList.currentRowChanged.connect(self.switch_page)
        self.captureBtn.clicked.connect(self.start_capture)
        self.toggleBtn.clicked.connect(self.toggle_nav)
        
        # 默认选中第一项
        self.navList.setCurrentRow(0)
        
        # 添加阴影效果
        self.add_shadow_effects()
    
    def add_shadow_effects(self):
        """添加阴影效果"""
        for card in [self.statCard1, self.statCard2, self.statCard3]:
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(20)
            shadow.setXOffset(0)
            shadow.setYOffset(2)
            shadow.setColor(QColor(0, 0, 0, 50))
            card.setGraphicsEffect(shadow)
    
    def toggle_nav(self):
        """切换导航栏显示/隐藏"""
        # 创建动画组
        animation_group = QParallelAnimationGroup()
        
        # 左侧面板动画
        nav_anim = QPropertyAnimation(self.leftPanel, b"maximumWidth")
        nav_anim.setDuration(300)
        nav_anim.setEasingCurve(QEasingCurve.Type.InOutQuad)
        
        if self.nav_visible:
            # 隐藏导航栏（保留按钮宽度）
            nav_anim.setStartValue(self.nav_expanded_width)
            nav_anim.setEndValue(self.nav_collapsed_width)
            self.toggleBtn.setText("☰")
            # 隐藏文字和列表
            self.logoLabel.hide()
            self.navList.hide()
        else:
            # 显示导航栏
            nav_anim.setStartValue(self.nav_collapsed_width)
            nav_anim.setEndValue(self.nav_expanded_width)
            self.toggleBtn.setText("✕")
            # 显示文字和列表
            self.logoLabel.show()
            self.navList.show()
        
        animation_group.addAnimation(nav_anim)
        
        # 启动动画
        animation_group.start()
        
        # 保存动画引用，防止被垃圾回收
        self._animation_group = animation_group
        
        # 切换状态
        self.nav_visible = not self.nav_visible
    
    def switch_page(self, index):
        """切换页面"""
        self.contentStack.setCurrentIndex(index)
    
    def start_capture(self):
        """开始截图分析"""
        self.captureBtn.setEnabled(False)
        self.captureBtn.setText("分析中...")
        QTimer.singleShot(100, self.do_capture)
    
    def do_capture(self):
        """执行截图分析"""
        try:
            result = run_and_store()
            QMessageBox.information(
                self, 
                "分析完成", 
                f"工作类型: {result['type']}\n工作描述: {result['description']}"
            )
            self.update_dashboard()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"截图分析失败: {str(e)}")
        finally:
            self.captureBtn.setEnabled(True)
            self.captureBtn.setText("开始截图分析")
    
    def update_dashboard(self):
        """更新仪表盘数据"""
        summary = get_daily_summary()
        
        if summary:
            self.statValue1.setText(f"{summary['记录条数']} 条")
            self.statValue2.setText(f"{summary['使用时长(小时)']} 小时")
            self.statValue3.setText(summary['主要工作'])
        else:
            self.statValue1.setText("0 条")
            self.statValue2.setText("0 小时")
            self.statValue3.setText("暂无")
        
        self.update_records_table()
    
    def update_records_table(self):
        """更新记录表格"""
        records = get_daily_records()
        self.recordsTable.setRowCount(len(records))
        
        for row, record in enumerate(records):
            self.recordsTable.setItem(row, 0, QTableWidgetItem(record['时间']))
            self.recordsTable.setItem(row, 1, QTableWidgetItem(record['工作类型']))
            self.recordsTable.setItem(row, 2, QTableWidgetItem(record['工作描述']))
        
        self.recordsTable.horizontalHeader().setStretchLastSection(True)
        self.recordsTable.setColumnWidth(0, 100)
        self.recordsTable.setColumnWidth(1, 80)


def main():
    """主函数"""
    app = QApplication(sys.argv)
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
