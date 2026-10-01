# 工作日报助手 - Fluent Design 版本

使用 PyQt6-Fluent-Widgets 库实现的现代化界面。

## 运行方式

```bash
# 双击 run.bat
# 或手动运行：
conda activate daily-AI
python UI/main_fluent.py
```

## 界面特性

- **Fluent Design 风格** - 微软流畅设计语言
- **导航栏** - 左侧导航，支持图标和文字
- **卡片组件** - 圆角阴影卡片
- **表格组件** - 现代化表格样式
- **响应式布局** - 自适应窗口大小

## 页面说明

1. **仪表盘** - 显示今日概览和最近记录
2. **截图分析** - 执行截图识别
3. **工作记录** - 查看今日记录列表
4. **数据统计** - 查看详细统计
5. **设置** - 系统设置信息

## 依赖库

- PyQt6
- PyQt6-Fluent-Widgets
- opencv-python
- numpy
- mss
- ollama

## 安装依赖

```bash
pip install PyQt6-Fluent-Widgets opencv-python numpy mss ollama
```
