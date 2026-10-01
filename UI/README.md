# 工作日报助手 - UI界面

## 文件说明

- `main_window.ui` - Qt Designer界面文件
- `styles.qss` - QSS样式表
- `main.py` - 主程序文件

## 运行方式

### 方式一：使用批处理文件（推荐）

1. 双击 `run.bat` 文件
2. 程序会自动激活conda环境并运行

### 方式二：手动运行

1. 打开命令行
2. 激活conda环境：
   ```bash
   conda activate daily-AI
   ```
3. 运行程序：
   ```bash
   python UI/main.py
   ```

## 界面功能

### 左侧导航栏
- **仪表盘** - 显示今日工作概览
- **截图分析** - 执行截图识别
- **工作记录** - 查看今日记录列表
- **数据统计** - 查看详细统计
- **设置** - 系统设置

### 右侧内容区
- 根据选择的导航项显示对应内容
- 支持页面切换动画

## 自定义界面

### 修改UI布局
1. 使用Qt Designer打开 `main_window.ui`
2. 调整控件位置和大小
3. 保存文件

### 修改样式
1. 编辑 `styles.qss` 文件
2. 修改颜色、字体、间距等
3. 重新运行程序查看效果

## 依赖库

- PyQt6
- opencv-python
- numpy
- mss
- ollama

## 安装依赖

```bash
pip install PyQt6 opencv-python numpy mss ollama
```
