# 工作日报助手 - 项目概览

## 技术栈
- 语言: Python
- GUI: PyQt6 + PyQt6-Fluent-Widgets (Fluent Design)
- AI识别: GLM (zhipuai) / Ollama (minicpm-v4.6)
- 图像处理: OpenCV, mss, numpy
- 数据存储: CSV文件

## 项目结构
- main.py: 空文件（入口未使用）
- screenshot.py: 截图识别核心模块（916行）
  - capture_screen(): 截图函数
  - glm_recognize(): GLM API识别
  - ollama_recognize(): Ollama本地识别
  - run_and_store(): 截图识别并存储
  - start_monitor()/stop_monitor(): 定时监控
- store.py: 数据存储模块（383行）
  - init_db(): 初始化CSV文件
  - read_summary()/write_summary(): 汇总数据读写
  - read_records()/write_records(): 记录数据读写
- UI/main_fluent.py: Fluent Design主界面（2000+行）
  - TodayWorkPage: 今日工作页面
  - ScreenshotPage: 截图分析页面
  - RecordsPage: 工作记录页面
  - TimelinePage: 工作时间线页面
- UI/main.py: 旧版UI（194行）
- data/: 数据目录
  - config.txt: 配置文件（缩放因子）
  - daily_summary.csv: 每日汇总
  - records.csv: 详细记录

## 核心功能
1. 屏幕截图自动识别工作类型（14种分类）
2. 隐私保护：自动脱敏处理
3. 工作时长统计
4. 热力图展示24小时工作分布
5. 定时监控模式
6. 显示器信息获取

## 工作类型（14种）
开发、沟通、生活、学习、设计、管理、文档、娱乐、产品、会议、运维、测试、数据分析、其他

## 依赖包
- PyQt6>=6.4.0
- PyQt6-Frameless-Window>=0.8.0
- PyQt6-Fluent-Widgets>=1.10.0
- opencv-python>=4.6.0
- numpy>=1.24.0
- Pillow>=9.0.0
- mss>=9.0.0
- ollama>=0.4.0
- darkdetect>=0.8.0
- zai (可选，用于GLM)
