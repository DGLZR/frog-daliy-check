# 工作日报助手

> 自动截图分析工作内容，智能生成工作日报，告别加班写周报。

## 项目特点

- **智能识别**：基于 AI 视觉模型（GLM / Ollama）自动识别屏幕工作内容
- **隐私保护**：自动脱敏处理，不记录联系人、账号、密码等敏感信息
- **本地存储**：所有数据仅存本地 CSV，不上传云端
- **实时统计**：24 小时热力图展示工作分布，自动计算专注时长
- **定时监控**：支持自定义间隔自动截图分析
- **现代界面**：基于 PyQt6-Fluent-Widgets 的 Fluent Design 风格

## 技术栈

| 模块 | 技术 |
|------|------|
| 语言 | Python 3.10 |
| GUI | PyQt6 + PyQt6-Fluent-Widgets |
| AI 模型 | GLM-4.6V-Flash / Ollama (MiniCPM-V) |
| 图像处理 | OpenCV + mss + numpy |
| 数据存储 | CSV |
| 打包 | PyInstaller |

## 项目结构

```
├── main.py              # 入口文件（空）
├── screenshot.py        # 截图识别核心模块
├── store.py             # 数据存储模块
├── requirements.txt     # 依赖列表
├── UI/
│   ├── main_fluent.py   # Fluent Design 主界面
│   └── main.py          # 旧版 UI
└── data/
    ├── config.txt       # 配置文件
    ├── daily_summary.csv # 每日汇总
    └── records.csv      # 详细记录
```

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python UI/main_fluent.py
```

## 工作类型

支持 14 种工作类型自动分类：

`开发` `沟通` `生活` `学习` `设计` `管理` `文档` `娱乐` `产品` `会议` `运维` `测试` `数据分析` `其他`

---

## TODO

- [ ] 
- [ ] 
- [ ] 
- [ ] 
- [ ] 
