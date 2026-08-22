"""
工作日报助手 - 数据存储模块

功能说明：
1. 使用CSV文件存储工作数据，方便直接用Excel查看
2. 记录每日截图分析次数、使用时长、主要工作等汇总信息
3. 记录每一条截图分析的详细数据
4. 支持按日期查询汇总和记录

文件结构：
- data/daily_summary.csv: 每日工作汇总数据
- data/records.csv: 每条截图分析记录
"""

import csv
import os
from datetime import datetime, timedelta, timezone
import time

# ==================== 时区工具 ====================

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


# ==================== 配置常量 ====================

# 数据库文件夹路径：打包后在exe同级目录，开发时在脚本所在目录
import sys
if getattr(sys, 'frozen', False):
    # 打包后的exe环境
    DB_DIR = os.path.join(os.path.dirname(sys.executable), 'data')
else:
    # 开发环境
    DB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')

# 每日汇总CSV文件路径
SUMMARY_FILE = os.path.join(DB_DIR, 'daily_summary.csv')

# 每条记录CSV文件路径
RECORDS_FILE = os.path.join(DB_DIR, 'records.csv')

# 工作类型列表，与screenshot.py中定义的类型保持一致
WORK_TYPES = ["开发", "沟通", "生活", "学习", "设计", "管理", "文档", "娱乐", "产品", "会议", "运维", "测试", "数据分析", "其他"]

# 活动记录表头（新增「消耗token数」列，用于记录每次识别消耗的 token）
RECORD_HEADERS = ['ID', '日期', '时间', '工作类型', '工作描述', '持续时长(分钟)', '消耗token数']


def format_token_count(n):
    """
    将 token 数量格式化为带量级单位的字符串

    规则：
        >= 10亿  -> xB（十亿）
        >= 100万 -> xM（百万）
        >= 1000  -> xK（千）
        否则原样显示

    参数：
        n: token 数量（数值或字符串）
    返回值：
        格式化后的字符串，如 "12.5K"、"3.2M"、"1.1B"、"860"
    """
    try:
        n = float(n)
    except (TypeError, ValueError):
        return "0"
    if n >= 1_000_000_000:
        return f"{n / 1_000_000_000:.1f}B"
    if n >= 1_000_000:
        return f"{n / 1_000_000:.1f}M"
    if n >= 1_000:
        return f"{n / 1_000:.1f}K"
    return str(int(n))


def _to_int(value, default=0):
    """安全地把字符串/数值转成 int"""
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def _parse_meta_value(line):
    """
    解析报告元信息行（形如 "**标签：** 值"）中的值

    由于 Markdown 加粗的右 "**" 紧跟在全角冒号之后，直接按 "：" 切分
    会得到 "** 值" 这样带前缀的结果，这里统一去掉前导 "*" 和空白。
    """
    value = line.split('：', 1)[-1].strip()
    return value.lstrip('*').strip()


# ==================== 初始化函数 ====================

def init_db():
    """
    初始化数据库
    
    功能：
    1. 检查data文件夹是否存在，不存在则创建
    2. 检查daily_summary.csv是否存在，不存在则创建并写入表头
    3. 检查records.csv是否存在，不存在则创建并写入表头
    
    参数：无
    返回值：无
    """
    # 创建data文件夹（如果不存在）
    if not os.path.exists(DB_DIR):
        os.makedirs(DB_DIR)
        print(f"已创建data文件夹: {DB_DIR}")
    
    # 创建每日汇总CSV文件（如果不存在）
    if not os.path.exists(SUMMARY_FILE):
        # 定义汇总表的列名
        headers = ['日期', '记录条数', '使用时长(小时)', '主要工作', '最早使用时间', '最晚使用时间']
        # 为每种工作类型添加时长列
        headers += [f'{t}时长(小时)' for t in WORK_TYPES]
        # 为每个小时添加记录条数列（00:00-23:00）
        headers += [f'{h:02d}:00记录数' for h in range(24)]
        
        # 使用utf-8-sig编码确保Excel打开中文不乱码
        with open(SUMMARY_FILE, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers)  # 写入表头
        print(f"已创建汇总文件: {SUMMARY_FILE}")
    
    # 创建记录CSV文件（如果不存在）
    if not os.path.exists(RECORDS_FILE):
        # 定义记录表的列名
        headers = RECORD_HEADERS
        
        with open(RECORDS_FILE, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers)  # 写入表头
        print(f"已创建记录文件: {RECORDS_FILE}")
    
    # 初始化报告模板数据库
    init_templates_db()


# ==================== 读取函数 ====================

def read_summary():
    """
    读取每日汇总数据
    
    功能：从daily_summary.csv读取所有日期的汇总数据
    
    参数：无
    返回值：字典，key为日期字符串，value为该日期的汇总数据字典
           例如：{'2025-01-15': {'日期': '2025-01-15', '记录条数': '5', ...}}
    """
    # 如果文件不存在，返回空字典
    if not os.path.exists(SUMMARY_FILE):
        return {}
    
    summaries = {}
    # 使用DictReader读取，每行会自动转换为字典
    with open(SUMMARY_FILE, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 以日期为key存储，方便按日期查找
            summaries[row['日期']] = row
    return summaries


def read_records():
    """
    读取所有记录
    
    功能：从records.csv读取所有截图分析记录
    
    参数：无
    返回值：列表，每个元素是一条记录的字典
           例如：[{'ID': '1', '日期': '2025-01-15', '时间': '09:30:00', ...}]
    """
    # 如果文件不存在，返回空列表
    if not os.path.exists(RECORDS_FILE):
        return []
    
    records = []
    with open(RECORDS_FILE, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            records.append(row)
    return records


# ==================== 写入函数 ====================

def write_summary(summaries):
    """
    写入每日汇总数据
    
    功能：将汇总数据字典写入daily_summary.csv（覆盖写入）
    
    参数：
        summaries: 字典，key为日期，value为汇总数据字典
    返回值：无
    """
    # 确保目录存在
    os.makedirs(DB_DIR, exist_ok=True)
    
    # 定义表头
    headers = ['日期', '记录条数', '使用时长(小时)', '主要工作', '最早使用时间', '最晚使用时间']
    headers += [f'{t}时长(小时)' for t in WORK_TYPES]
    # 为每个小时添加记录条数列（00:00-23:00）
    headers += [f'{h:02d}:00记录数' for h in range(24)]
    
    # 覆盖写入整个文件
    with open(SUMMARY_FILE, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()  # 写入表头
        # 按日期排序写入，保持文件有序
        for date in sorted(summaries.keys()):
            writer.writerow(summaries[date])


def write_records(records):
    """
    写入记录数据
    
    功能：将记录列表写入records.csv（覆盖写入）
    
    参数：
        records: 列表，每个元素是一条记录的字典
    返回值：无
    """
    # 确保目录存在
    os.makedirs(DB_DIR, exist_ok=True)
    
    # 定义表头（含消耗token数列）
    headers = RECORD_HEADERS
    
    # 覆盖写入整个文件
    # restval='0'：旧记录缺少「消耗token数」列时默认补 0（向后兼容）
    # extrasaction='ignore'：忽略记录里多余的字段，避免写入报错
    with open(RECORDS_FILE, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=headers, restval='0', extrasaction='ignore')
        writer.writeheader()  # 写入表头
        for record in records:
            writer.writerow(record)


# ==================== 辅助函数 ====================

def get_next_id(records):
    """
    获取下一条记录的ID
    
    功能：计算新记录应该使用的ID号（自增）
    
    参数：
        records: 当前所有记录的列表
    返回值：整数，下一个可用的ID
    """
    # 如果没有记录，从1开始
    if not records:
        return 1
    # 找到现有最大ID，加1
    return max(int(r['ID']) for r in records) + 1


# ==================== 查询函数 ====================

def get_daily_summary(date=None):
    """
    获取某天的汇总数据
    
    功能：查询指定日期的工作汇总信息
    
    参数：
        date: 日期字符串，格式为'YYYY-MM-DD'，默认为今天
    返回值：
        如果该日期有数据，返回字典包含汇总信息
        如果该日期无数据，返回None
    """
    # 默认查询今天
    if date is None:
        date = get_now().strftime('%Y-%m-%d')
    
    # 读取所有汇总数据
    summaries = read_summary()
    # 返回指定日期的数据，不存在则返回None
    return summaries.get(date, None)


def get_daily_records(date=None):
    """
    获取某天的所有记录
    
    功能：查询指定日期的所有截图分析记录
    
    参数：
        date: 日期字符串，格式为'YYYY-MM-DD'，默认为今天
    返回值：
        列表，包含该日期的所有记录字典
    """
    # 默认查询今天
    if date is None:
        date = get_now().strftime('%Y-%m-%d')
    
    # 读取所有记录
    records = read_records()
    # 筛选出指定日期的记录
    return [r for r in records if r['日期'] == date]


def print_today_summary():
    """
    打印今日工作汇总
    
    功能：在控制台打印今日的工作汇总和记录详情，便于快速查看
    
    参数：无
    返回值：无（直接打印到控制台）
    """
    # 打印汇总标题
    print("\n" + "="*50)
    print("今日工作汇总")
    print("="*50)
    
    # 获取今日汇总数据
    summary = get_daily_summary()
    if summary:
        # 打印各项汇总信息
        print(f"日期: {summary['日期']}")
        print(f"截图分析次数: {summary['记录条数']}")
        print(f"使用时长: {summary.get('使用时长(小时)', '0')} 小时")
        print(f"主要工作: {summary.get('主要工作', '暂无')}")
        print(f"最早使用: {summary.get('最早使用时间', '--:--')}")
        print(f"最晚使用: {summary.get('最晚使用时间', '--:--')}")
        
        # 打印各工作类型的时长（只显示有记录的类型）
        print("\n各类工作时长:")
        for t in WORK_TYPES:
            hours = float(summary.get(f'{t}时长(小时)', '0'))
            if hours > 0:  # 只显示时长大于0的类型
                print(f"  {t}: {hours:.2f} 小时")
        
        # 打印每小时记录数（只显示有记录的小时）
        print("\n每小时记录数:")
        for h in range(24):
            count = int(summary.get(f'{h:02d}:00记录数', '0'))
            if count > 0:  # 只显示记录数大于0的小时
                print(f"  {h:02d}:00 - {count} 条")
    else:
        print("今日暂无记录")
    
    # 打印记录详情标题
    print("\n" + "-"*50)
    print("今日记录详情")
    print("-"*50)
    
    # 获取今日所有记录
    records = get_daily_records()
    if records:
        # 打印每条记录的时间、类型和描述
        for r in records:
            print(f"[{r['时间']}] {r['工作类型']}: {r['工作描述']}")
    else:
        print("暂无记录")
    
    print("="*50)


def print_daily_summary(date=None):
    """
    打印指定日期的工作汇总
    
    功能：在控制台打印指定日期的工作汇总和记录详情，便于查看历史数据
    
    参数：
        date: 日期字符串，格式为'YYYY-MM-DD'，默认为今天
              例如：'2025-01-15' 表示2025年1月15日
    返回值：无（直接打印到控制台）
    
    使用示例：
        print_daily_summary()           # 打印今天的汇总
        print_daily_summary('2025-01-15')  # 打印2025年1月15日的汇总
    """
    # 默认查询今天
    if date is None:
        date = get_now().strftime('%Y-%m-%d')
    
    # 打印汇总标题
    print("\n" + "="*50)
    print(f"{date} 工作汇总")
    print("="*50)
    
    # 获取指定日期的汇总数据
    summary = get_daily_summary(date)
    if summary:
        # 打印各项汇总信息
        print(f"日期: {summary['日期']}")
        print(f"截图分析次数: {summary['记录条数']}")
        print(f"使用时长: {summary.get('使用时长(小时)', '0')} 小时")
        print(f"主要工作: {summary.get('主要工作', '暂无')}")
        print(f"最早使用: {summary.get('最早使用时间', '--:--')}")
        print(f"最晚使用: {summary.get('最晚使用时间', '--:--')}")
        
        # 打印各工作类型的时长（只显示有记录的类型）
        print("\n各类工作时长:")
        for t in WORK_TYPES:
            hours = float(summary.get(f'{t}时长(小时)', '0'))
            if hours > 0:  # 只显示时长大于0的类型
                print(f"  {t}: {hours:.2f} 小时")
        
        # 打印每小时记录数（只显示有记录的小时）
        print("\n每小时记录数:")
        for h in range(24):
            # 获取当前小时的记录数，如果为空则默认为'0'
            hour_count = summary.get(f'{h:02d}:00记录数', '0') or '0'
            count = int(hour_count)
            if count > 0:  # 只显示记录数大于0的小时
                print(f"  {h:02d}:00 - {count} 条")
    else:
        print(f"{date} 暂无记录")
    
    # 打印记录详情标题
    print("\n" + "-"*50)
    print(f"{date} 记录详情")
    print("-"*50)
    
    # 获取指定日期的所有记录
    records = get_daily_records(date)
    if records:
        # 打印每条记录的时间、类型和描述
        for r in records:
            print(f"[{r['时间']}] {r['工作类型']}: {r['工作描述']}")
    else:
        print("暂无记录")
    
    print("="*50)


# ==================== 主程序入口 ====================

if __name__ == '__main__':
    from screenshot import run_and_store
    
    time.sleep(1)
    # 初始化数据库
    init_db()
    
    # 测试：运行一次截图识别并存储
    print("运行截图识别...")
    result = run_and_store()
    print(f"识别结果: {result}")
    
    # 打印今日汇总
    print_today_summary()


# ==================== 报告模板存储 ====================

# 报告模板CSV文件路径
TEMPLATES_FILE = os.path.join(DB_DIR, 'report_templates.csv')


def get_default_templates():
    """
    获取默认报告模板
    
    返回值：模板列表
    """
    try:
        from template import get_default_templates as _get_templates
        return _get_templates()
    except ImportError:
        # 如果 template.py 不存在，返回最小模板
        return [
            {
                "name": "工作日报",
                "intro": "标准工作日报模板",
                "desc": "标准工作日报模板，包含今日完成、进展、问题、计划。",
                "is_cloud": True,
                "prompt": "## 工作日报\n\n### 今日完成\n{{完成事项}}\n\n### 当前进展\n{{进展}}\n\n### 遇到的问题\n{{问题}}\n\n### 明日计划\n{{计划}}"
            }
        ]


def init_templates_db():
    """
    初始化报告模板数据库
    
    功能：检查报告模板CSV文件是否存在，不存在则从默认模板创建
    """
    if not os.path.exists(TEMPLATES_FILE):
        # 从默认模板创建
        templates = get_default_templates()
        write_templates(templates)
        print(f"已从默认模板创建: {TEMPLATES_FILE}")


def read_templates():
    """
    读取报告模板数据
    
    返回值：列表，每个元素是一个模板字典
    """
    if not os.path.exists(TEMPLATES_FILE):
        init_templates_db()
    
    templates = []
    with open(TEMPLATES_FILE, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 转换 is_cloud 为布尔值
            row['is_cloud'] = row.get('is_cloud', '').lower() == 'true'
            # 确保 intro 字段存在
            if 'intro' not in row:
                row['intro'] = row.get('desc', '')[:20] + '...' if len(row.get('desc', '')) > 20 else row.get('desc', '')
            templates.append(row)
    
    return templates if templates else DEFAULT_TEMPLATES.copy()


def write_templates(templates):
    """
    写入报告模板数据
    
    参数：
        templates: 列表，每个元素是一个模板字典
    """
    # 确保目录存在
    os.makedirs(os.path.dirname(TEMPLATES_FILE), exist_ok=True)
    
    headers = ['name', 'intro', 'desc', 'is_cloud', 'prompt']
    
    with open(TEMPLATES_FILE, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for template in templates:
            # 确保 is_cloud 是字符串
            template['is_cloud'] = str(template.get('is_cloud', False))
            # 确保 intro 字段存在
            if 'intro' not in template:
                template['intro'] = template.get('desc', '')[:20] + '...' if len(template.get('desc', '')) > 20 else template.get('desc', '')
            writer.writerow(template)


def add_template(template):
    """
    添加一个新的报告模板
    
    参数：
        template: 字典，包含 name, desc, is_cloud, prompt
    """
    templates = read_templates()
    templates.append(template)
    write_templates(templates)


def delete_template(index):
    """
    删除指定索引的报告模板
    
    参数：
        index: 模板索引
    
    返回值：布尔值，表示是否删除成功
    """
    templates = read_templates()
    if 0 <= index < len(templates) and len(templates) > 1:
        templates.pop(index)
        write_templates(templates)
        return True
    return False


def update_template(index, template):
    """
    更新指定索引的报告模板
    
    参数：
        index: 模板索引
        template: 新的模板数据
    """
    templates = read_templates()
    if 0 <= index < len(templates):
        templates[index] = template
        write_templates(templates)


def export_templates(file_path):
    """
    导出报告模板到指定文件
    
    参数：
        file_path: 导出文件路径
    
    返回值：布尔值，表示是否导出成功
    """
    try:
        templates = read_templates()
        headers = ['name', 'intro', 'desc', 'is_cloud', 'prompt']
        
        with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
            writer.writeheader()
            for template in templates:
                template['is_cloud'] = str(template.get('is_cloud', False))
                # 确保 intro 字段存在
                if 'intro' not in template:
                    template['intro'] = template.get('desc', '')[:30] + '...' if len(template.get('desc', '')) > 30 else template.get('desc', '')
                writer.writerow(template)
        return True
    except Exception as e:
        print(f"导出模板失败: {e}")
        return False


def import_templates(file_path):
    """
    从指定文件导入报告模板
    
    参数：
        file_path: 导入文件路径
    
    返回值：导入的模板列表，失败返回空列表
    """
    try:
        templates = []
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                row['is_cloud'] = row.get('is_cloud', '').lower() == 'true'
                templates.append(row)
        return templates
    except Exception as e:
        print(f"导入模板失败: {e}")
        return []


# ==================== 报告文件管理 ====================

# 报告文件夹路径
REPORT_DIR = os.path.join(DB_DIR, 'report')


def init_report_dir():
    """初始化报告文件夹"""
    if not os.path.exists(REPORT_DIR):
        os.makedirs(REPORT_DIR)
        print(f"已创建报告文件夹: {REPORT_DIR}")


def save_report(title, content, report_type="日报", template_name="默认", tokens=None):
    """
    保存报告为 .md 文件
    
    参数：
        title: 报告标题
        content: 报告内容
        report_type: 报告类型（日报/周报/月报）
        template_name: 使用的模板名称
        tokens: 本次生成消耗的 token 字典 {'input':…, 'output':…, 'total':…}，可为 None
    
    返回值：
        保存的文件路径
    """
    init_report_dir()
    
    # 生成文件名：日期_时间_类型.md
    now = get_now()
    filename = f"{now.strftime('%Y%m%d_%H%M%S')}_{report_type}.md"
    filepath = os.path.join(REPORT_DIR, filename)
    
    # token 用量（缺省为 0）
    tokens = tokens or {}
    t_input = _to_int(tokens.get('input', 0))
    t_output = _to_int(tokens.get('output', 0))
    t_total = _to_int(tokens.get('total', 0), t_input + t_output)
    
    # 生成报告内容
    report_content = f"""# {title}

**生成时间：** {now.strftime('%Y-%m-%d %H:%M:%S')}
**报告类型：** {report_type}
**使用模板：** {template_name}
**消耗Token：** {t_total}
**输入Token：** {t_input}
**输出Token：** {t_output}

---

{content}
"""
    
    # 保存文件
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(report_content)
    
    print(f"报告已保存: {filepath}")
    return filepath


def get_report_list():
    """
    获取报告列表
    
    返回值：
        列表，每个元素是报告信息字典
    """
    init_report_dir()
    
    reports = []
    
    # 遍历报告文件夹
    for filename in os.listdir(REPORT_DIR):
        if filename.endswith('.md'):
            filepath = os.path.join(REPORT_DIR, filename)
            
            # 读取文件内容获取元信息
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 解析元信息
                title = filename.replace('.md', '')
                report_type = "日报"
                generate_time = ""
                template_name = "默认"
                word_count = len(content)
                token_total = 0
                token_input = 0
                token_output = 0
                
                # 简单解析标题
                lines = content.split('\n')
                for line in lines:
                    if line.startswith('# '):
                        title = line[2:].strip()
                    if '**报告类型：**' in line:
                        report_type = _parse_meta_value(line)
                    if '**生成时间：**' in line:
                        generate_time = _parse_meta_value(line)
                    if '**使用模板：**' in line:
                        template_name = _parse_meta_value(line)
                    if '**消耗Token：**' in line:
                        token_total = _to_int(_parse_meta_value(line))
                    if '**输入Token：**' in line:
                        token_input = _to_int(_parse_meta_value(line))
                    if '**输出Token：**' in line:
                        token_output = _to_int(_parse_meta_value(line))
                
                # 格式化时间
                if generate_time:
                    try:
                        dt = datetime.strptime(generate_time, '%Y-%m-%d %H:%M:%S')
                        today = get_now().strftime('%Y-%m-%d')
                        if dt.strftime('%Y-%m-%d') == today:
                            generate_time = f"今日 {dt.strftime('%H:%M')}"
                        else:
                            generate_time = dt.strftime('%m月%d日 %H:%M')
                    except:
                        pass
                
                reports.append({
                    'filename': filename,
                    'filepath': filepath,
                    'title': title,
                    'type': report_type,
                    'status': '已完成',
                    'time': generate_time,
                    'word_count': word_count,
                    'output_method': '直接输出',
                    'model': template_name,
                    'token_total': token_total,
                    'token_input': token_input,
                    'token_output': token_output
                })
            except Exception as e:
                print(f"读取报告失败 {filename}: {e}")
    
    # 按文件名排序（最新的在前面）
    reports.sort(key=lambda x: x['filename'], reverse=True)
    
    return reports


def get_token_stats():
    """
    统计 token 用量（活动分析 + 报告生成）

    返回值：字典
        today_analysis / all_analysis  —— 活动分析消耗 token（今日 / 全部）
        today_report / all_report      —— 报告生成消耗 token（今日 / 全部，取报告总 token）
        today_report_output / all_report_output —— 报告输出 token（今日 / 全部）
        today_total / all_total        —— 合计（今日 / 全部）
    """
    today = get_today()

    # ---- 活动分析 token（来自 records.csv 的「消耗token数」列）----
    today_analysis = 0
    all_analysis = 0
    for r in read_records():
        tokens = _to_int(r.get('消耗token数', 0))
        all_analysis += tokens
        if r.get('日期', '') == today:
            today_analysis += tokens

    # ---- 报告生成 token（来自各报告 .md 文件头）----
    today_report = 0
    all_report = 0
    today_report_output = 0
    all_report_output = 0
    for rep in get_report_list():
        total = _to_int(rep.get('token_total', 0))
        output = _to_int(rep.get('token_output', 0))
        all_report += total
        all_report_output += output
        # get_report_list 将当天报告的时间格式化为「今日 HH:MM」
        if str(rep.get('time', '')).startswith('今日'):
            today_report += total
            today_report_output += output

    return {
        'today_analysis': today_analysis,
        'today_report': today_report,
        'today_report_output': today_report_output,
        'today_total': today_analysis + today_report,
        'all_analysis': all_analysis,
        'all_report': all_report,
        'all_report_output': all_report_output,
        'all_total': all_analysis + all_report,
    }


def read_report(filepath):
    """
    读取报告内容
    
    参数：
        filepath: 报告文件路径
    
    返回值：
        报告内容字符串
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"读取报告失败: {e}")
        return ""


def delete_report(filepath):
    """
    删除报告文件
    
    参数：
        filepath: 报告文件路径
    
    返回值：
        布尔值，表示是否删除成功
    """
    try:
        if os.path.exists(filepath):
            os.remove(filepath)
            print(f"报告已删除: {filepath}")
            return True
        return False
    except Exception as e:
        print(f"删除报告失败: {e}")
        return False


# ==================== 设置管理 ====================

# 设置文件路径
SETTINGS_FILE = os.path.join(DB_DIR, 'settings.csv')

# 默认设置
DEFAULT_SETTINGS = {
    'auto_check_update': 'true'
}


def init_settings():
    """初始化设置文件"""
    if not os.path.exists(SETTINGS_FILE):
        os.makedirs(DB_DIR, exist_ok=True)
        with open(SETTINGS_FILE, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['key', 'value'])
            for key, value in DEFAULT_SETTINGS.items():
                writer.writerow([key, value])
        print(f"已创建设置文件: {SETTINGS_FILE}")


def read_settings():
    """
    读取设置
    
    返回值：
        字典，包含所有设置项
    """
    init_settings()
    settings = DEFAULT_SETTINGS.copy()
    
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                settings[row['key']] = row['value']
    except Exception as e:
        print(f"读取设置失败: {e}")
    
    return settings


def write_setting(key, value):
    """
    写入单个设置
    
    参数：
        key: 设置键名
        value: 设置值
    """
    settings = read_settings()
    settings[key] = str(value).lower() if isinstance(value, bool) else str(value)
    
    with open(SETTINGS_FILE, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['key', 'value'])
        for k, v in settings.items():
            writer.writerow([k, v])


def get_setting(key, default=None):
    """
    获取单个设置值
    
    参数：
        key: 设置键名
        default: 默认值
    
    返回值：
        设置值
    """
    settings = read_settings()
    return settings.get(key, default)
