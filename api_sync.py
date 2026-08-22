"""
数据同步模块 - 与服务器同步数据
"""

import requests
import json
import os
import sys
from datetime import datetime

# API 配置
API_BASE_URL = "http://129.204.12.226"

# 登录状态文件
def get_login_email():
    """获取当前登录的邮箱"""
    # 打包后在exe同级目录，开发时在脚本所在目录
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    config_file = os.path.join(base_dir, 'data', 'login_state.json')
    
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                state = json.load(f)
                email = state.get('email')
                print(f"[同步] 获取到登录邮箱: {email}")
                return email
        except Exception as e:
            print(f"[同步] 读取登录状态失败: {e}")
    else:
        print(f"[同步] 登录状态文件不存在: {config_file}")
    return None


def sync_work_record(date, time, work_type, description, duration, tokens=None):
    """
    同步工作记录到服务器
    
    参数:
        date: 日期 (YYYY-MM-DD)
        time: 时间 (HH:MM:SS)
        work_type: 工作类型
        description: 工作描述
        duration: 持续时长(分钟)
        tokens: 本次识别消耗的 token 用量字典 {'input':…, 'output':…, 'total':…}，可为 None
    
    返回值:
        (success: bool, message: str)
    """
    email = get_login_email()
    if not email:
        print("[同步] 未登录，跳过同步工作记录")
        return False, "未登录"
    
    try:
        data = {
            "email": email,
            "date": date,
            "time": time,
            "work_type": work_type,
            "description": description,
            "duration": duration,
            "tokens": tokens or {}
        }
        print(f"[同步] 同步工作记录: {data}")
        response = requests.post(
            f"{API_BASE_URL}/api/user/record",
            json=data,
            timeout=5
        )
        result = response.json()
        print(f"[同步] 工作记录响应: {result}")
        return result.get('success', False), result.get('message', '')
    except Exception as e:
        print(f"[同步] 同步工作记录失败: {e}")
        return False, str(e)


def sync_daily_summary(summary_data):
    """
    同步每日汇总到服务器
    
    参数:
        summary_data: 汇总数据字典
    
    返回值:
        (success: bool, message: str)
    """
    email = get_login_email()
    if not email:
        print("[同步] 未登录，跳过同步每日汇总")
        return False, "未登录"
    
    try:
        # 构建请求数据
        data = {"email": email}
        
        # 映射字段
        field_map = {
            '日期': 'date',
            '记录条数': 'record_count',
            '使用时长(小时)': 'usage_hours',
            '主要工作': 'main_work',
            '最早使用时间': 'earliest_time',
            '最晚使用时间': 'latest_time'
        }
        
        for cn_key, en_key in field_map.items():
            if cn_key in summary_data:
                data[en_key] = summary_data[cn_key]
        
        # 工作类型时长
        work_types = ["开发", "沟通", "生活", "学习", "设计", "管理", "文档", 
                      "娱乐", "产品", "会议", "运维", "测试", "数据分析", "其他"]
        for wt in work_types:
            key = f'{wt}时长(小时)'
            if key in summary_data:
                data[f'{wt}_hours'] = summary_data[key]
        
        # 每小时记录数
        for h in range(24):
            key = f'{h:02d}:00记录数'
            if key in summary_data:
                data[f'hour_{h:02d}'] = summary_data[key]
        
        print(f"[同步] 同步每日汇总: {data}")
        response = requests.post(
            f"{API_BASE_URL}/api/user/daily-summary",
            json=data,
            timeout=5
        )
        result = response.json()
        print(f"[同步] 每日汇总响应: {result}")
        return result.get('success', False), result.get('message', '')
    except Exception as e:
        print(f"[同步] 同步每日汇总失败: {e}")
        return False, str(e)


def sync_login(email, password):
    """
    同步登录到服务器
    
    参数:
        email: 邮箱
        password: 密码
    
    返回值:
        (success: bool, message: str)
    """
    try:
        response = requests.post(
            f"{API_BASE_URL}/api/login",
            json={"email": email, "password": password},
            timeout=10
        )
        result = response.json()
        return result.get('success', False), result.get('message', '')
    except Exception as e:
        return False, str(e)


def get_user_stats():
    """
    获取用户统计信息
    
    返回值:
        (success: bool, stats: dict or message: str)
    """
    email = get_login_email()
    if not email:
        return False, "未登录"
    
    try:
        response = requests.get(
            f"{API_BASE_URL}/api/user/stats/{email}",
            timeout=5
        )
        result = response.json()
        if result.get('success'):
            return True, result.get('stats', {})
        return False, result.get('message', '获取失败')
    except Exception as e:
        return False, str(e)


def sync_report_generated():
    """
    同步报告生成事件到服务器
    
    返回值:
        (success: bool, message: str)
    """
    email = get_login_email()
    if not email:
        print("[同步] 未登录，跳过同步报告生成事件")
        return False, "未登录"
    
    try:
        print(f"[同步] 同步报告生成事件: {email}")
        response = requests.post(
            f"{API_BASE_URL}/api/user/report-generated",
            json={"email": email},
            timeout=5
        )
        result = response.json()
        print(f"[同步] 报告生成事件响应: {result}")
        return result.get('success', False), result.get('message', '')
    except Exception as e:
        print(f"[同步] 同步报告生成事件失败: {e}")
        return False, str(e)


def upload_report(title, content, report_type="日报", tokens=None):
    """
    上传报告内容到服务器
    
    参数:
        title: 报告标题
        content: 报告内容（Markdown格式）
        report_type: 报告类型（日报/周报/月报）
        tokens: 本次生成消耗的 token 用量字典 {'input':…, 'output':…, 'total':…}，可为 None
    
    返回值:
        (success: bool, message: str)
    """
    email = get_login_email()
    if not email:
        print("[同步] 未登录，跳过上传报告")
        return False, "未登录"
    
    try:
        print(f"[同步] 上传报告: {title}")
        response = requests.post(
            f"{API_BASE_URL}/api/user/upload-report",
            json={
                "email": email,
                "title": title,
                "content": content,
                "type": report_type,
                "tokens": tokens or {}
            },
            timeout=30
        )
        result = response.json()
        print(f"[同步] 上传报告响应: {result}")
        return result.get('success', False), result.get('message', '')
    except Exception as e:
        print(f"[同步] 上传报告失败: {e}")
        return False, str(e)
