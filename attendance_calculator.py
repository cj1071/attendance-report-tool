#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤统计计算器 - 工时计算和夜班补贴模块
"""

from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional
import pandas as pd

class AttendanceCalculator:
    """考勤统计计算器"""
    
    def __init__(self):
        # 夜班补贴标准（元/人/日）
        self.night_allowance_rate = 10.0
        # 夜班补贴最低工时要求（小时）
        self.night_allowance_min_hours = 11.5
        
    def parse_time_string(self, time_str: str) -> Optional[float]:
        """
        解析时间字符串为小时数（24小时制）
        
        Args:
            time_str: 时间字符串，如 "08:30", "20:00"
            
        Returns:
            float: 小时数，如 8.5, 20.0；解析失败返回None
        """
        if not time_str or pd.isna(time_str):
            return None
            
        try:
            # 处理字符串格式
            time_str = str(time_str).strip()
            
            # 如果是Excel时间格式（浮点数）
            if time_str.replace('.', '').isdigit():
                hours = float(time_str) * 24
                return hours % 24
            
            # 处理 HH:MM 格式
            if ':' in time_str:
                parts = time_str.split(':')
                if len(parts) == 2:
                    hours = int(parts[0])
                    minutes = int(parts[1])
                    return hours + minutes / 60.0
            
            # 处理纯数字格式（假设为小时）
            if time_str.isdigit():
                return float(time_str)
                
        except (ValueError, TypeError):
            pass
            
        return None
    
    def is_night_shift(self, start_time: float, end_time: float) -> bool:
        """
        判断是否为夜班
        
        夜班判定标准：
        - 上工时间在 20:00 或之后 (start_time >= 20.0)
        - 或者跨天且上工时间在凌晨 (end_time < start_time and start_time < 20.0)
        
        简化规则：
        - 20:00 及以后上工 → 夜班
        - 其他 → 白班
        
        Args:
            start_time: 上工时间（小时）
            end_time: 下工时间（小时）
            
        Returns:
            bool: True表示夜班，False表示白班
        """
        if start_time is None or end_time is None:
            return False
        
        # 上工时间 >= 20:00 即为夜班
        if start_time >= 20.0:
            return True
        
        # 跨天且上工在夜间时段（如凌晨上工）
        if end_time < start_time and start_time < 8.0:
            return True
        
        return False
    
    def calculate_total_hours(self, start_time: float, end_time: float) -> float:
        """
        计算总工作时长（支持跨天）
        
        Args:
            start_time: 上工时间（小时）
            end_time: 下工时间（小时）
            
        Returns:
            float: 总工作时长（小时）
        """
        if start_time is None or end_time is None:
            return 0.0
        
        if end_time < start_time:
            # 跨天计算：(24 - 上工时间) + 下工时间
            return (24.0 - start_time) + end_time
        else:
            # 同一天：下工时间 - 上工时间
            return end_time - start_time
    
    def calculate_day_shift_hours(self, start_time: float, end_time: float) -> float:
        """
        计算白班有效工时
        
        扣减逻辑（按优先级顺序判断）：
        1. 上工 > 17:00 → 不扣
        2. 下工 ≤ 11:00 → 不扣
        3. 上工 > 11:00 且 ≤ 17:00 → 扣0.5h
        4. 下工 ≤ 17:00 且 上工 ≤ 11:00 → 扣0.5h
        5. 上工 ≤ 11:00 且 下工 ≥ 17:00 → 扣1h（默认情况）
        6. 其他情况按 (下工 - 上工) 计算（无扣减）
        
        Args:
            start_time: 上工时间（小时）
            end_time: 下工时间（小时）
            
        Returns:
            float: 有效工时（小时）
        """
        if start_time is None or end_time is None:
            return 0.0
        
        total_hours = self.calculate_total_hours(start_time, end_time)
        
        # 按优先级顺序判断扣减规则
        if start_time > 17.0:
            # 上工 > 17:00 → 不扣
            return total_hours
        elif end_time <= 11.0:
            # 下工 ≤ 11:00 → 不扣
            return total_hours
        elif 11.0 < start_time <= 17.0:
            # 上工 > 11:00 且 ≤ 17:00 → 扣0.5h
            return max(0.0, total_hours - 0.5)
        elif end_time <= 17.0 and start_time <= 11.0:
            # 下工 ≤ 17:00 且 上工 ≤ 11:00 → 扣0.5h
            return max(0.0, total_hours - 0.5)
        elif start_time <= 11.0 and end_time >= 17.0:
            # 上工 ≤ 11:00 且 下工 ≥ 17:00 → 扣1h（默认情况）
            return max(0.0, total_hours - 1.0)
        else:
            # 其他情况按总时长计算（无扣减）
            return total_hours
    
    def calculate_night_shift_hours(self, start_time: float, end_time: float) -> float:
        """
        计算夜班有效工时
        
        所有夜班统一扣除 0.5 小时休息时间
        
        Args:
            start_time: 上工时间（小时）
            end_time: 下工时间（小时）
            
        Returns:
            float: 有效工时（小时）
        """
        if start_time is None or end_time is None:
            return 0.0
        
        total_hours = self.calculate_total_hours(start_time, end_time)
        return max(0.0, total_hours - 0.5)
    
    def calculate_working_hours(self, start_time_str: str, end_time_str: str) -> Dict:
        """
        计算工作时长和班次信息
        
        Args:
            start_time_str: 上工时间字符串
            end_time_str: 下工时间字符串
            
        Returns:
            dict: 包含班次类型、总时长、有效工时等信息
        """
        start_time = self.parse_time_string(start_time_str)
        end_time = self.parse_time_string(end_time_str)
        
        if start_time is None or end_time is None:
            return {
                'shift_type': '无效',
                'total_hours': 0.0,
                'effective_hours': 0.0,
                'is_night_shift': False,
                'start_time': None,
                'end_time': None
            }
        
        is_night = self.is_night_shift(start_time, end_time)
        total_hours = self.calculate_total_hours(start_time, end_time)
        
        if is_night:
            effective_hours = self.calculate_night_shift_hours(start_time, end_time)
            shift_type = '夜班'
        else:
            effective_hours = self.calculate_day_shift_hours(start_time, end_time)
            shift_type = '白班'
        
        return {
            'shift_type': shift_type,
            'total_hours': round(total_hours, 2),
            'effective_hours': round(effective_hours, 2),
            'is_night_shift': is_night,
            'start_time': start_time,
            'end_time': end_time
        }
    
    def calculate_night_allowance(self, effective_hours: float, is_night_shift: bool) -> float:
        """
        计算夜班补贴
        
        发放条件（同时满足）：
        1. 当日为夜班
        2. 扣除0.5小时休息后的有效夜班工时 ≥ 11.5 小时
        
        Args:
            effective_hours: 有效工时
            is_night_shift: 是否为夜班
            
        Returns:
            float: 夜班补贴金额（元）
        """
        if is_night_shift and effective_hours >= self.night_allowance_min_hours:
            return self.night_allowance_rate
        return 0.0
    
    def format_time(self, hours: float) -> str:
        """
        将小时数格式化为时间字符串
        
        Args:
            hours: 小时数
            
        Returns:
            str: 格式化的时间字符串，如 "08:30"
        """
        if hours is None:
            return ""
        
        hours = hours % 24  # 确保在24小时范围内
        hour_part = int(hours)
        minute_part = int((hours - hour_part) * 60)
        return f"{hour_part:02d}:{minute_part:02d}"
    
    def process_attendance_record(self, record: Dict) -> Dict:
        """
        处理单条考勤记录
        
        Args:
            record: 包含姓名、上工时间、下工时间等信息的字典
            
        Returns:
            dict: 处理后的考勤统计信息
        """
        start_time_str = record.get('上工时间', '')
        end_time_str = record.get('下工时间', '')
        
        # 计算工时信息
        work_info = self.calculate_working_hours(start_time_str, end_time_str)
        
        # 计算夜班补贴
        night_allowance = self.calculate_night_allowance(
            work_info['effective_hours'], 
            work_info['is_night_shift']
        )
        
        return {
            'name': record.get('姓名', ''),
            'company': record.get('劳务公司', ''),
            'date': record.get('日期', ''),
            'start_time': start_time_str,
            'end_time': end_time_str,
            'start_time_formatted': self.format_time(work_info['start_time']),
            'end_time_formatted': self.format_time(work_info['end_time']),
            'shift_type': work_info['shift_type'],
            'total_hours': work_info['total_hours'],
            'effective_hours': work_info['effective_hours'],
            'night_allowance': night_allowance,
            'is_night_shift': work_info['is_night_shift']
        }

# 测试代码
if __name__ == "__main__":
    calculator = AttendanceCalculator()
    
    # 测试用例
    test_cases = [
        {"name": "张三", "上工时间": "08:00", "下工时间": "17:00"},  # 白班正常
        {"name": "李四", "上工时间": "20:00", "下工时间": "08:00"},  # 夜班跨天
        {"name": "王五", "上工时间": "22:00", "下工时间": "06:00"},  # 夜班跨天
        {"name": "赵六", "上工时间": "12:00", "下工时间": "18:00"},  # 白班晚到
        {"name": "钱七", "上工时间": "20:00", "下工时间": "07:30"},  # 夜班不满11.5h
    ]
    
    print("🧪 考勤计算器测试")
    print("=" * 80)
    
    for i, case in enumerate(test_cases, 1):
        result = calculator.process_attendance_record(case)
        print(f"\n测试用例 {i}: {result['name']}")
        print(f"  上工时间: {result['start_time']} → 下工时间: {result['end_time']}")
        print(f"  班次类型: {result['shift_type']}")
        print(f"  总工时: {result['total_hours']}h")
        print(f"  有效工时: {result['effective_hours']}h")
        print(f"  夜班补贴: {result['night_allowance']}元")
