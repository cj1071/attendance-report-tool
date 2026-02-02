#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化的报表生成脚本 - 修复版本
"""

import os
import sys
from excel_report_generator_fixed import ExcelReportGenerator

def find_input_file():
    """在当前目录查找输入文件"""
    current_dir = os.getcwd()
    
    # 查找可能的输入文件
    for filename in os.listdir(current_dir):
        if filename.endswith(('.xls', '.xlsx')) and '劳务签到表' in filename:
            return os.path.join(current_dir, filename)
    
    # 如果没找到，列出所有Excel文件
    excel_files = [f for f in os.listdir(current_dir) if f.endswith(('.xls', '.xlsx'))]
    
    if not excel_files:
        print("❌ 当前目录没有找到Excel文件")
        return None
    
    print("📁 当前目录的Excel文件:")
    for i, filename in enumerate(excel_files, 1):
        print(f"  {i}. {filename}")
    
    try:
        choice = input("\n请选择要处理的文件编号 (直接回车选择第1个): ").strip()
        if not choice:
            choice = "1"
        
        index = int(choice) - 1
        if 0 <= index < len(excel_files):
            return os.path.join(current_dir, excel_files[index])
        else:
            print("❌ 无效的选择")
            return None
    except ValueError:
        print("❌ 请输入有效的数字")
        return None

def main():
    print("🚀 员工工时报表生成工具 (修复版)")
    print("=" * 50)
    
    # 查找输入文件
    input_file = find_input_file()
    if not input_file:
        sys.exit(1)
    
    print(f"📄 处理文件: {os.path.basename(input_file)}")
    
    try:
        # 创建报表生成器
        generator = ExcelReportGenerator()
        
        # 读取输入文件
        print("\n📖 正在读取Excel文件...")
        generator.read_input_excel(input_file)
        
        if not generator.raw_data:
            print("❌ 没有读取到有效数据")
            sys.exit(1)
        
        # 为每个公司生成报表
        print(f"\n📊 开始生成报表...")
        generated_files = []
        
        for company in sorted(generator.companies):
            print(f"  正在生成 {company} 的报表...")
            report_info = generator.generate_company_report(company)
            if report_info:
                filepath = generator.save_company_report(report_info, os.getcwd())
                generated_files.append(filepath)
        
        print(f"\n✅ 报表生成完成!")
        print(f"📁 共生成 {len(generated_files)} 个文件:")
        for filepath in generated_files:
            print(f"  ✓ {os.path.basename(filepath)}")
        
        print(f"\n💡 文件保存在: {os.getcwd()}")
        print("\n🎯 新功能:")
        print("  ✓ 按原始数据顺序排列员工")
        print("  ✓ 多次签到动态扩展列")
        print("  ✓ 斑马纹区分不同员工")
        print("  ✓ 表头12号字体，数据10号字体")
        print("  ✓ 除A1外所有表头居中对齐")
        
    except Exception as e:
        print(f"❌ 处理失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
