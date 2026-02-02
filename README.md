# 📊 员工工时报表生成工具

一个纯 Python 工具，用于将劳务签到表转换为按公司分组的月度考勤报表。

## 🎯 功能特点

- ✅ **无数据库依赖** - 纯文件处理
- ✅ **跨年处理** - 智能处理跨年份的工时数据
- ✅ **时间解析** - 正确解析 Excel 时间格式
- ✅ **按公司分组** - 自动按劳务公司生成独立报表
- ✅ **双报表输出** - 员工工时表 + 考勤统计表

## 📁 输入输出格式

### 输入文件
```
XX月劳务签到表.xls
├── 6.2 (工作表)    # 6月2日数据
├── 6.3 (工作表)    # 6月3日数据
└── ...
```

每个工作表包含：序号、姓名、劳务公司、上工时间、下工时间、备注

### 输出文件
```
employee_hours-公司A.xlsx      # 公司A的月度工时报表
employee_hours-公司B.xlsx      # 公司B的月度工时报表
attendance_stats-公司A.xlsx    # 公司A的考勤统计报表
attendance_stats-公司B.xlsx    # 公司B的考勤统计报表
```

## 🚀 快速开始

### GUI 界面（推荐）
```bash
python gui_app.py
```

### 命令行
```bash
python run_report_fixed.py
```

## 📦 安装依赖

```bash
pip install -r requirements.txt
```

依赖包：pandas, openpyxl, xlrd, numpy

## 🔧 核心特性

### 智能年份推断
- 2025年1月导入12月数据 → 自动识别为2024年12月
- 2025年2月导入1月数据 → 自动识别为2025年1月

### 时间格式处理
- 支持 Excel 时间序列号
- 支持字符串时间格式
- 自动处理 1899 年基准日期问题

## 🏗️ CI/CD

项目使用 GitHub Actions 进行跨平台构建：
- Windows (.exe)
- macOS (.app)
- Linux

通过 tag 推送触发构建：
```bash
git tag v1.0.0
git push origin v1.0.0
```

## � 注意事项

1. **文件格式**：支持 .xls 和 .xlsx 格式
2. **工作表命名**：必须使用 "月.日" 格式（如：6.2）
3. **数据列**：姓名、劳务公司、上工、下工列必须存在

## 📄 License

MIT
