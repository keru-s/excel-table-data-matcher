#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据匹配工具启动脚本

检查依赖并启动工具
"""

import sys
import os

def check_dependencies():
    """检查必要的依赖"""
    required_packages = [
        'PyQt6',
        'pandas',
        'numpy',
        'xlrd',
        'openpyxl',
        'psutil'
    ]

    missing_packages = []

    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    return missing_packages

def main():
    """主函数"""
    print("Excel数据匹配工具 v2.0 (性能优化版)")
    print("=" * 50)

    # 检查Python版本
    if sys.version_info < (3, 8):
        print("错误: 需要Python 3.8或更高版本")
        print(f"当前版本: {sys.version}")
        input("按回车键退出...")
        return

    # 检查依赖
    print("检查依赖包...")
    missing = check_dependencies()

    if missing:
        print(f"缺少以下依赖包: {', '.join(missing)}")
        print("\n请运行以下命令安装依赖:")
        print("python install_dependencies.py")
        input("\n按回车键退出...")
        return

    print("✓ 所有依赖包检查通过")
    print("\n启动Excel数据匹配工具...")

    try:
        # 导入并启动工具
        from excel_tool import main as excel_main
        excel_main()
    except Exception as e:
        print(f"\n启动失败: {e}")
        input("按回车键退出...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n程序被用户中断")
    except Exception as e:
        print(f"\n程序出错: {e}")
        input("按回车键退出...")
