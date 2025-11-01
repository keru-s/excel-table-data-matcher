#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具依赖安装脚本

自动安装所需的Python包
"""

import subprocess
import sys
import os

def install_package(package):
    """安装单个包"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError:
        return False

def check_package(package_name):
    """检查包是否已安装"""
    try:
        __import__(package_name)
        return True
    except ImportError:
        return False

def main():
    """主安装函数"""
    print("Excel数据匹配工具 - 依赖安装程序")
    print("=" * 50)

    # 检查Python版本
    if sys.version_info < (3, 8):
        print("错误: 需要Python 3.8或更高版本")
        print(f"当前版本: {sys.version}")
        return False

    print(f"Python版本: {sys.version}")
    print()

    # 需要安装的包列表
    packages = [
        ("pandas", "pandas>=2.0.0"),
        ("numpy", "numpy>=1.24.0"),
        ("PyQt6", "PyQt6>=6.5.0"),
        ("xlrd", "xlrd>=2.0.0"),
        ("openpyxl", "openpyxl>=3.1.0"),
        ("psutil", "psutil>=5.9.0"),
        ("chardet", "chardet>=5.0.0")
    ]

    print("检查已安装的包...")

    to_install = []
    for package_name, package_spec in packages:
        if check_package(package_name):
            print(f"✓ {package_name} 已安装")
        else:
            print(f"✗ {package_name} 未安装")
            to_install.append(package_spec)

    if not to_install:
        print("\n所有依赖包都已安装！")
        return True

    print(f"\n需要安装 {len(to_install)} 个包:")
    for package in to_install:
        print(f"  - {package}")

    # 询问用户是否继续
    response = input("\n是否继续安装? (Y/n): ").strip().lower()
    if response in ['n', 'no']:
        print("安装已取消")
        return False

    print("\n开始安装依赖包...")

    # 升级pip
    print("升级pip...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        print("✓ pip升级成功")
    except subprocess.CalledProcessError:
        print("⚠ pip升级失败，继续安装其他包...")

    # 安装包
    success_count = 0
    for package in to_install:
        print(f"\n安装 {package}...")
        if install_package(package):
            print(f"✓ {package} 安装成功")
            success_count += 1
        else:
            print(f"✗ {package} 安装失败")

    print(f"\n安装完成: {success_count}/{len(to_install)} 个包安装成功")

    if success_count == len(to_install):
        print("\n🎉 所有依赖包安装成功！")
        print("\n现在可以运行Excel工具了:")
        print("python excel_tool.py")
        return True
    else:
        print("\n⚠ 部分包安装失败，请手动安装失败的包")
        return False

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n安装被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n安装过程中出错: {e}")
        sys.exit(1)
