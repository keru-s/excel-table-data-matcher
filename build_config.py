#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 打包配置脚本
使用 PyInstaller 打包 Excel 工具为 exe 文件

使用方法:
1. 在 Windows 系统中运行此脚本
2. 或者使用 auto-py-to-exe 工具时参考此配置
"""

import os
import sys

# PyInstaller 配置参数
PYINSTALLER_CONFIG = {
    # 主入口文件
    'entry_point': 'start_excel_tool.py',

    # 输出配置
    'name': 'Excel数据匹配工具',
    'onefile': False,  # 桌面程序推荐 onedir，启动更快、问题更少
    'windowed': True,  # 不显示控制台窗口

    # 图标文件（如果有的话）
    'icon': 'assets/app_icon.ico',

    # 需要包含的数据文件和目录
    'add_data': [
        # ('source_path', 'destination_path_in_exe')
    ],

    # 需要包含的二进制文件
    'add_binary': [
        # ('source_path', 'destination_path_in_exe')
    ],

    # 隐藏导入（解决某些模块无法自动检测的问题）
    'hidden_imports': [
        'PyQt6.QtCore',
        'PyQt6.QtWidgets',
        'PyQt6.QtGui',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.properties',
        'numpy.core._methods',
        'numpy.lib.format',
        'xlrd.xldate',
        'openpyxl.cell',
        'openpyxl.workbook',
        'psutil._psutil_windows',
        'chardet',
    ],

    # 排除的模块（减小文件大小）
    'excludes': [
        'matplotlib',
        'tkinter',
        'test',
        'unittest',
        'pydoc',
        'doctest',
    ],

    # 其他选项
    'clean': True,  # 清理临时文件
    'noconfirm': True,  # 不询问覆盖
}

def generate_pyinstaller_command():
    """生成 PyInstaller 命令行"""
    cmd_parts = ['pyinstaller']

    # 基本参数
    if PYINSTALLER_CONFIG['onefile']:
        cmd_parts.append('--onefile')

    if PYINSTALLER_CONFIG['windowed']:
        cmd_parts.append('--windowed')

    if PYINSTALLER_CONFIG['clean']:
        cmd_parts.append('--clean')

    if PYINSTALLER_CONFIG['noconfirm']:
        cmd_parts.append('--noconfirm')

    # 名称
    cmd_parts.extend(['--name', f'"{PYINSTALLER_CONFIG["name"]}"'])

    # 图标
    if PYINSTALLER_CONFIG['icon']:
        cmd_parts.extend(['--icon', PYINSTALLER_CONFIG['icon']])

    # 隐藏导入
    for module in PYINSTALLER_CONFIG['hidden_imports']:
        cmd_parts.extend(['--hidden-import', module])

    # 排除模块
    for module in PYINSTALLER_CONFIG['excludes']:
        cmd_parts.extend(['--exclude-module', module])

    # 添加数据文件
    for src, dst in PYINSTALLER_CONFIG['add_data']:
        cmd_parts.extend(['--add-data', f'{src};{dst}'])

    # 入口文件
    cmd_parts.append(PYINSTALLER_CONFIG['entry_point'])

    return ' '.join(cmd_parts)

def generate_spec_file():
    """生成 .spec 文件内容"""
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['{PYINSTALLER_CONFIG["entry_point"]}'],
    pathex=[],
    binaries={PYINSTALLER_CONFIG["add_binary"]},
    datas={PYINSTALLER_CONFIG["add_data"]},
    hiddenimports={PYINSTALLER_CONFIG["hidden_imports"]},
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes={PYINSTALLER_CONFIG["excludes"]},
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{PYINSTALLER_CONFIG["name"]}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console={'not ' + str(PYINSTALLER_CONFIG["windowed"])},
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon={repr(PYINSTALLER_CONFIG["icon"]) if PYINSTALLER_CONFIG["icon"] else None},
)
'''
    return spec_content

def main():
    """主函数"""
    print("Excel工具打包配置生成器")
    print("=" * 50)

    # 生成 PyInstaller 命令
    cmd = generate_pyinstaller_command()
    print("PyInstaller 命令:")
    print(cmd)
    print()

    # 生成 spec 文件
    spec_content = generate_spec_file()
    spec_filename = f'{PYINSTALLER_CONFIG["name"]}.spec'

    try:
        with open(spec_filename, 'w', encoding='utf-8') as f:
            f.write(spec_content)
        print(f"✓ 已生成 spec 文件: {spec_filename}")
    except Exception as e:
        print(f"✗ 生成 spec 文件失败: {e}")

    print()
    print("使用说明:")
    print("1. 方法一：直接运行上面的 PyInstaller 命令")
    print("2. 方法二：使用生成的 spec 文件")
    print(f"   pyinstaller {spec_filename}")
    print("3. 方法三：使用 auto-py-to-exe 图形界面工具")

if __name__ == "__main__":
    main()
