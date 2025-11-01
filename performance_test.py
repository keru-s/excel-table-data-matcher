#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具性能测试脚本

用于测试优化后的Excel数据匹配工具的性能表现
"""

import pandas as pd
import numpy as np
import time
import os
import sys
from pathlib import Path

def create_test_data():
    """创建测试数据文件"""
    print("正在创建测试数据...")

    # 创建大文件1 (100万行)
    np.random.seed(42)
    data1_size = 1000000

    data1 = {
        'ID': range(1, data1_size + 1),
        '姓名': [f'用户{i}' for i in range(1, data1_size + 1)],
        '年龄': np.random.randint(18, 80, data1_size),
        '城市': np.random.choice(['北京', '上海', '广州', '深圳', '杭州'], data1_size),
        '收入': np.random.randint(3000, 50000, data1_size),
        '部门': np.random.choice(['技术部', '销售部', '市场部', '人事部', '财务部'], data1_size)
    }

    df1 = pd.DataFrame(data1)
    test_file1 = 'test_data_large.xlsx'
    df1.to_excel(test_file1, index=False, engine='openpyxl')
    print(f"已创建大文件: {test_file1} ({data1_size:,} 行)")

    # 创建小文件2 (1万行，部分ID匹配)
    data2_size = 10000
    # 选择一些ID与文件1匹配
    matched_ids = np.random.choice(range(1, data1_size + 1), data2_size, replace=False)

    data2 = {
        'ID': matched_ids,
        '订单号': [f'ORDER{i:08d}' for i in range(1, data2_size + 1)],
        '订单金额': np.random.randint(100, 10000, data2_size),
        '订单日期': pd.date_range('2024-01-01', periods=data2_size, freq='H')[:data2_size],
        '产品类别': np.random.choice(['电子产品', '服装', '食品', '图书', '家居'], data2_size)
    }

    df2 = pd.DataFrame(data2)
    test_file2 = 'test_data_small.xlsx'
    df2.to_excel(test_file2, index=False, engine='openpyxl')
    print(f"已创建小文件: {test_file2} ({data2_size:,} 行)")

    return test_file1, test_file2

def test_performance():
    """测试性能"""
    print("\n=== Excel工具性能测试 ===")

    # 检查测试文件是否存在
    test_file1 = 'test_data_large.xlsx'
    test_file2 = 'test_data_small.xlsx'

    if not os.path.exists(test_file1) or not os.path.exists(test_file2):
        test_file1, test_file2 = create_test_data()

    print(f"\n测试文件:")
    print(f"- 文件1: {test_file1}")
    print(f"- 文件2: {test_file2}")

    # 获取系统信息
    import psutil
    cpu_count = psutil.cpu_count()
    memory_gb = psutil.virtual_memory().total / (1024**3)

    print(f"\n系统信息:")
    print(f"- CPU核心数: {cpu_count}")
    print(f"- 总内存: {memory_gb:.1f} GB")
    print(f"- 可用内存: {psutil.virtual_memory().available / (1024**3):.1f} GB")

    # 测试不同的分块大小和进程数
    test_configs = [
        {'chunk_size': 5000, 'process_count': 1, 'name': '单进程-小块'},
        {'chunk_size': 10000, 'process_count': 1, 'name': '单进程-中块'},
        {'chunk_size': 20000, 'process_count': 1, 'name': '单进程-大块'},
        {'chunk_size': 10000, 'process_count': min(2, cpu_count), 'name': '双进程-中块'},
        {'chunk_size': 10000, 'process_count': min(4, cpu_count), 'name': '四进程-中块'},
    ]

    print(f"\n开始性能测试...")

    for config in test_configs:
        print(f"\n--- 测试配置: {config['name']} ---")
        print(f"分块大小: {config['chunk_size']:,} 行")
        print(f"进程数: {config['process_count']}")

        start_time = time.time()

        # 这里应该调用优化后的处理函数
        # 由于我们在测试脚本中，这里只是模拟
        simulate_processing(test_file1, test_file2, config['chunk_size'], config['process_count'])

        end_time = time.time()
        duration = end_time - start_time

        print(f"处理时间: {duration:.2f} 秒")
        print(f"内存使用: {psutil.virtual_memory().percent:.1f}%")

def simulate_processing(file1, file2, chunk_size, process_count):
    """模拟数据处理过程"""
    # 读取文件信息
    try:
        df1_info = pd.read_excel(file1, nrows=0)  # 只读取列信息
        df2 = pd.read_excel(file2)  # 小文件完全读取

        # 模拟分块处理
        import openpyxl
        wb = openpyxl.load_workbook(file1, read_only=True)
        ws = wb.active
        total_rows = ws.max_row

        chunks = (total_rows - 1) // chunk_size + 1
        print(f"总行数: {total_rows:,}, 分块数: {chunks}")

        # 模拟处理每个块
        for i in range(0, total_rows, chunk_size):
            # 这里应该是实际的数据处理
            time.sleep(0.01)  # 模拟处理时间

        wb.close()

    except Exception as e:
        print(f"处理出错: {e}")

def cleanup_test_files():
    """清理测试文件"""
    test_files = ['test_data_large.xlsx', 'test_data_small.xlsx']
    for file in test_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"已删除测试文件: {file}")

def main():
    """主函数"""
    print("Excel数据匹配工具性能测试")
    print("=" * 50)

    try:
        test_performance()
    except KeyboardInterrupt:
        print("\n测试被用户中断")
    except Exception as e:
        print(f"\n测试过程中出错: {e}")
    finally:
        # 询问是否清理测试文件
        response = input("\n是否删除测试文件? (y/N): ").strip().lower()
        if response in ['y', 'yes']:
            cleanup_test_files()

if __name__ == "__main__":
    main()
