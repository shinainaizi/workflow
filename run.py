#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
社区活动台账自动化工作流 - 一键运行入口

工作流程:
1. 读取策划方案(.docx) + 现场照片 + 回忆记录
2. AI提取结构化信息
3. 生成活动台账(.docx)
4. 生成汇总报告

使用方法:
    python run.py --scheme 策划方案.docx --photos 照片文件夹 --notes 回忆记录.txt
    python run.py --batch 批量处理文件夹
"""

import argparse
import os
import sys
import yaml
from pathlib import Path

def safe_print(msg):
    """Safe print that handles encoding issues on Windows"""
    try:
        print(msg)
    except UnicodeEncodeError:
        safe_msg = msg.encode('ascii', 'ignore').decode('ascii')
        if safe_msg.strip():
            print(safe_msg)
        else:
            print("[Non-ASCII content]")

# 将core目录加入Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'core'))

from extract import ActivityExtractor
from generate_ledger import LedgerGenerator
from generate_report import ReportGenerator

def load_config():
    """加载配置文件"""
    config_path = Path(__file__).parent / 'config.yaml'
    if config_path.exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    return {}

def process_single_activity(scheme_path, photos_dir, notes_path, config):
    """处理单个活动"""
    safe_print("=" * 60)
    safe_print("开始处理活动台账自动生成...")
    safe_print("=" * 60)
    
    # Step 1: 提取结构化信息
    safe_print("\n[1/4] 正在从策划方案中提取信息...")
    extractor = ActivityExtractor(config)
    activity_data = extractor.extract(scheme_path, photos_dir, notes_path)
    safe_print(f"  提取完成: {activity_data.get('activity_theme', '未知活动')}")
    
    # Step 2: 生成台账
    safe_print("\n[2/4] 正在生成活动台账...")
    ledger_gen = LedgerGenerator(config)
    ledger_path = ledger_gen.generate(activity_data)
    safe_print(f"  台账已生成: {ledger_path}")
    
    # Step 3: 生成汇总报告（如果有多个台账）
    safe_print("\n[3/4] 正在更新汇总报告...")
    report_gen = ReportGenerator(config)
    report_path = report_gen.update_report(activity_data, ledger_path)
    safe_print(f"  报告已更新: {report_path}")
    
    safe_print("\n" + "=" * 60)
    safe_print("处理完成！")
    safe_print(f"台账文件: {ledger_path}")
    safe_print(f"报告文件: {report_path}")
    safe_print("=" * 60)
    
    return ledger_path, report_path

def batch_process(batch_dir, config):
    """批量处理多个活动"""
    batch_path = Path(batch_dir)
    results = []
    
    # 查找所有包含策划方案的子目录
    for subdir in batch_path.iterdir():
        if subdir.is_dir():
            scheme_files = list(subdir.glob('*.docx'))
            if scheme_files:
                scheme_path = scheme_files[0]  # 取第一个docx作为方案
                photos_dir = subdir / 'photos' if (subdir / 'photos').exists() else None
                notes_path = subdir / 'notes.txt' if (subdir / 'notes.txt').exists() else None
                
safe_print(f"\n正在处理: {subdir.name}")
                try:
                    ledger_path, report_path = process_single_activity(
                        str(scheme_path), 
                        str(photos_dir) if photos_dir else None,
                        str(notes_path) if notes_path else None,
                        config
                    )
                    results.append((subdir.name, ledger_path, report_path))
                except Exception as e:
                    safe_print(f"  处理失败: {e}")
                    results.append((subdir.name, None, str(e)))
    
    return results

def main():
    parser = argparse.ArgumentParser(
        description='社区活动台账自动化工作流',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python run.py --scheme 活动方案.docx --photos ./photos --notes 回忆.txt
  python run.py --batch ./batch_input
        """
    )
    
    parser.add_argument('--scheme', '-s', help='活动策划方案文件路径(.docx)')
    parser.add_argument('--photos', '-p', help='现场照片文件夹路径')
    parser.add_argument('--notes', '-n', help='回忆记录文本文件路径(.txt)')
    parser.add_argument('--batch', '-b', help='批量处理目录（包含多个活动子文件夹）')
    parser.add_argument('--output', '-o', default='outputs', help='输出目录')
    
    args = parser.parse_args()
    
    # 加载配置
    config = load_config()
    if args.output:
        config['output']['ledger_dir'] = os.path.join(args.output, 'ledger')
        config['output']['report_dir'] = os.path.join(args.output, 'report')
    
    # 检查API配置
    if not config.get('qianfan', {}).get('api_key') or config['qianfan']['api_key'] == 'your-api-key-here':
        safe_print("错误: 请先配置百度千帆API密钥！")
        safe_print("请编辑 config.yaml 文件，填入你的 API Key 和 Secret Key")
        safe_print("获取地址: https://console.bce.baidu.com/qianfan/ais/console/applicationConsole/application")
        sys.exit(1)
    
    try:
        if args.batch:
            # 批量模式
            results = batch_process(args.batch, config)
            safe_print("\n" + "=" * 60)
            safe_print("批量处理完成，结果汇总:")
            for name, ledger, report in results:
                status = "成功" if ledger else f"失败: {report}"
                safe_print(f"  {name}: {status}")
        elif args.scheme:
            # 单文件模式
            process_single_activity(args.scheme, args.photos, args.notes, config)
        else:
            safe_print("请提供 --scheme 参数指定策划方案，或使用 --batch 批量处理")
            parser.print_help()
            
    except Exception as e:
        safe_print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
