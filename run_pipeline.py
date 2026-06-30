import os
import sys
import glob
import argparse

# 添加项目根目录到 sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core_engine.action_dispatcher import PipelineEngine

def run_single_config(config_path):
    """执行单个 YAML 配置"""
    if not os.path.exists(config_path):
        print(f"❌ 错误: 找不到配置文件 {config_path}")
        return False
        
    try:
        engine = PipelineEngine(config_path)
        engine.run()
        return True
    except Exception as e:
        print(f"❌ 执行 {config_path} 失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def run_all_configs(config_dir="configs"):
    """扫描 configs 目录下所有的 yaml 文件并全部执行"""
    yaml_files = glob.glob(os.path.join(config_dir, "*.yaml"))
    
    if not yaml_files:
        print(f"⚠️ 在 {config_dir} 目录下没有找到任何 YAML 配置文件。")
        return
        
    print(f"🔍 发现 {len(yaml_files)} 个任务配置，准备批量执行...")
    
    success_count = 0
    for config_path in yaml_files:
        if run_single_config(config_path):
            success_count += 1
            
    print(f"\n🎉 批量执行完毕！成功: {success_count}/{len(yaml_files)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="行业数据库自动化生成 Pipeline")
    parser.add_argument("-c", "--config", type=str, help="指定要执行的 YAML 配置文件路径")
    parser.add_argument("-a", "--all", action="store_true", help="执行 configs 目录下的所有配置")
    
    args = parser.parse_args()
    
    if args.config:
        run_single_config(args.config)
    elif args.all:
        run_all_configs()
    else:
        # 默认行为：提示用法
        print("请指定要执行的任务！")
        print("执行单个任务: python run_pipeline.py -c configs/db_aviation.yaml")
        print("执行全部任务: python run_pipeline.py --all")