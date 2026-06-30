import os
import sys

# 添加项目根目录到 sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core_engine.action_dispatcher import PipelineEngine

if __name__ == "__main__":
    config_path = os.path.join("configs", "db_aviation.yaml")
    engine = PipelineEngine(config_path)
    engine.run()
