import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(BASE_DIR)

from core_engine.action_dispatcher import PipelineEngine

if __name__ == "__main__":
    config_path = os.path.join(BASE_DIR, "configs", "db_baijiu2.yaml")
    engine = PipelineEngine(config_path)
    engine.run()
