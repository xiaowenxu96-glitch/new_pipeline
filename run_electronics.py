import os
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core_engine.action_dispatcher import PipelineEngine

if __name__ == "__main__":
    config_path = os.path.join("configs", "db_electronics.yaml")
    engine = PipelineEngine(config_path)
    engine.run()
