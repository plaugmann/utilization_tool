from pathlib import Path
from typing import Dict

import yaml

BASE_PATH = Path(__file__).parents[1]
CONFIG_PATH = BASE_PATH / "config"


def load_app_config() -> Dict:
    path = CONFIG_PATH / "app.yaml"
    if path.exists():
        with open(path, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    return {}
