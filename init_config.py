import json
from pathlib import Path

def init_config():
    config = {
        "RootFolder": str(Path.home()),
        "window": {
            "width": 400,
            "height": 400,
            "inner_margin_width": 12,
            "inner_margin_height": 12,
            "screen_margins": 10
        },
        "window_position": {"x": 100, "y": 100},
        "sidebar_relative_x": 420,
        "sidebar_relative_y": 0,
    }
    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)
    print("config.json 已初始化")

if __name__ == "__main__":
    init_config()
