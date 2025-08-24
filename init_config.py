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
    "window_position": {
        "x": 658,
        "y": 154
    },
    "sidebar_position": {
        "relative_offset_x": 0,
        "relative_offset_y": 0,
        "adsorbed_edge": ""
    },
    "sidebar_relative_x": -5,
    "sidebar_relative_y": 400,
    "window_width": 300,
    "window_height": 400,
    "window_inner_margin_width": 12,
    "window_inner_margin_height": 12,
    "window_screen_margins": 10,
    "sidebar_margin": 20,
    "sidebar_width": 200,
    "marks": {}
    }
    with open("config.json", "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)
    print("config.json 已初始化")

if __name__ == "__main__":
    init_config()
