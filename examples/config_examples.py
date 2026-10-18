# -*- coding: utf-8 -*-
"""
config_examples.py — 多学科配置示例

score_assigner.py 的顶部配置是模块级常量。切换学科 / 课程时，
既可以直接改 score_assigner.py 顶部，也可以在本文件里新增一套配置后调用：

    from examples.config_examples import CONFIGS, apply_config
    import score_assigner as sa
    apply_config(sa, CONFIGS["理工科_平时70_期末30"])
    sa.run(sa.FILE_PATH)
"""

CONFIGS = {
    # 示例表（examples/demo_scores.xlsx）所用的默认配置
    "默认_平时80_出勤20": {
        "FILE_PATH": r"examples/demo_scores.xlsx",
        "SHEET_NAME": "Sheet1",
        "DATA_START_ROW": 2,
        "ID_COL": "B",
        "NAME_COL": "C",
        "ATTEND_COL": "H",
        "HW_COL": "J",
        "TARGET_COL": "L",
        "HW_MIN": 60,
        "HW_MAX": 99,
        "WEIGHT_I": 0.8,
        "WEIGHT_J": 0.2,
        "J_VALUE": 100,
    },
    # 理工科常见口径：平时分权重 70%，出勤 30%
    "理工科_平时70_出勤30": {
        "FILE_PATH": r"examples/demo_scores.xlsx",
        "SHEET_NAME": "Sheet1",
        "DATA_START_ROW": 2,
        "ATTEND_COL": "H",
        "HW_COL": "J",
        "TARGET_COL": "L",
        "HW_MIN": 60,
        "HW_MAX": 100,
        "WEIGHT_I": 0.7,
        "WEIGHT_J": 0.3,
        "J_VALUE": 100,
    },
    # 实践类课程：平时分区间放宽到 50~100
    "实践类_区间50至100": {
        "FILE_PATH": r"examples/demo_scores.xlsx",
        "SHEET_NAME": "Sheet1",
        "DATA_START_ROW": 2,
        "ATTEND_COL": "H",
        "HW_COL": "J",
        "TARGET_COL": "L",
        "HW_MIN": 50,
        "HW_MAX": 100,
        "WEIGHT_I": 0.6,
        "WEIGHT_J": 0.4,
        "J_VALUE": 100,
    },
}


def apply_config(module, config: dict) -> None:
    """把一套配置写进 score_assigner 模块的同名常量。"""
    for key, value in config.items():
        if not hasattr(module, key):
            raise AttributeError(f"score_assigner 中不存在配置项 {key!r}")
        setattr(module, key, value)


def reachable_range(config: dict) -> tuple:
    """给定配置下的总评可达区间（出勤按 J_VALUE 常满算，取整口径同 Excel ROUND）。"""
    i, j = config["WEIGHT_I"], config["WEIGHT_J"]

    def _round(x: float) -> int:   # 半进位，避免 Python 银行家舍入
        return int(x + 0.5) if x >= 0 else -int(-x + 0.5)

    lo = _round(i * config["HW_MIN"] + j * config["J_VALUE"])
    hi = _round(i * config["HW_MAX"] + j * config["J_VALUE"])
    return lo, hi


if __name__ == "__main__":
    for name, cfg in CONFIGS.items():
        lo, hi = reachable_range(cfg)
        print(f"{name:24s} 平时分 [{cfg['HW_MIN']}, {cfg['HW_MAX']}]  →  可达总评 [{lo}, {hi}]")
