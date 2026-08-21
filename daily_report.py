#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
daily_report.py
===============
教育出版 · 学科销量日报工具（单文件版）
-----------------------------------------------
一个文件搞定「Quick BI 明细表 → 目标达成分析 → 结构化/智能体报告」。
既可以在 Codex 里用，也可以在 WorkBuddy 里用；分析逻辑完全相同，
唯一平台相关的只有「从 Quick BI 下载 Excel」这一步。

=====================================================================
一、不同大模型工具所需的浏览器插件（⚠️ 重点）
=====================================================================
「分析 + 报告」部分与平台无关；唯一平台相关的是**从 Quick BI 下载明细表 Excel**。
这一步在不同大模型工具里用的浏览器自动化能力不同：

| 工具       | 下载数据用的浏览器能力                                | 是否需要额外安装插件                |
|------------|-------------------------------------------------------|-------------------------------------|
| Codex      | OpenAI 内置 Browser 工具（对话即可驱动浏览器）         | 否，运行时自带，自动复用登录态      |
| WorkBuddy  | Playwright MCP（`mcp__playwright__browser_*`）或八爪鱼连接器 | 是，需先在 WorkBuddy 设置里连接 Playwright MCP / 八爪鱼 |

两者最终都只产出一个本地 Excel 文件，后续分析逻辑 100% 相同。
如果暂时不想接浏览器插件，两种工具都支持「本地文件模式」：
手动从 Quick BI 导出 Excel 后，用 `--source` 直接指向该文件即可（见下方用法）。

=====================================================================
二、怎么用（两种大模型工具都适用）
=====================================================================
【Codex 里】
  直接把本文件放进对话/仓库，然后用自然语言说，例如：
  “用 daily_report.py 帮我下载并生成 26暑 销量日报”
  → Codex 会读取本文件顶部的说明，用内置浏览器从 Quick BI 导出 Excel，
    再运行 `python daily_report.py --source <导出的xlsx> --current 26暑` 产出报告。

【WorkBuddy 里】
  1) 先在 WorkBuddy 设置里连接「Playwright MCP」（或八爪鱼连接器）；
  2) 把本文件放进项目目录，用自然语言说，例如：
     “按 daily_report.py 的说明，用 Playwright 从 Quick BI 下载 Excel 并生成 26暑 日报”
  → WorkBuddy 按本文件里的 Playwright 操作清单完成下载，再运行脚本产出报告。

【通用 / 本地文件模式（两种工具都支持）】
  如果你已经手动导出了 Excel：
  python daily_report.py --source 平台数据明细_明细表.xlsx --current 26暑

=====================================================================
三、依赖
=====================================================================
  Python 3.10+ ，仅需 pandas + openpyxl：
  pip install pandas openpyxl

=====================================================================
四、目标与口径（看哪一学期，就对照哪一学期的目标）
=====================================================================
  - 小学数学 = 小数同步 + 小数培优（原始宽表只到“小学数学”粒度，故实际销量按
    “小学数学”合计与“小数同步+小数培优”合并目标对照；如需表内拆分，请在
    SUBJECT_TARGET_KEYS / business_subject() 处按你的 Quick BI 实际维度调整）
  - 语文     = 小学语文 + 初中语文
  - 英语     = 小学英语 + 初中英语
  - 大盘     = 上述全部学科合计
  目标值内嵌在下方 TARGETS 中，按学期给出，可直接修改。

=====================================================================
五、智能体驱动（不死板）
=====================================================================
  本工具只产出「事实数据」（JSON / 文本 / HTML），真正“数据分析 → 机会问题”
  的叙述由调用它的智能体（Codex / WorkBuddy）根据数据实际情况生成：
  - 用 `--render agent_prompt` 可导出一段写给智能体的分析 prompt（含完整数据 +
    分析任务 + 输出格式要求），由智能体续写日报正文；
  - 不带 LLM 时，`--render text` / `--render html` 提供可读的事实兜底报表。
=====================================================================
"""

from __future__ import annotations

import argparse
import html
import json
import sys
from pathlib import Path
from typing import Optional

import pandas as pd


__version__ = "1.0.0"


# =====================================================================
# 配置区（可自由修改）
# =====================================================================

# 各学期销量目标（单位：本）。学期键与 Quick BI 列名一致；看哪一学期就对照哪一学期。
# 来源：业务看板截图。小学数学=小数同步+小数培优；语文=小学语文+初中语文；
# 英语=小学英语+初中英语；大盘=全部合计。
TARGETS = {
    "subject_targets": {
        "1小数同步": {"25春": 307470, "25暑": 270611, "25秋": 334518, "26春": 327058, "26暑": 308930, "26秋": 377212},
        "2小数培优": {"25春": 224929, "25暑": 175476, "25秋": 226738, "26春": 217970, "26暑": 188356, "26秋": 243464},
        "3初中数学": {"25春": 123848, "25暑": 212603, "25秋": 183772, "26春": 139992, "26暑": 252792, "26秋": 218346},
        "4小学语文": {"25春": 296049, "25暑": 237373, "25秋": 278766, "26春": 248423, "26暑": 223603, "26秋": 257772},
        "5初中语文": {"25春": 30491,  "25暑": 70628,  "25秋": 42890,  "26春": 30150,  "26暑": 73218,  "26秋": 46318},
        "6小学英语": {"25春": 118191, "25暑": 140708, "25秋": 155765, "26春": 123370, "26暑": 173250, "26秋": 187143},
        "7初中英语": {"25春": 38401,  "25暑": 70044,  "25秋": 54257,  "26春": 39913,  "26暑": 74291,  "26秋": 54992},
        "8初中理化": {"25春": 25508,  "25暑": 63397,  "25秋": 40233,  "26春": 27734,  "26暑": 64163,  "26秋": 41044},
        "9其他":     {"25春": 1145,   "25暑": 9803,   "25秋": 2339,   "26春": 342,    "26暑": 1502,   "26秋": 1030},
    },
    "group_targets": {
        "小学数学": {"members": ["1小数同步", "2小数培优"], "25春": 532399, "25暑": 446087, "25秋": 561256, "26春": 545028, "26暑": 497286, "26秋": 620676},
        "语文":     {"members": ["4小学语文", "5初中语文"], "25春": 326540, "25暑": 308001, "25秋": 321656, "26春": 278573, "26暑": 296821, "26秋": 304090},
        "英语":     {"members": ["6小学英语", "7初中英语"], "25春": 156592, "25暑": 210752, "25秋": 210022, "26春": 163283, "26暑": 247541, "26秋": 242135},
        "大盘":     {"members": ["1小数同步", "2小数培优", "3初中数学", "4小学语文", "5初中语文", "6小学英语", "7初中英语", "8初中理化", "9其他"],
                     "25春": 1166032, "25暑": 1250643, "25秋": 1319278, "26春": 1154952, "26暑": 1360105, "26秋": 1427322},
    },
}

# 原始 Quick BI 明细表字段（与导出的 Excel 列名保持一致）
INST_ID = "机构id"
INST_NAME = "机构名"
REGION = "大区"
PROVINCE = "省份"
CITY = "城市"
OPERATOR = "运营人"
SUBJECT = "学科"
SYSTEM = "体系"
GRADE = "年级"
VERSION = "版本"
COOP_DATE = "合作日期"

# 学期列（按时间顺序，与源表保持一致；仅有秋学期带“同期”列）
ALL_SEMESTERS = ["23秋", "24暑", "24秋", "25春", "25暑", "25秋", "26寒", "26春", "26暑", "26秋"]
YOY_COL = {  # 当前学期 -> 同期列（用于同比）；非秋学期无同期列
    "23秋": "23秋同期", "24秋": "24秋同期", "25秋": "25秋同期",
}

# 清洗规则
DROP_REGIONS = {"其他", "代理"}

# 新签口径：合作日期 ∈ [NEW_SIGN_START, NEW_SIGN_END) 且本期销量 > 20
NEW_SIGN_START = 20250101
NEW_SIGN_END = 20260101

# 默认分析期
DEFAULT_CURRENT = "26暑"
DEFAULT_PRIOR = "26春"
DEFAULT_YOY = "25暑"

# 原始 学科 -> 业务口径（9 类）。注意：原始宽表无“小数同步/小数培优”细分，
# 故 小学数学 作为合并口径，与目标“小数同步+小数培优”之和对照。
def business_subject(s: object) -> str:
    s = str(s)
    if s == "小学数学":
        return "小学数学"
    if s == "初中数学":
        return "初中数学"
    if s == "小学语文":
        return "小学语文"
    if s == "初中语文":
        return "初中语文"
    if s == "小学英语":
        return "小学英语"
    if s == "初中英语":
        return "初中英语"
    if s in ("初中物理", "初中化学"):
        return "初中理化"
    return "其他"


BUSINESS_SUBJECTS = ["小学数学", "初中数学", "小学语文", "初中语文", "小学英语", "初中英语", "初中理化", "其他"]

# 业务口径 -> 对应 TARGETS["subject_targets"] 的键（用于逐科对照目标）
SUBJECT_TARGET_KEYS = {
    "小学数学": ["1小数同步", "2小数培优"],
    "初中数学": ["3初中数学"],
    "小学语文": ["4小学语文"],
    "初中语文": ["5初中语文"],
    "小学英语": ["6小学英语"],
    "初中英语": ["7初中英语"],
    "初中理化": ["8初中理化"],
    "其他": ["9其他"],
}

# 分组 -> 原始 学科 列表（用于按原始数据汇总实际销量）
GROUP_RAW_SUBJECTS = {
    "小学数学": ["小学数学"],
    "语文": ["小学语文", "初中语文"],
    "英语": ["小学英语", "初中英语"],
}


# =====================================================================
# 下载适配层（平台相关部分：仅此处不同）
# =====================================================================

def build_download_recipe(platform: str, report_url: str, save_dir: str = "./downloads") -> list:
    """返回给智能体执行的浏览器操作清单（平台相关）。

    platform: "codex" 或 "workbuddy"
      - codex：使用 OpenAI 内置浏览器工具（运行时自带，无需安装插件）
      - workbuddy：使用 Playwright MCP（需先在 WorkBuddy 连接 Playwright MCP / 八爪鱼）
    """
    if platform == "codex":
        return [
            {"step": 1, "action": "open", "url": report_url, "note": "Codex 内置浏览器，自动复用已登录的 Quick BI 会话"},
            {"step": 2, "action": "wait_for", "selector": ".quickbi-report, [data-testid='report-canvas']"},
            {"step": 3, "action": "click", "selector": "button:has-text('导出'), button:has-text('下载'), .icon-export", "note": "UI 文案可能不同，必要时截图确认"},
            {"step": 4, "action": "click", "selector": "span:has-text('Excel'), .download-excel"},
            {"step": 5, "action": "download", "save_dir": save_dir, "expected_pattern": "平台数据明细_明细表_*.xlsx"},
            {"step": 6, "action": "validate", "row_count_min": 1000, "column_count_min": 10},
        ]
    # workbuddy —— 通过 Playwright MCP 工具完成
    return [
        {"step": 1, "tool": "mcp__playwright__browser_navigate", "params": {"url": report_url}, "note": "需先在 WorkBuddy 连接 Playwright MCP；复用已登录会话"},
        {"step": 2, "tool": "mcp__playwright__browser_wait_for", "params": {"selector": ".quickbi-report, [data-testid='report-canvas']", "timeout": 30000}},
        {"step": 3, "tool": "mcp__playwright__browser_click", "params": {"selector": "button:has-text('导出'), button:has-text('下载'), .icon-export"}},
        {"step": 4, "tool": "mcp__playwright__browser_click", "params": {"selector": "span:has-text('Excel'), .download-excel"}},
        {"step": 5, "tool": "mcp__playwright__browser_wait_for", "params": {"state": "download", "timeout": 30000}, "note": "等待下载完成，文件名形如 平台数据明细_明细表_*.xlsx"},
        {"step": 6, "action": "move_download_to", "save_dir": save_dir},
    ]


# =====================================================================
# 分析层（平台无关）
# =====================================================================

def load_excel(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"源文件不存在：{path}")
    return pd.read_excel(path, engine="openpyxl")


def _to_biz_frame(df: pd.DataFrame) -> pd.DataFrame:
    """原始宽表 -> 加业务口径列、学期列转数值的清洗后行级表。"""
    df = df.copy()
    semesters = [s for s in ALL_SEMESTERS if s in df.columns]
    for col in semesters:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    dim_cols = [INST_ID, INST_NAME, REGION, PROVINCE, CITY, OPERATOR, SUBJECT, SYSTEM, GRADE, VERSION]
    for col in dim_cols:
        if col in df.columns:
            df[col] = df[col].fillna("缺失").astype(str)

    if COOP_DATE in df.columns:
        df[COOP_DATE] = pd.to_numeric(df[COOP_DATE], errors="coerce")

    if REGION in df.columns:
        df = df[~df[REGION].isin(DROP_REGIONS)].copy()

    if INST_ID in df.columns:
        df[INST_ID] = df[INST_ID].astype(str)
    else:
        raise ValueError(f"源表缺少必要字段：{INST_ID}")

    df["_biz"] = df[SUBJECT].map(business_subject)
    return df


def collapse_orgs(df: pd.DataFrame) -> pd.DataFrame:
    """按机构 ID 汇总为机构级表（销量求和，维度取首条非空）。"""
    semesters = [s for s in ALL_SEMESTERS if s in df.columns]
    grouped = df.groupby(INST_ID, as_index=False)
    sums = grouped[semesters].sum()
    meta_cols = [c for c in [INST_ID, INST_NAME, REGION, PROVINCE, CITY, OPERATOR, SUBJECT, SYSTEM, GRADE, VERSION, "_biz"] if c in df.columns]
    meta = grouped.first()[meta_cols].set_index(INST_ID)
    if COOP_DATE in df.columns:
        first_coop = grouped[COOP_DATE].min().reset_index().rename(columns={COOP_DATE: "_first_coop_date"})
        sums = sums.merge(first_coop, on=INST_ID, how="left")
    org_df = sums.merge(meta, left_on=INST_ID, right_index=True, how="left")
    return org_df


def flag_org_types(org_df: pd.DataFrame, current: str, prior: str) -> pd.DataFrame:
    """标记新签 / 老机构 / 续用 / 流失。"""
    df = org_df.copy()
    coop = pd.to_numeric(df.get("_first_coop_date"), errors="coerce").fillna(0).astype(int)
    df["_is_new"] = (coop >= NEW_SIGN_START) & (coop < NEW_SIGN_END) & (df[current] > 20)
    df["_is_old"] = coop < NEW_SIGN_START
    cur_pos = df[current] > 20
    prior_pos = df[prior] > 20
    ratio = df[current] / df[prior].replace(0, float("nan"))
    df["_is_renew"] = df["_is_old"] & prior_pos & cur_pos & (ratio >= 0.8)
    df["_is_churn"] = df["_is_old"] & prior_pos & ~df["_is_renew"]
    return df


def pct_change(current: float, base: float) -> Optional[float]:
    if base == 0 or pd.isna(base):
        return None
    return (current - base) / base


def achievement_rate(actual: float, target: float) -> Optional[float]:
    if target == 0 or pd.isna(target):
        return None
    return actual / target


def org_summary(scope_raw: pd.DataFrame, org_df: pd.DataFrame, current: str, prior: str, yoy: str) -> dict:
    """计算某一范围（原始行子集）的核心指标。"""
    cur_sales = int(scope_raw[current].sum())
    prior_sales = int(scope_raw[prior].sum())
    yoy_sales = int(scope_raw[yoy].sum()) if yoy in scope_raw.columns else 0

    orgs_in_scope = scope_raw[INST_ID].unique()
    sub_org = org_df[org_df[INST_ID].isin(orgs_in_scope)]

    new_orgs = int(sub_org["_is_new"].sum())
    old_orgs = int(sub_org["_is_old"].sum())
    renew_orgs = int(sub_org["_is_renew"].sum())
    churn_orgs = int(sub_org["_is_churn"].sum())
    renewal_rate = (renew_orgs / (renew_orgs + churn_orgs)) if (renew_orgs + churn_orgs) > 0 else None

    return {
        "sales": {
            "current": cur_sales,
            "prior": prior_sales,
            "yoy": yoy_sales,
            "mom_change": pct_change(cur_sales, prior_sales),
            "yoy_change": pct_change(cur_sales, yoy_sales),
        },
        "orgs": {"current": int((scope_raw[current] > 0).sum()), "prior": int((scope_raw[prior] > 0).sum())},
        "new_sign": {"orgs": new_orgs, "sales": int(sub_org[current].sum())},
        "old": {"orgs": old_orgs, "renew_orgs": renew_orgs, "churn_orgs": churn_orgs, "renewal_rate": renewal_rate},
    }


def aggregate_dim(raw_flagged: pd.DataFrame, dim: str, current: str, prior: str, yoy: str, limit: Optional[int] = None) -> list:
    """按任意维度聚合（销量 + 新签/续用/流失机构数），从行级表计算，保证体系/年级/版本等可正确拆分。"""
    if dim not in raw_flagged.columns:
        return []
    rows = []
    flag_cols = ["_is_new", "_is_renew", "_is_churn"]
    for key, sub in raw_flagged.groupby(dim, dropna=False):
        cur_sales = int(sub[current].sum())
        prior_sales = int(sub[prior].sum())
        yoy_sales = int(sub[yoy].sum()) if yoy in sub.columns else 0
        new_orgs = int(sub["_is_new"].sum())
        renew_orgs = int(sub["_is_renew"].sum())
        churn_orgs = int(sub["_is_churn"].sum())
        renewal_rate = (renew_orgs / (renew_orgs + churn_orgs)) if (renew_orgs + churn_orgs) > 0 else None
        rows.append({
            "dimension": dim, "value": str(key),
            "sales_current": cur_sales, "sales_prior": prior_sales, "sales_yoy": yoy_sales,
            "sales_delta": cur_sales - prior_sales,
            "mom_change": pct_change(cur_sales, prior_sales),
            "yoy_change": pct_change(cur_sales, yoy_sales),
            "new_orgs": new_orgs, "renew_orgs": renew_orgs, "churn_orgs": churn_orgs,
            "renewal_rate": renewal_rate,
        })
    rows.sort(key=lambda x: x["sales_current"], reverse=True)
    if limit:
        rows = rows[:limit]
    return rows


def subject_actual(raw: pd.DataFrame, biz: str, current: str) -> int:
    return int(raw[raw["_biz"] == biz][current].sum())


def subject_target_value(biz: str, sem: str) -> Optional[int]:
    keys = SUBJECT_TARGET_KEYS.get(biz, [])
    total = 0
    found = False
    for k in keys:
        v = TARGETS["subject_targets"].get(k, {}).get(sem)
        if v is not None:
            total += v
            found = True
    return total if found else None


def build_report_data(source: str | Path, current: str = DEFAULT_CURRENT, prior: str = DEFAULT_PRIOR, yoy: str = DEFAULT_YOY) -> dict:
    """构建完整报告数据（只输出事实，不含文本结论）。"""
    raw = load_excel(source)
    raw = _to_biz_frame(raw)
    org_df = collapse_orgs(raw)
    org_df = flag_org_types(org_df, current=current, prior=prior)

    # 把机构级标记合并回行级，便于按维度统计新签/续用/流失机构数
    flag_cols = org_df[[INST_ID, "_is_new", "_is_old", "_is_renew", "_is_churn"]]
    raw_flagged = raw.merge(flag_cols, on=INST_ID, how="left")

    report = {
        "meta": {
            "source_file": str(Path(source).name),
            "source_path": str(Path(source).absolute()),
            "generated_at": pd.Timestamp.now().isoformat(),
            "current_semester": current, "prior_semester": prior, "yoy_semester": yoy,
            "row_count_raw": int(len(raw)),
        },
        "overall": {}, "groups": {}, "business_subjects": {}, "dimensions": {}, "facts": {},
    }

    # 大盘
    overall = org_summary(raw_flagged, org_df, current, prior, yoy)
    tgt = TARGETS["group_targets"]["大盘"].get(current)
    if tgt is not None:
        actual = overall["sales"]["current"]
        overall["target"] = {"semester": current, "value": int(tgt),
                             "achievement_rate": achievement_rate(actual, tgt), "gap": actual - int(tgt),
                             "gap_pct": pct_change(actual, tgt)}
    report["overall"] = overall

    # 分组（小学数学 / 语文 / 英语）
    for gname, members in GROUP_RAW_SUBJECTS.items():
        scope = raw_flagged[raw_flagged[SUBJECT].isin(members)]
        if scope.empty:
            continue
        grp = org_summary(scope, org_df, current, prior, yoy)
        gt = TARGETS["group_targets"][gname].get(current)
        if gt is not None:
            actual = grp["sales"]["current"]
            grp["target"] = {"semester": current, "value": int(gt),
                             "achievement_rate": achievement_rate(actual, gt), "gap": actual - int(gt),
                             "gap_pct": pct_change(actual, gt)}
        report["groups"][gname] = {
            "summary": grp,
            "by_system": aggregate_dim(scope, SYSTEM, current, prior, yoy),
            "by_grade": aggregate_dim(scope, GRADE, current, prior, yoy),
            "by_version": aggregate_dim(scope, VERSION, current, prior, yoy),
            "by_region": aggregate_dim(scope, REGION, current, prior, yoy),
            "by_operator": aggregate_dim(scope, OPERATOR, current, prior, yoy, limit=30),
        }

    # 业务口径逐科对照目标
    for biz in BUSINESS_SUBJECTS:
        actual = subject_actual(raw_flagged, biz, current)
        tgt = subject_target_value(biz, current)
        item = {"actual": actual, "target": tgt}
        if tgt is not None:
            item["achievement_rate"] = achievement_rate(actual, tgt)
            item["gap"] = actual - tgt
        report["business_subjects"][biz] = item

    # 大盘维度下钻
    report["dimensions"] = {
        "by_subject": aggregate_dim(raw_flagged, SUBJECT, current, prior, yoy),
        "by_biz": aggregate_dim(raw_flagged, "_biz", current, prior, yoy),
        "by_system": aggregate_dim(raw_flagged, SYSTEM, current, prior, yoy),
        "by_region": aggregate_dim(raw_flagged, REGION, current, prior, yoy),
        "by_operator": aggregate_dim(raw_flagged, OPERATOR, current, prior, yoy, limit=30),
    }

    report["facts"] = _extract_facts(report)
    return report


def _extract_facts(report: dict) -> dict:
    facts = {
        "overall_achievement": report["overall"].get("target", {}).get("achievement_rate"),
        "worst_achievement_group": None, "best_achievement_group": None,
        "worst_sales_delta_biz": None, "best_sales_delta_biz": None,
        "worst_system_delta": None, "best_system_delta": None,
    }
    group_rates = []
    for name, data in report["groups"].items():
        tgt = data["summary"].get("target")
        if tgt and tgt.get("achievement_rate") is not None:
            group_rates.append((name, tgt["achievement_rate"], tgt["gap"]))
    if group_rates:
        group_rates.sort(key=lambda x: x[1])
        facts["worst_achievement_group"] = {"name": group_rates[0][0], "achievement_rate": group_rates[0][1], "gap": group_rates[0][2]}
        facts["best_achievement_group"] = {"name": group_rates[-1][0], "achievement_rate": group_rates[-1][1], "gap": group_rates[-1][2]}

    by_biz = report["dimensions"].get("by_biz", [])
    deltas = [r for r in by_biz if r.get("sales_delta") is not None]
    if deltas:
        deltas.sort(key=lambda x: x["sales_delta"])
        facts["worst_sales_delta_biz"] = deltas[0]
        facts["best_sales_delta_biz"] = deltas[-1]

    by_sys = report["dimensions"].get("by_system", [])
    sd = [r for r in by_sys if r.get("sales_delta") is not None]
    if sd:
        sd.sort(key=lambda x: x["sales_delta"])
        facts["worst_system_delta"] = sd[0]
        facts["best_system_delta"] = sd[-1]
    return facts


# =====================================================================
# 报告层（智能体驱动 + 兜底）
# =====================================================================

def build_agent_prompt(report_data: dict, focus: Optional[str] = None) -> str:
    """生成给智能体/LLM 的分析 prompt（含完整数据 + 任务 + 格式要求）。"""
    meta = report_data.get("meta", {})
    facts = report_data.get("facts", {})

    def _rate(v):
        return "无目标/未计算" if v is None else f"{v * 100:.1f}%"

    def _ach(item):
        if not item:
            return "无"
        return f"{item['name']}（达成率 {item['achievement_rate'] * 100:.1f}%，缺口 {item['gap']:,} 本）"

    def _delta(item):
        if not item:
            return "无"
        sign = "+" if item.get("sales_delta", 0) >= 0 else ""
        return f"{item['value']}（{item['dimension']}）：{item['sales_current']:,} 本，环比 {sign}{item['sales_delta']:,} 本"

    prompt = f"""你是一位资深教育出版业务分析师。请根据以下销售数据生成一份日报。

## 分析任务
1. 先说明大盘整体情况（销量、同比、环比、目标达成率）。
2. 再聚焦关键分组（小学数学 / 语文 / 英语），分析目标达成、新签/老机构续用、体系/年级/区域下钻。
3. 识别关键问题：不要只看跌幅，要综合销量体量、目标缺口、机构数变化判断；体量极小的下滑可忽略。
4. 识别机会点：哪些体系/区域/年级在增长、新签贡献高、达成率领先。
5. 给出 2-3 条具体、可执行的行动建议（含优先级 P0/P1）。

## 重要原则
- 不死板套模板：根据数据实际情况决定重点。
- 目标对照：所有实际销量必须对照同学期的目标，计算达成率与缺口。
- 小学数学 = 小数同步 + 小数培优；语文 = 小学语文 + 初中语文；英语 = 小学英语 + 初中英语。
- 措辞中性：风险用“需重点关注并跟进续购转化”等温和表述，避免“打击/处置/诉讼”等词。
- 数据精度：销量保留一位小数（以“万”为单位），百分比一位小数，机构数为整数。
- 机构层面典型机构控制在 3 家以内；结尾以“续购进度需重点关注”收束。

## 报告数据（JSON）
```json
{json.dumps(report_data, ensure_ascii=False, indent=2, default=_default_serializer)}
```

## 快速事实索引
- 大盘目标达成率：{_rate(facts.get('overall_achievement'))}
- 达成率最低分组：{_ach(facts.get('worst_achievement_group'))}
- 达成率最高分组：{_ach(facts.get('best_achievement_group'))}
- 销量变化最大（下滑）业务口径：{_delta(facts.get('worst_sales_delta_biz'))}
- 销量变化最大（增长）业务口径：{_delta(facts.get('best_sales_delta_biz'))}
- 体系层面变化最大（下滑）：{_delta(facts.get('worst_system_delta'))}
- 体系层面变化最大（增长）：{_delta(facts.get('best_system_delta'))}
{("## 本次聚焦\n请重点分析 **" + focus + "** 的情况，但也不要忽略大盘上下文。") if focus else ""}

请输出符合以下结构的日报：
## 一、整体情况
## 二、关键问题
## 三、机会点
## 四、行动建议
"""
    return prompt


def _default_serializer(obj):
    if hasattr(obj, "item"):
        return obj.item()
    if hasattr(obj, "isoformat"):
        return obj.isoformat()
    return str(obj)


def _fmt_pct(v):
    return "—" if v is None else f"{v * 100:+.1f}%"


def render_text(report_data: dict, focus: Optional[str] = None) -> str:
    """纯文本兜底报表（只列事实）。"""
    meta = report_data.get("meta", {})
    overall = report_data.get("overall", {})
    groups = report_data.get("groups", {})
    dims = report_data.get("dimensions", {})
    lines = [f"教育出版销售日报（{meta.get('current_semester', '')}）",
             f"源文件：{meta.get('source_file', '')} | 生成：{meta.get('generated_at', '')[:19]}", ""]
    lines.append("【大盘】")
    lines.extend(_summary_lines(overall))
    lines.append("")
    if groups:
        lines.append("【学科分组】")
        for name, data in groups.items():
            if focus and name != focus:
                continue
            lines.append(f"\n{name}:")
            lines.extend(_summary_lines(data["summary"], indent="  "))
            for dk in ["by_system", "by_grade", "by_region"]:
                tbl = data.get(dk, [])
                if tbl:
                    lines.append(f"\n  {dk}（Top 8）：")
                    lines.append(_table_text(tbl[:8]))
        lines.append("")
    lines.append("【大盘维度下钻】")
    for dk, tbl in dims.items():
        if tbl:
            lines.append(f"\n{dk}（Top 10）：")
            lines.append(_table_text(tbl[:10]))
    return "\n".join(lines)


def _summary_lines(s: dict, indent: str = "") -> list:
    lines = []
    sales = s.get("sales", {})
    target = s.get("target", {})
    orgs = s.get("orgs", {})
    old = s.get("old", {})
    new = s.get("new_sign", {})
    if sales.get("current") is not None:
        lines.append(f"{indent}本期销量：{sales['current']:,} 本")
    if sales.get("yoy_change") is not None:
        lines.append(f"{indent}同比：{_fmt_pct(sales['yoy_change'])}")
    if sales.get("mom_change") is not None:
        lines.append(f"{indent}环比：{_fmt_pct(sales['mom_change'])}")
    if target:
        lines.append(f"{indent}目标：{target['value']:,} 本 | 达成率：{target['achievement_rate'] * 100:.1f}% | 缺口：{target['gap']:+,} 本")
    if orgs:
        lines.append(f"{indent}购买机构数：{orgs.get('current'):,} 家")
    if old:
        lines.append(f"{indent}老机构：{old['orgs']:,} 家，续用 {old['renew_orgs']:,} 家，流失 {old.get('churn_orgs', 0):,} 家，续用率 {_fmt_pct(old.get('renewal_rate'))}")
    if new:
        lines.append(f"{indent}新签机构：{new['orgs']:,} 家，新签销量 {new['sales']:,} 本")
    return lines


def _table_text(rows: list) -> str:
    if not rows:
        return "  无数据"
    headers = ["维度值", "本期", "上期", "环比变化", "同比", "新签", "续用", "流失"]
    keys = ["value", "sales_current", "sales_prior", "sales_delta", "yoy_change", "new_orgs", "renew_orgs", "churn_orgs"]

    def fmt(row, k):
        v = row.get(k)
        if v is None:
            return "—"
        if isinstance(v, str):
            return v
        if isinstance(v, float):
            return f"{v * 100:+.1f}%" if k in ("mom_change", "yoy_change") else f"{v:.1f}"
        return f"{v:,}"

    widths = [max(len(headers[i]), max(len(fmt(r, keys[i])) for r in rows)) for i in range(len(headers))]
    sep = " | "
    header = sep.join(headers[i].ljust(widths[i]) for i in range(len(headers)))
    rule = sep.join("-" * w for w in widths)
    body = "\n".join(sep.join(fmt(r, keys[i]).ljust(widths[i]) for i in range(len(headers))) for r in rows)
    return f"  {header}\n  {rule}\n  {body}"


def render_html(report_data: dict, title: str = "销售日报") -> str:
    """自包含 HTML 兜底报表。"""
    meta = report_data.get("meta", {})
    overall = report_data.get("overall", {})
    groups = report_data.get("groups", {})
    dims = report_data.get("dimensions", {})
    facts = report_data.get("facts", {})

    sections = [_section("大盘概览", _overall_html(overall))]
    for name, data in groups.items():
        parts = [_overall_html(data["summary"])]
        for dk in ["by_system", "by_grade", "by_version", "by_region"]:
            tbl = data.get(dk, [])
            if tbl:
                parts.append(f"<h4>{dk}</h4>" + _table_html(tbl))
        sections.append(_section(name, "\n".join(parts)))
    dim_parts = []
    for dk, tbl in dims.items():
        if tbl:
            dim_parts.append(f"<h4>{dk}</h4>" + _table_html(tbl))
    if dim_parts:
        sections.append(_section("大盘维度下钻", "\n".join(dim_parts)))
    facts_html = _facts_html(facts)
    if facts_html:
        sections.append(_section("关键事实索引（供智能体解读）", facts_html))
    data_json = html.escape(json.dumps(report_data, ensure_ascii=False, default=str))

    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title>
<style>
:root{{--bg:#f8fafc;--card:#fff;--ink:#0f172a;--muted:#64748b;--line:#e2e8f0;--blue:#2563eb;--red:#dc2626;--green:#16a34a}}
@media(prefers-color-scheme:dark){{:root{{--bg:#0f172a;--card:#1e293b;--ink:#f1f5f9;--muted:#94a3b8;--line:#334155}}}}
body{{margin:0;background:var(--bg);color:var(--ink);font:14px/1.6 system-ui,Microsoft YaHei,sans-serif}}
main{{max-width:1200px;margin:0 auto;padding:24px}}
h1{{font-size:26px;margin:0 0 6px}}h2{{font-size:20px;margin:0 0 12px}}h3{{font-size:16px;margin:18px 0 8px}}
.meta{{color:var(--muted);font-size:13px;margin-bottom:20px}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin-bottom:16px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px}}
.kpi{{background:#eff6ff;border-radius:8px;padding:12px}}.kpi .v{{font-size:22px;font-weight:700;color:var(--blue)}}
.kpi .l{{font-weight:600;margin-top:4px}}.kpi .d{{font-size:12px;color:var(--muted)}}
table{{border-collapse:collapse;width:100%;margin-top:8px}}
th,td{{border-bottom:1px solid var(--line);padding:8px;text-align:left;font-size:13px}}
th{{color:var(--muted);font-weight:600;cursor:pointer}}.up{{color:var(--green)}}.down{{color:var(--red)}}
pre{{background:#f1f5f9;padding:12px;border-radius:8px;overflow:auto;font-size:12px}}
</style></head><body><main>
<h1>{html.escape(title)}</h1>
<div class="meta">学期：{html.escape(meta.get('current_semester',''))} | 源文件：{html.escape(meta.get('source_file',''))} | 生成：{html.escape(str(meta.get('generated_at',''))[:19])}</div>
{''.join(sections)}
<div class="card"><h2>原始数据 JSON（便于复制给智能体）</h2><pre>{data_json}</pre></div>
</main><script>
window.addEventListener('load',()=>{{document.querySelectorAll('th').forEach(th=>{{th.addEventListener('click',()=>{{
const t=th.closest('table'),tb=t.querySelector('tbody'),i=Array.from(th.parentNode.children).indexOf(th);
const rows=Array.from(tb.querySelectorAll('tr')),asc=th.dataset.o!=='asc';th.dataset.o=asc?'asc':'desc';
rows.sort((a,b)=>{{const av=a.children[i].innerText.replace(/,/g,''),bv=b.children[i].innerText.replace(/,/g,''),an=parseFloat(av),bn=parseFloat(bv);
if(!isNaN(an)&&!isNaN(bn))return asc?an-bn:bn-an;return asc?av.localeCompare(bv):bv.localeCompare(av);}});
rows.forEach(r=>tb.appendChild(r));}});}});}});
</script></body></html>"""


def _section(title: str, body: str) -> str:
    return f'<div class="card"><h2>{html.escape(title)}</h2>{body}</div>' if isinstance(title, str) else body


def _overall_html(s: dict) -> str:
    sales = s.get("sales", {})
    target = s.get("target", {})
    orgs = s.get("orgs", {})
    old = s.get("old", {})
    new = s.get("new_sign", {})

    def pct(v):
        return "—" if v is None else f"{v * 100:+.1f}%"

    kpis = [
        ("本期销量", f"{sales.get('current'):,}" if sales.get('current') is not None else "—", f"同比 {pct(sales.get('yoy_change'))} · 环比 {pct(sales.get('mom_change'))}"),
        ("购买机构数", f"{orgs.get('current'):,}" if orgs.get('current') is not None else "—", f"上期 {orgs.get('prior'):,}"),
        ("新签机构", f"{new.get('orgs'):,}" if new else "—", f"新签销量 {new.get('sales'):,}" if new else ""),
        ("老机构续用", f"{old.get('renew_orgs'):,}" if old else "—", f"续用率 {pct(old.get('renewal_rate'))}" if old else ""),
    ]
    if target:
        kpis.insert(0, ("目标达成率", f"{target.get('achievement_rate') * 100:.1f}%", f"缺口 {target.get('gap'):+,} 本 · 目标 {target.get('value'):,} 本"))
    cards = "".join(f'<div class="kpi"><div class="v">{html.escape(str(v))}</div><div class="l">{html.escape(l)}</div><div class="d">{html.escape(d)}</div></div>' for l, v, d in kpis)
    return f'<div class="kpis">{cards}</div>'


def _table_html(rows: list) -> str:
    if not rows:
        return "<p>无数据</p>"
    keys = ["value", "sales_current", "sales_prior", "sales_delta", "mom_change", "yoy_change", "new_orgs", "renew_orgs", "churn_orgs", "renewal_rate"]
    headers = {"value": "维度值", "sales_current": "本期", "sales_prior": "上期", "sales_delta": "环比变化", "mom_change": "环比", "yoy_change": "同比", "new_orgs": "新签", "renew_orgs": "续用", "churn_orgs": "流失", "renewal_rate": "续用率"}
    present = [k for k in keys if any(k in r for r in rows)]

    def fmt(k, v):
        if v is None:
            return "—"
        if isinstance(v, str):
            return v
        if k in ("mom_change", "yoy_change", "renewal_rate"):
            cls = "up" if v > 0 else ("down" if v < 0 else "")
            return f'<span class="{cls}">{v * 100:.1f}%</span>'
        if k == "sales_delta":
            cls = "up" if v > 0 else ("down" if v < 0 else "")
            return f'<span class="{cls}">{v:+,}</span>'
        return f"{v:,}"

    thead = "<tr>" + "".join(f"<th>{html.escape(headers.get(k, k))}</th>" for k in present) + "</tr>"
    tbody = "".join("<tr>" + "".join(f"<td>{fmt(k, r.get(k))}</td>" for k in present) + "</tr>" for r in rows)
    return f"<table><thead>{thead}</thead><tbody>{tbody}</tbody></table>"


def _facts_html(facts: dict) -> str:
    lines = []
    if facts.get("overall_achievement") is not None:
        lines.append(f"大盘目标达成率：{facts['overall_achievement'] * 100:.1f}%")
    if facts.get("worst_achievement_group"):
        g = facts["worst_achievement_group"]
        lines.append(f"达成率最低分组：{g['name']}（{g['achievement_rate'] * 100:.1f}%，缺口 {g['gap']:,} 本）")
    if facts.get("best_achievement_group"):
        g = facts["best_achievement_group"]
        lines.append(f"达成率最高分组：{g['name']}（{g['achievement_rate'] * 100:.1f}%，缺口 {g['gap']:,} 本）")
    if facts.get("worst_sales_delta_biz"):
        d = facts["worst_sales_delta_biz"]
        lines.append(f"销量下滑最大业务口径：{d['value']}（{d['sales_delta']:,} 本）")
    if facts.get("best_sales_delta_biz"):
        d = facts["best_sales_delta_biz"]
        lines.append(f"销量增长最大业务口径：{d['value']}（+{d['sales_delta']:,} 本）")
    return "<ul>" + "".join(f"<li>{html.escape(l)}</li>" for l in lines) + "</ul>" if lines else ""


# =====================================================================
# CLI
# =====================================================================

def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="daily_report.py",
        description="教育出版学科销量日报工具（单文件版）。分析 Quick BI 明细表，生成结构化/智能体报告。",
    )
    parser.add_argument("--source", "-s", help="Quick BI 明细表 Excel 路径（本地文件模式）")
    parser.add_argument("--download", choices=["codex", "workbuddy"], default=None,
                        help="生成对应平台的浏览器下载操作清单（由智能体执行），不分析数据")
    parser.add_argument("--report-url", default="https://bi.aliyuncs.com/", help="Quick BI 报表页地址（配合 --download）")
    parser.add_argument("--save-dir", default="./downloads", help="下载 Excel 保存目录（配合 --download）")
    parser.add_argument("--current", "-c", default=DEFAULT_CURRENT, help=f"本期学期（默认 {DEFAULT_CURRENT}）")
    parser.add_argument("--prior", "-p", default=DEFAULT_PRIOR, help=f"上期学期（默认 {DEFAULT_PRIOR}）")
    parser.add_argument("--yoy", default=DEFAULT_YOY, help=f"同比学期（默认 {DEFAULT_YOY}）")
    parser.add_argument("--output-dir", "-o", default="./outputs", help="输出目录（默认 ./outputs）")
    parser.add_argument("--focus", "-f", default=None, help="聚焦分组，如 小学数学 / 语文 / 英语")
    parser.add_argument("--render", "-r", default="json", help="输出类型，可逗号分隔多个：text,html,agent_prompt,json（默认 json）")
    args = parser.parse_args(argv)

    # 仅输出下载清单
    if args.download:
        recipe = build_download_recipe(args.download, args.report_url, args.save_dir)
        print(f"【{args.download} 下载操作清单】请交给对应智能体执行：")
        print(json.dumps(recipe, ensure_ascii=False, indent=2))
        print(f"\n执行完成后，用以下命令分析：\n  python daily_report.py --source {args.save_dir}/平台数据明细_明细表_*.xlsx --current {args.current}")
        return 0

    if not args.source:
        parser.error("需提供 --source（Excel 路径）；或用 --download codex|workbuddy 生成下载清单")

    source = Path(args.source)
    if not source.exists():
        print(f"错误：源文件不存在 {source}", file=sys.stderr)
        return 1

    report = build_report_data(source, current=args.current, prior=args.prior, yoy=args.yoy)
    out = Path(args.output_dir)
    out.mkdir(parents=True, exist_ok=True)

    # 支持一次分析、多种格式输出（逗号分隔）
    renders = [r.strip() for r in args.render.split(",") if r.strip()]
    valid = {"text", "html", "agent_prompt", "json"}
    for r in renders:
        if r not in valid:
            parser.error(f"未知 --render 值：{r}（可选：text, html, agent_prompt, json）")

    for r in renders:
        if r == "json":
            p = out / f"report_{args.current}.json"
            p.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=_default_serializer), encoding="utf-8")
            print(f"已保存 JSON：{p}")
        elif r == "text":
            p = out / f"report_{args.current}.txt"
            p.write_text(render_text(report, focus=args.focus), encoding="utf-8")
            print(f"已保存文本报告：{p}")
        elif r == "html":
            p = out / f"report_{args.current}.html"
            p.write_text(render_html(report, title=f"销售日报 {args.current}"), encoding="utf-8")
            print(f"已保存 HTML 报告：{p}")
        elif r == "agent_prompt":
            p = out / f"agent_prompt_{args.current}.txt"
            p.write_text(build_agent_prompt(report, focus=args.focus), encoding="utf-8")
            print(f"已保存智能体 prompt：{p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
