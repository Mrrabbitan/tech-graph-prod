#!/usr/bin/env python3
"""Blueprint-style 五智能体上下游流程图：1 决策 + 1 规划 + 3 执行."""
from __future__ import annotations

import os

W, H = 1640, 1180

BG = "#0a1628"
GRID = "#112240"
PANEL = "#0d1f3c"
STROKE = "#00b4d8"
TEXT = "#caf0f8"
TEXT_MUTED = "#90e0ef"
LABEL = "#48cae4"

C_CTRL = "#a78bfa"
C_READ = "#00b4d8"
C_WRITE = "#06d6a0"
C_FB = "#f77f00"
C_DATA = "#ffffff"
C_DEC = "#fcd34d"


def esc(t: str) -> str:
    return (
        t.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def rect(x, y, w, h, fill=PANEL, stroke=STROKE, sw=1.2, rx=4):
    return (
        f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" ry="{rx}" '
        f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'
    )


def txt(x, y, s, size=11, fill=TEXT, weight="normal", anchor="start"):
    return (
        f'<text x="{x}" y="{y}" font-size="{size}" fill="{fill}" font-weight="{weight}" '
        f'text-anchor="{anchor}" font-family="Courier New,monospace">{esc(s)}</text>'
    )


def text_block(x, y, lines, size=10, fill=TEXT, leading=4):
    out = []
    dy = 0
    for line in lines:
        out.append(txt(x, y + dy, line, size=size, fill=fill))
        dy += size + leading
    return "\n".join(out)


def defs():
    return """
  <defs>
    <pattern id="grid" width="30" height="30" patternUnits="userSpaceOnUse">
      <path d="M 30 0 L 0 0 0 30" fill="none" stroke="#112240" stroke-width="0.5"/>
    </pattern>
    <marker id="m_ctl" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#a78bfa"/>
    </marker>
    <marker id="m_read" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#00b4d8"/>
    </marker>
    <marker id="m_write" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#06d6a0"/>
    </marker>
    <marker id="m_fb" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#f77f00"/>
    </marker>
    <marker id="m_data" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#ffffff"/>
    </marker>
    <marker id="m_dec" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#fcd34d"/>
    </marker>
  </defs>
"""


_MARK = {
    C_CTRL: "m_ctl",
    C_READ: "m_read",
    C_WRITE: "m_write",
    C_FB: "m_fb",
    C_DATA: "m_data",
    C_DEC: "m_dec",
}


def line_arrow(x1, y1, x2, y2, color, label="", dashed=False):
    dash = 'stroke-dasharray="6 4"' if dashed else ""
    mid = ""
    if label:
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        w = max(60, len(label) * 6 + 10)
        mid = (
            f'<rect x="{mx - w / 2}" y="{my - 14}" width="{w}" height="16" fill="{BG}" opacity="0.92"/>'
            f'<text x="{mx}" y="{my - 2}" font-size="9" fill="{TEXT_MUTED}" text-anchor="middle" '
            f'font-family="Courier New,monospace">{esc(label)}</text>'
        )
    return (
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="2" '
        f'{dash} marker-end="url(#{_MARK[color]})" fill="none"/>' + mid
    )


def poly(points, color, dashed=False, label="", label_at=None):
    dash = 'stroke-dasharray="6 4"' if dashed else ""
    pts = " ".join(f"{px},{py}" for px, py in points)
    out = (
        f'<polyline points="{pts}" fill="none" stroke="{color}" stroke-width="2" '
        f'{dash} marker-end="url(#{_MARK[color]})"/>'
    )
    if label and label_at:
        lx, ly = label_at
        w = max(80, len(label) * 6 + 10)
        out += (
            f'<rect x="{lx - w / 2}" y="{ly - 14}" width="{w}" height="16" fill="{BG}" opacity="0.92"/>'
            f'<text x="{lx}" y="{ly - 2}" font-size="9" fill="{TEXT_MUTED}" text-anchor="middle" '
            f'font-family="Courier New,monospace">{esc(label)}</text>'
        )
    return out


def card(x, y, w, h, tag, title, bullets, accent=C_CTRL):
    out = [
        rect(x, y, w, h, fill=PANEL, stroke=accent, sw=1.6),
        txt(x + 10, y + 18, tag, size=10, fill=LABEL, weight="700"),
        txt(x + 10, y + 38, title, size=13, fill=TEXT, weight="700"),
    ]
    out.append(text_block(x + 10, y + 58, bullets, size=10))
    return "\n".join(out)


def diamond(cx, cy, w, h, label):
    pts = f"{cx},{cy - h / 2} {cx + w / 2},{cy} {cx},{cy + h / 2} {cx - w / 2},{cy}"
    return (
        f'<polygon points="{pts}" fill="{PANEL}" stroke="{C_DEC}" stroke-width="1.5"/>'
        + txt(cx, cy + 4, label, size=11, fill=TEXT, weight="700", anchor="middle")
    )


def main():
    out_path = os.path.normpath(
        os.path.join(os.path.dirname(__file__), "..", "output", "quanke-five-agent-flow.svg")
    )
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    L: list[str] = []
    L.append(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}">')
    L.append("<style>text { font-family: 'Courier New', 'Lucida Console', monospace; }</style>")
    L.append(defs())
    L.append(f'<rect width="{W}" height="{H}" fill="{BG}"/>')
    L.append(f'<rect width="{W}" height="{H}" fill="url(#grid)" opacity="0.55"/>')

    L.append(txt(W / 2, 38, "全客存量 · 五智能体上下游流程（1 决策 + 1 规划 + 3 执行）", size=20, fill=TEXT, weight="700", anchor="middle"))
    L.append(
        txt(
            W / 2,
            58,
            "Decision → Planning → (Audience Build · Strategy Publish & Push · Outcome & Governance)",
            size=11,
            fill=TEXT_MUTED,
            anchor="middle",
        )
    )

    # Top row: input channels
    L.append(rect(28, 80, W - 56, 88, fill="#0b1c30"))
    L.append(txt(40, 102, "INPUT / 触点 & 上游", size=11, fill=LABEL, weight="700"))
    chips = [
        ("CB 智慧推荐弹窗", 40, 120),
        ("10010 来话 ASR", 240, 120),
        ("全客对话式策划 NLU", 420, 120),
        ("画像快照（标签集市+能开实时）", 640, 120),
        ("会话上下文 / 多轮槽位", 920, 120),
        ("运营人员人机协同（HITL）", 1160, 120),
    ]
    for label, x, y in chips:
        w = max(140, len(label) * 7 + 24)
        L.append(rect(x, y - 14, w, 26, fill="#0e2540", stroke=STROKE, sw=0.9, rx=12))
        L.append(txt(x + 8, y + 4, label, size=10, fill=TEXT))

    # Decision
    dx, dy, dw, dh = 60, 200, 560, 160
    L.append(
        card(
            dx,
            dy,
            dw,
            dh,
            "AGENT-1 / DECISION",
            "决策智能体",
            [
                "意图分类 INT-01~04 + 置信度",
                "场景判定 L1/L2（FTTR/单转融/宽带感知…）",
                "槽位抽取 StrategyConfig（时间·业务·触点·免打扰…）",
                "路由决策：planning | execution | clarify",
            ],
        )
    )

    # Planning
    px, py, pw, ph = 1020, 200, 560, 160
    L.append(
        card(
            px,
            py,
            pw,
            ph,
            "AGENT-2 / PLANNING",
            "规划智能体",
            [
                "RAG 召回：场景最优策略 Top-K + 原子配置切片",
                "Neo4j 推理：意图→场景→特征→产品→约束",
                "组装：客群 + 产品 + 触点 + 卖点/话术",
                "档位排序、约束校验、HITL 候选输出",
            ],
            accent=C_CTRL,
        )
    )

    # Decision arrow with diamond
    diamond_cx, diamond_cy = 820, dy + dh / 2
    L.append(line_arrow(dx + dw, diamond_cy, diamond_cx - 60, diamond_cy, C_CTRL, "结构化意图"))
    L.append(diamond(diamond_cx, diamond_cy, 120, 64, "route?"))
    L.append(line_arrow(diamond_cx + 60, diamond_cy, px - 6, diamond_cy, C_DEC, "planning"))
    # clarify back to input
    L.append(
        poly(
            [
                (diamond_cx, diamond_cy + 32),
                (diamond_cx, dy + dh + 30),
                (W / 2, dy + dh + 30),
                (W / 2, 168 + 4),
            ],
            C_FB,
            label="clarify · 反问/补全",
            label_at=(W / 2 - 130, dy + dh + 22),
        )
    )

    # Execution row: 3 agents
    ex_y = 410
    ex_w = 480
    ex_h = 200
    gap = 40
    total_w = ex_w * 3 + gap * 2
    ex0 = (W - total_w) / 2

    e_cards = [
        (
            "AGENT-3a / EXECUTION",
            "客群圈选执行",
            [
                "MCP: 客群圈选 · 客群生成 · 规模预估",
                "标签客群创建 · 切片预览",
                "写：全客客户群表 / 标签快照",
                "回：custgroup_id, scale, label_diff",
            ],
        ),
        (
            "AGENT-3b / EXECUTION",
            "策略写入与触点下发",
            [
                "MCP: 策略方案构建 · 策略配置写入",
                "触点下发：CB弹窗 · 10010 · 外呼 · 工单",
                "校验：免打扰/黑名单/产品互斥",
                "回：strategy_id, publish_status, push_records",
            ],
        ),
        (
            "AGENT-3c / EXECUTION",
            "效果回流与治理",
            [
                "采集 W/C/O 漏斗 + D_CREATE 时间字段",
                "10 档效能自动归类 · TOP-K 入选",
                "标签治理回流（频繁变更/未采纳/纠错）",
                "A/B 实验结果 → L3 情景记忆",
            ],
        ),
    ]
    e_x_list = []
    for i, (tag, title, bullets) in enumerate(e_cards):
        ex = ex0 + i * (ex_w + gap)
        e_x_list.append(ex)
        L.append(card(ex, ex_y, ex_w, ex_h, tag, title, bullets, accent=C_WRITE))

    # planning -> 3 executors fan-out
    fan_y = ex_y - 26
    L.append(line_arrow(px + pw / 2, py + ph, px + pw / 2, fan_y, C_CTRL))
    bus_left = e_x_list[0] + ex_w / 2
    bus_right = e_x_list[-1] + ex_w / 2
    L.append(f'<line x1="{bus_left}" y1="{fan_y}" x2="{bus_right}" y2="{fan_y}" stroke="{C_CTRL}" stroke-width="2"/>')
    L.append(
        f'<rect x="{(px + pw / 2) - 70}" y="{fan_y - 22}" width="140" height="16" fill="{BG}" opacity="0.92"/>'
        + txt(px + pw / 2, fan_y - 10, "可执行 StrategyPlan", size=9, fill=TEXT_MUTED, anchor="middle")
    )
    for i, ex in enumerate(e_x_list):
        cx = ex + ex_w / 2
        L.append(line_arrow(cx, fan_y, cx, ex_y - 2, C_CTRL, ""))

    # execution sequencing E1 -> E2 -> E3 (with labels)
    L.append(
        line_arrow(
            e_x_list[0] + ex_w,
            ex_y + ex_h / 2,
            e_x_list[1] - 4,
            ex_y + ex_h / 2,
            C_CTRL,
            "客群就绪",
        )
    )
    L.append(
        line_arrow(
            e_x_list[1] + ex_w,
            ex_y + ex_h / 2,
            e_x_list[2] - 4,
            ex_y + ex_h / 2,
            C_CTRL,
            "已发布·已推送",
        )
    )

    # MCP toolchain bar
    mb_y = ex_y + ex_h + 28
    mb_h = 76
    L.append(rect(28, mb_y, W - 56, mb_h, fill="#0b2238"))
    L.append(txt(40, mb_y + 22, "MCP TOOLCHAIN / 工具执行层", size=11, fill=LABEL, weight="700"))
    L.append(
        txt(
            40,
            mb_y + 44,
            "客群圈选 · 客群生成 · 产品推荐 · 触点推荐 · 话术生成 · 策略方案构建 · 策略配置写入 · 触点下发 · 效果回流",
            size=10,
            fill=TEXT_MUTED,
        )
    )
    L.append(
        txt(
            40,
            mb_y + 62,
            "适配触点：CB 弹窗 · 10010 来话 · 外呼 · 工单驱动 · 营业员工作台",
            size=10,
            fill=TEXT_MUTED,
        )
    )

    # 5 memory layer
    mem_y = mb_y + mb_h + 24
    L.append(txt(28, mem_y - 6, "LAYERED MEMORY / 五层记忆", size=11, fill=LABEL, weight="700"))
    cell_w = (W - 56 - 4 * 8) // 5
    tiers = [
        ("L1", "感知 Sensory", "事件流 · ASR · Token"),
        ("L2", "工作 Working", "上下文 · 槽位 · 草稿"),
        ("L3", "情景 Episodic", "W/C/O · A/B · 历史方案"),
        ("L4", "语义 Semantic", "标签 · KG · 意图 · 同义词"),
        ("L5", "程序 Procedural", "工具 Schema · Prompt · 规则"),
    ]
    for i, (idx, name, store) in enumerate(tiers):
        x0 = 28 + i * (cell_w + 8)
        L.append(rect(x0, mem_y, cell_w, 78, fill=PANEL, stroke=STROKE, sw=1))
        L.append(txt(x0 + 8, mem_y + 18, idx, size=10, fill=LABEL))
        L.append(txt(x0 + 8, mem_y + 34, name, size=11, fill=TEXT, weight="700"))
        L.append(txt(x0 + 8, mem_y + 52, store, size=9, fill=TEXT_MUTED))

    # Memory interactions
    # Decision reads L2/L4
    L.append(
        poly(
            [
                (dx + dw / 2, dy + dh),
                (dx + dw / 2, mb_y - 6),
            ],
            C_READ,
        )
    )
    # Planning reads L4 / L5 / L3
    L.append(
        poly(
            [
                (px + pw / 2 - 80, py + ph),
                (px + pw / 2 - 80, mem_y + 39),
                (28 + 3 * (cell_w + 8) + cell_w / 2, mem_y + 39),
            ],
            C_READ,
            label="RAG+KG",
            label_at=(px + pw / 2 - 80, mem_y - 12),
        )
    )

    # Execution writes L3 (E3) and L5 (governance writes)
    e3_cx = e_x_list[2] + ex_w / 2
    L.append(
        poly(
            [
                (e3_cx, ex_y + ex_h),
                (e3_cx, mem_y + 39),
                (28 + 2 * (cell_w + 8) + cell_w / 2, mem_y + 39),
            ],
            C_WRITE,
            dashed=True,
            label="效果写入",
            label_at=(e3_cx + 80, mem_y - 12),
        )
    )
    L.append(
        poly(
            [
                (e3_cx + 40, ex_y + ex_h + 20),
                (W - 80, ex_y + ex_h + 20),
                (W - 80, mem_y + 39),
                (28 + 4 * (cell_w + 8) + cell_w / 2, mem_y + 39),
            ],
            C_WRITE,
            dashed=True,
            label="规则/Prompt 治理",
            label_at=(W - 200, ex_y + ex_h + 12),
        )
    )

    # Data foundation
    fy = mem_y + 78 + 18
    fh = 112
    L.append(rect(28, fy, W - 56, fh, fill="#081426"))
    L.append(txt(40, fy + 22, "DATA & KNOWLEDGE BASE / 数据底座", size=11, fill=LABEL, weight="700"))
    L.append(
        text_block(
            40,
            fy + 38,
            [
                "全客：历史策略配置 · 策略效果(W/C/O/D_CREATE) · 原子配置项 Schema",
                "省分：产品特征 · 场景预设 · 热门话术 · 标签/档位/激励/约束 · 100+ 对话语料",
                "双引擎：向量库（RAG 切片） + Neo4j（意图-场景-产品-约束）",
            ],
            size=10,
        )
    )

    # Feedback loop: E3 -> Decision (re-plan)
    L.append(
        poly(
            [
                (e3_cx + 60, ex_y + 12),
                (W - 36, ex_y + 12),
                (W - 36, dy - 14),
                (dx + dw / 2, dy - 14),
                (dx + dw / 2, dy - 4),
            ],
            C_FB,
            label="重规划 · 下次会话个性化",
            label_at=(W - 240, dy - 22),
        )
    )

    # Legend at the very bottom
    ly0 = H - 36
    L.append(rect(28, ly0 - 12, W - 56, 36, fill="#081426", stroke=STROKE, sw=0.8))
    L.append(txt(40, ly0 + 6, "LEGEND", size=10, fill=TEXT, weight="700"))
    legend = [
        (C_CTRL, "控制流 orchestrate"),
        (C_DEC, "决策路由 route"),
        (C_READ, "读记忆/知识 retrieve"),
        (C_WRITE, "写记忆/治理 write"),
        (C_FB, "重规划 / 澄清 loop"),
        (C_DATA, "原始数据 ingest"),
    ]
    lx0 = 130
    step = (W - 56 - 130 - 20) // len(legend)
    for col, (c, lab) in enumerate(legend):
        xb = lx0 + col * step
        L.append(f'<rect x="{xb}" y="{ly0 - 4}" width="18" height="10" fill="{c}"/>')
        L.append(txt(xb + 24, ly0 + 5, lab, size=9, fill=TEXT_MUTED))

    L.append("</svg>")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print("Wrote", out_path)


if __name__ == "__main__":
    main()
