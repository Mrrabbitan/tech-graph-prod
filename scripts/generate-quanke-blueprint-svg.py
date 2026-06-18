#!/usr/bin/env python3
"""Emit Blueprint-style SVG: 决策/规划/执行 三智能体 + 五层记忆 + 语义箭头图例."""
from __future__ import annotations

import os

W, H = 1440, 1050

BG = "#0a1628"
GRID = "#112240"
PANEL = "#0d1f3c"
STROKE = "#00b4d8"
TEXT = "#caf0f8"
TEXT_MUTED = "#90e0ef"
LABEL = "#48cae4"

COLOR_CONTROL = "#a78bfa"
COLOR_READ = "#00b4d8"
COLOR_WRITE = "#06d6a0"
COLOR_FEEDBACK = "#f77f00"
COLOR_DATA = "#ffffff"


def esc(t: str) -> str:
    return (
        t.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def rect(x, y, rw, rh, fill=PANEL, stroke=STROKE, sw=1.2, rx=4):
    return (
        f'<rect x="{x}" y="{y}" width="{rw}" height="{rh}" rx="{rx}" ry="{rx}" '
        f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'
    )


def text_block(x, y, lines, size=11, fill=TEXT, weight="normal"):
    out = []
    dy = 0
    for line in lines:
        out.append(
            f'<text x="{x}" y="{y + dy}" font-size="{size}" fill="{fill}" font-weight="{weight}" '
            f'font-family="Courier New,monospace">{esc(line)}</text>'
        )
        dy += size + 4
    return "\n".join(out)


def arrow_defs():
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
  </defs>
"""


def line_arrow(x1, y1, x2, y2, color, mid_label="", dashed=False):
    dash = 'stroke-dasharray="6 4"' if dashed else ""
    mid_id = {
        COLOR_CONTROL: "m_ctl",
        COLOR_READ: "m_read",
        COLOR_WRITE: "m_write",
        COLOR_FEEDBACK: "m_fb",
        COLOR_DATA: "m_data",
    }[color]
    mid = ""
    if mid_label:
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        mid = (
            f'<rect x="{mx - 42}" y="{my - 14}" width="84" height="16" fill="{BG}" opacity="0.92"/>'
            f'<text x="{mx}" y="{my - 2}" font-size="9" fill="{TEXT_MUTED}" text-anchor="middle" '
            f'font-family="Courier New,monospace">{esc(mid_label)}</text>'
        )
    return (
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="2" '
        f'{dash} marker-end="url(#{mid_id})" fill="none"/>' + mid
    )


def poly_arrow(points, color, mid_id_key):
    pts = " ".join(f"{px},{py}" for px, py in points)
    return (
        f'<polyline points="{pts}" fill="none" stroke="{color}" stroke-width="2" '
        f'marker-end="url(#{mid_id_key})"/>'
    )


def agent_card(x, y, w, h, tag, title, bullets):
    parts = [
        rect(x, y, w, h, fill=PANEL, stroke=COLOR_CONTROL, sw=1.6),
        f'<text x="{x + 10}" y="{y + 18}" font-size="10" fill="{LABEL}" font-weight="700" '
        f'font-family="Courier New,monospace">{esc(tag)}</text>',
        f'<text x="{x + 10}" y="{y + 36}" font-size="13" fill="{TEXT}" font-weight="700" '
        f'font-family="Courier New,monospace">{esc(title)}</text>',
        text_block(x + 10, y + 52, bullets, size=10),
    ]
    return "\n".join(parts)


def memory_row(x, y, w, h, idx, name, store):
    return "\n".join(
        [
            rect(x, y, w, h, fill="#0d1f3c", stroke=STROKE, sw=1),
            f'<text x="{x + 8}" y="{y + 18}" font-size="10" fill="{LABEL}" font-family="Courier New,monospace">{esc(idx)}</text>',
            f'<text x="{x + 8}" y="{y + 34}" font-size="11" fill="{TEXT}" font-weight="700" '
            f'font-family="Courier New,monospace">{esc(name)}</text>',
            f'<text x="{x + 8}" y="{y + 52}" font-size="9" fill="{TEXT_MUTED}" font-family="Courier New,monospace">{esc(store)}</text>',
        ]
    )


def main():
    out_path = os.path.join(os.path.dirname(__file__), "..", "output", "quanke-three-agent-memory.svg")
    out_path = os.path.normpath(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    lines: list[str] = []
    lines.append(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}">')
    lines.append("<style>text { font-family: 'Courier New', 'Lucida Console', monospace; }</style>")
    lines.append(arrow_defs())
    lines.append(f'<rect width="{W}" height="{H}" fill="{BG}"/>')
    lines.append(f'<rect width="{W}" height="{H}" fill="url(#grid)" opacity="0.55"/>')

    lines.append(
        f'<text x="{W / 2}" y="38" font-size="20" fill="{TEXT}" text-anchor="middle" font-weight="700" '
        f'font-family="Courier New,monospace">全客存量 · 三智能体 × 五层记忆 · 总架构</text>'
    )
    lines.append(
        f'<text x="{W / 2}" y="58" font-size="11" fill="{TEXT_MUTED}" text-anchor="middle" '
        f'font-family="Courier New,monospace">'
        f'Decision / Planning / Execution  ·  Sensory→Working→Episodic→Semantic→Procedural</text>'
    )

    lx, ly = 28, 88
    lw, lh = 220, 200
    lines.append(rect(lx, ly, lw, lh))
    lines.append(
        f'<text x="{lx + 10}" y="{ly + 22}" font-size="11" fill="{LABEL}" font-weight="700" '
        f'font-family="Courier New,monospace">INPUT / 触点</text>'
    )
    lines.append(
        text_block(
            lx + 10,
            ly + 40,
            [
                "· CB智慧推荐弹窗事件",
                "· 10010来话 / 新客服ASR",
                "· 全客对话式策划(NLU)",
                "· 标签集市+能开实时画像",
                "· 会话上下文 session",
            ],
            size=10,
        )
    )

    ax = 268
    ay = 88
    aw = 292
    ah = 168
    gap = 18
    lines.append(
        agent_card(
            ax,
            ay,
            aw,
            ah,
            "DECISION",
            "决策智能体",
            [
                "意图识别 INT-01~04",
                "场景判定 L1/L2",
                "槽位抽取 StrategyConfig",
                "路由 planning|exec|clarify",
            ],
        )
    )
    lines.append(
        agent_card(
            ax + aw + gap,
            ay,
            aw,
            ah,
            "PLANNING",
            "规划智能体",
            [
                "RAG 场景最优策略Top-K",
                "Neo4j 场景→特征→产品",
                "客群/产品/触点/话术组装",
                "档位排序·约束校验",
            ],
        )
    )
    lines.append(
        agent_card(
            ax + 2 * (aw + gap),
            ay,
            aw,
            ah,
            "EXECUTION",
            "执行智能体",
            [
                "MCP工具链调用",
                "客群创建·策略写入全客",
                "触点下发·工单/弹窗",
                "效果回流 W/C/O 漏斗",
            ],
        )
    )

    mx, my = ax, ay + ah + 28
    mw = 1192
    mh = 72
    lines.append(rect(mx, my, mw, mh, fill="#0b2238"))
    lines.append(
        f'<text x="{mx + 12}" y="{my + 22}" font-size="11" fill="{LABEL}" font-weight="700" '
        f'font-family="Courier New,monospace">MCP TOOLCHAIN / 工具执行层</text>'
    )
    lines.append(
        f'<text x="{mx + 12}" y="{my + 44}" font-size="10" fill="{TEXT_MUTED}" font-family="Courier New,monospace">'
        f'客群圈选 · 客群生成 · 产品推荐 · 触点推荐 · 话术生成 · 策略方案构建 · 策略配置写入</text>'
    )

    mem_y = my + mh + 24
    lines.append(
        f'<text x="28" y="{mem_y - 6}" font-size="11" fill="{LABEL}" font-weight="700" '
        f'font-family="Courier New,monospace">LAYERED MEMORY / 五层记忆</text>'
    )
    cell_w = (W - 56 - 4 * 8) // 5
    tiers = [
        ("L1", "感知 Sensory", "事件流·ASR·原始Token"),
        ("L2", "工作 Working", "多轮上下文·槽位草稿"),
        ("L3", "情景 Episodic", "触达/接触/订购·A/B"),
        ("L4", "语义 Semantic", "标签库·KG·意图库·同义词"),
        ("L5", "程序 Procedural", "工具Schema·Prompt·档位规则"),
    ]
    for i, (idx, name, store) in enumerate(tiers):
        x0 = 28 + i * (cell_w + 8)
        lines.append(memory_row(x0, mem_y, cell_w, 86, idx, name, store))

    fx, fy = 28, mem_y + 86 + 20
    fw, fh = W - 56, 112
    lines.append(rect(fx, fy, fw, fh, fill="#081426"))
    lines.append(
        f'<text x="{fx + 12}" y="{fy + 22}" font-size="11" fill="{LABEL}" font-weight="700" '
        f'font-family="Courier New,monospace">{esc("DATA & KNOWLEDGE BASE / 数据底座")}</text>'
    )
    lines.append(
        text_block(
            fx + 12,
            fy + 38,
            [
                "全客：历史策略配置·策略效果(W/C/O/D_CREATE)·原子配置项Schema",
                "省分：产品特征·场景预设产品/触点·热门话术·标签/档位/激励/约束·对话语料100+",
                "双引擎：向量库(RAG切片) + Neo4j(意图-场景-产品-约束)",
            ],
            size=10,
        )
    )

    rx, ry = W - 248, 88
    rw, rh = 220, 200
    lines.append(rect(rx, ry, rw, rh))
    lines.append(
        f'<text x="{rx + 10}" y="{ry + 22}" font-size="11" fill="{LABEL}" font-weight="700" '
        f'font-family="Courier New,monospace">OUTPUT / 交付</text>'
    )
    lines.append(
        text_block(
            rx + 10,
            ry + 40,
            [
                "结构化 StrategyPlan",
                "策略画布可发布JSON",
                "弹窗卖点/话术/跳转",
                "触点后评价回流",
            ],
            size=10,
        )
    )

    # Arrows
    lines.append(line_arrow(lx + lw, ay + ah - 18, ax - 6, ay + ah - 18, COLOR_DATA, "事件/文本"))
    lines.append(line_arrow(ax + aw, ay + ah - 18, ax + aw + gap - 2, ay + ah - 18, COLOR_CONTROL, "结构化意图"))
    lines.append(
        line_arrow(ax + 2 * aw + gap, ay + ah - 18, ax + 2 * aw + 2 * gap - 2, ay + ah - 18, COLOR_CONTROL, "方案草案")
    )

    cx1 = ax + aw / 2
    cx2 = ax + aw + gap + aw / 2
    cx3 = ax + 2 * (aw + gap) + aw / 2
    my_mid = mem_y + 43
    for cx in (cx1, cx2, cx3):
        lines.append(line_arrow(cx, ay + ah, cx, my_mid, COLOR_READ, "检索"))

    ex = ax + 2 * (aw + gap) + aw / 2
    lines.append(line_arrow(ex, ay + ah + 4, min(ex + 140, rx - 20), my_mid, COLOR_WRITE, "效果写入", dashed=True))

    fb_y = fy + fh + 14
    lines.append(
        poly_arrow(
            [
                (rx + rw / 2, ry + rh),
                (rx + rw / 2, fb_y),
                (lx + lw / 2, fb_y),
                (lx + lw / 2, ly + lh - 8),
            ],
            COLOR_FEEDBACK,
            "m_fb",
        )
    )
    lines.append(
        f'<text x="{(rx + rw / 2 + lx + lw / 2) / 2}" y="{fb_y - 6}" font-size="10" fill="{COLOR_FEEDBACK}" '
        f'text-anchor="middle" font-family="Courier New,monospace">重规划·澄清·下次会话个性化</text>'
    )

    lines.append(
        line_arrow(fx + fw / 2, fy, ax + aw + gap + aw / 2, my + mh + 2, COLOR_READ, "RAG+KG")
    )

    # Legend
    lyg = H - 30
    lines.append(rect(28, lyg - 10, W - 56, 42, fill="#081426", stroke=STROKE, sw=0.8))
    lines.append(
        f'<text x="40" y="{lyg + 4}" font-size="10" fill="{TEXT}" font-weight="700" '
        f'font-family="Courier New,monospace">LEGEND</text>'
    )
    legend = [
        (COLOR_CONTROL, "控制流 orchestrate"),
        (COLOR_READ, "读记忆/知识 retrieve"),
        (COLOR_WRITE, "写情景记忆 feedback"),
        (COLOR_FEEDBACK, "重规划 loop"),
        (COLOR_DATA, "原始数据 ingest"),
    ]
    lx0 = 120
    for col, (c, lab) in enumerate(legend):
        xb = lx0 + col * 248
        lines.append(f'<rect x="{xb}" y="{lyg - 2}" width="18" height="10" fill="{c}"/>')
        lines.append(
            f'<text x="{xb + 24}" y="{lyg + 7}" font-size="9" fill="{TEXT_MUTED}" font-family="Courier New,monospace">'
            f'{esc(lab)}</text>'
        )

    lines.append("</svg>")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("Wrote", out_path)


if __name__ == "__main__":
    main()
