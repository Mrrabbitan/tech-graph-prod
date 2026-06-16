#!/usr/bin/env python3
"""Generate the 5 Scale AI guide diagrams as editable drawio XML files.

Usage:
    python3 scripts/generate-scale-ai-diagrams.py

Outputs (under output/):
    scale-ai-01-data-engine-loop.drawio
    scale-ai-02-gold-dataset-qc.drawio
    scale-ai-03-deployment-matrix.drawio
    scale-ai-04-industry-cases.drawio
    scale-ai-05-hub-spoke-repos.drawio

All diagrams share the project's default 5-color palette (blue / indigo /
orange / emerald / slate).  Pure-Python, stdlib only.
"""
from __future__ import annotations

import html
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"

# --- Palette (mirrors .agents/skills/tech-graph-prod/SKILL.md defaults) ----
PALETTE = {
    "L1": {"bg": "#eff6ff", "border": "#93c5fd", "accent": "#1d4ed8"},   # blue
    "L2": {"bg": "#eef2ff", "border": "#a5b4fc", "accent": "#4338ca"},   # indigo
    "L3": {"bg": "#fff7ed", "border": "#fdba74", "accent": "#c2410c"},   # orange (core)
    "L4": {"bg": "#ecfdf5", "border": "#6ee7b7", "accent": "#047857"},   # emerald
    "L5": {"bg": "#f8fafc", "border": "#94a3b8", "accent": "#334155"},   # slate
}
INK = "#0f172a"
MUTED = "#64748b"
SLATE_INK = "#1e293b"
SLATE_SOFT = "#475569"
FONT = "PingFang SC,Helvetica,sans-serif"
FONT_EN = "Helvetica,sans-serif"


# --- Mxfile builder --------------------------------------------------------
def _xa(s: str) -> str:
    """Escape an arbitrary string for use inside an XML attribute value.

    drawio stores HTML inside the `value=` attribute, so we must convert
    `<`, `>`, `&`, `"` into XML entities before embedding.  draw.io itself
    will unescape the value when rendering the cell (because the cell style
    has `html=1`).
    """
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


class Mxfile:
    def __init__(self, title: str, width: int = 1440, height: int = 1050):
        self.title = title
        self.width = width
        self.height = height
        self.cells: list[str] = []
        self._counter = 1000

    def nid(self) -> str:
        self._counter += 1
        return f"c{self._counter}"

    def rect(self, x: float, y: float, w: float, h: float, value: str,
             style: str, cid: Optional[str] = None) -> str:
        cid = cid or self.nid()
        self.cells.append(
            f'        <mxCell id="{cid}" parent="1" style="{style}" '
            f'value="{_xa(value)}" vertex="1">'
        )
        self.cells.append(
            f'          <mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" '
            f'as="geometry" />'
        )
        self.cells.append('        </mxCell>')
        return cid

    def text(self, x: float, y: float, w: float, h: float, value: str,
             *, size: int = 12, color: str = INK, bold: bool = False,
             italic: bool = False, align: str = "left",
             font: str = FONT) -> str:
        style_bits = [
            "text", "html=1", f"align={align}", "verticalAlign=middle",
            "strokeColor=none", "fillColor=none",
            f"fontSize={size}", f"fontColor={color}", f"fontFamily={font}",
        ]
        if bold:
            style_bits.append("fontStyle=1")
        elif italic:
            style_bits.append("fontStyle=2")
        return self.rect(x, y, w, h, value, ";".join(style_bits) + ";")

    def edge(self, src: Optional[str], tgt: Optional[str], style: str,
             value: str = "",
             waypoints: Optional[list[tuple[float, float]]] = None,
             source_xy: Optional[tuple[float, float]] = None,
             target_xy: Optional[tuple[float, float]] = None) -> str:
        cid = self.nid()
        attrs = [f'id="{cid}"', 'parent="1"', f'style="{style}"',
                 f'value="{_xa(value)}"', 'edge="1"']
        if src:
            attrs.append(f'source="{src}"')
        if tgt:
            attrs.append(f'target="{tgt}"')
        self.cells.append(f'        <mxCell {" ".join(attrs)}>')
        self.cells.append('          <mxGeometry relative="1" as="geometry">')
        if source_xy:
            self.cells.append(
                f'            <mxPoint x="{source_xy[0]}" '
                f'y="{source_xy[1]}" as="sourcePoint" />'
            )
        if target_xy:
            self.cells.append(
                f'            <mxPoint x="{target_xy[0]}" '
                f'y="{target_xy[1]}" as="targetPoint" />'
            )
        if waypoints:
            self.cells.append('            <Array as="points">')
            for wx, wy in waypoints:
                self.cells.append(f'              <mxPoint x="{wx}" y="{wy}" />')
            self.cells.append('            </Array>')
        self.cells.append('          </mxGeometry>')
        self.cells.append('        </mxCell>')
        return cid

    def render(self) -> str:
        parts = [
            '<mxfile host="Electron" agent="scripts/generate-scale-ai-diagrams.py" '
            'version="29.6.1">',
            f'  <diagram name="{html.escape(self.title)}" id="scale-ai-diagram">',
            f'    <mxGraphModel dx="1566" dy="1071" grid="1" gridSize="10" '
            f'guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" '
            f'pageScale="1" pageWidth="{self.width}" '
            f'pageHeight="{self.height}" math="0" shadow="0">',
            '      <root>',
            '        <mxCell id="0" />',
            '        <mxCell id="1" parent="0" />',
            *self.cells,
            '      </root>',
            '    </mxGraphModel>',
            '  </diagram>',
            '</mxfile>',
        ]
        return "\n".join(parts) + "\n"

    def write(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(self.render(), encoding="utf-8")


# --- Common style helpers -------------------------------------------------
def style_rounded(fill: str, stroke: str, *, stroke_width: float = 1.3,
                  arc: int = 12, opacity: int = 100,
                  dashed: bool = False, shadow: bool = False) -> str:
    bits = [
        "rounded=1", "whiteSpace=wrap", "html=1",
        f"fillColor={fill}", f"strokeColor={stroke}",
        f"strokeWidth={stroke_width}", f"arcSize={arc}",
        f"opacity={opacity}",
    ]
    if dashed:
        bits += ["dashed=1", "dashPattern=6 4"]
    if shadow:
        bits.append("shadow=1")
    return ";".join(bits) + ";"


def style_card(fill: str, stroke: str, *, accent: str = INK,
               stroke_width: float = 1.4, font_size: int = 12,
               arc: int = 12, valign: str = "middle",
               align: str = "center") -> str:
    return (
        f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill};"
        f"strokeColor={stroke};strokeWidth={stroke_width};arcSize={arc};"
        f"fontSize={font_size};fontColor={INK};fontFamily={FONT};"
        f"verticalAlign={valign};align={align};spacingTop=4;spacingLeft=8;"
        f"spacingRight=8;"
    )


def style_chip(fill: str, *, font_color: str = "#ffffff",
               font_size: int = 10.5) -> str:
    return (
        f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill};"
        f"strokeColor=none;fontColor={font_color};fontSize={font_size};"
        f"fontStyle=1;fontFamily={FONT};arcSize=50;verticalAlign=middle;"
        f"align=center;"
    )


def style_chip_outlined(stroke: str, *, font_color: Optional[str] = None,
                        font_size: int = 10.5) -> str:
    color = font_color or stroke
    return (
        f"rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;"
        f"strokeColor={stroke};strokeWidth=1.2;fontColor={color};"
        f"fontSize={font_size};fontStyle=1;fontFamily={FONT};arcSize=50;"
        f"verticalAlign=middle;align=center;"
    )


def style_edge(color: str, *, stroke_width: float = 1.6, dashed: bool = False,
               end_arrow: str = "classic", start_arrow: str = "none",
               curved: bool = False, rounded: bool = True,
               edge_style: str = "orthogonalEdgeStyle") -> str:
    bits = [
        f"edgeStyle={edge_style}",
        f"rounded={1 if rounded else 0}",
        "orthogonalLoop=1",
        "jettySize=auto",
        "html=1",
        f"strokeColor={color}",
        f"strokeWidth={stroke_width}",
        f"endArrow={end_arrow}", "endFill=1", "endSize=8",
        f"startArrow={start_arrow}",
        f"curved={1 if curved else 0}",
        f"fontColor={SLATE_SOFT}", f"fontFamily={FONT}", "fontSize=10.5",
    ]
    if dashed:
        bits += ["dashed=1", "dashPattern=6 4"]
    return ";".join(bits) + ";"


def title_block(mf: Mxfile, title: str, subtitle: str, y: int = 22,
                title_color: str = INK) -> None:
    mf.text(0, y, mf.width, 34, html.escape(title), size=22, bold=True,
            color=title_color, align="center")
    mf.text(0, y + 34, mf.width, 22,
            f'<span style="font-style:italic;">{html.escape(subtitle)}</span>',
            size=12, color=MUTED, align="center")


def footer_note(mf: Mxfile, note: str, y: Optional[int] = None) -> None:
    y = y if y is not None else mf.height - 32
    mf.text(0, y, mf.width, 22, html.escape(note), size=10.5, color=MUTED,
            italic=True, align="center")


# ========================================================================
# Diagram 1 — Data Engine 5-stage Loop + HITL
# ========================================================================
def build_diagram_1() -> Mxfile:
    mf = Mxfile("Scale AI · 数据引擎闭环 + HITL")
    title_block(
        mf,
        "Scale AI · 数据引擎五阶段闭环 + HITL 人机协作",
        "Data Engine Loop · Ingestion → Curation → Labeling → QC → Iteration · "
        "Drift Signal 触发再训练",
    )

    # --- HITL band (top overlay) ---
    hitl_x, hitl_y, hitl_w, hitl_h = 60, 110, 1320, 132
    mf.rect(
        hitl_x, hitl_y, hitl_w, hitl_h, "",
        style_rounded(PALETTE["L3"]["bg"], PALETTE["L3"]["border"],
                      stroke_width=1.4, arc=4, opacity=95, dashed=True),
    )
    mf.text(hitl_x + 14, hitl_y + 12, 280, 24,
            '<b>Human-in-the-Loop · 人机协作</b>',
            size=14, bold=True, color=PALETTE["L3"]["accent"])
    mf.text(hitl_x + 14, hitl_y + 38, 280, 20,
            "为概率系统提供伦理 / 安全缓冲",
            size=10.5, italic=True, color=SLATE_SOFT)

    hitl_cards = [
        ("专家校验", "Expert Verification", "复杂边界场景的人工标注与争议裁决", 600),
        ("金牌盲测", "Gold Set Blind Test", "新版本必须通过的准确率 / 偏见门禁", 870),
        ("RLHF 反馈", "Human Feedback RL", "用人类排名微调价值观与输出对齐", 1140),
    ]
    hitl_ids = []
    for zh, en, detail, x in hitl_cards:
        cid = mf.rect(
            x, hitl_y + 28, 240, 92, "",
            style_card(
                "#ffffff", PALETTE["L3"]["accent"],
                stroke_width=1.4, font_size=12, valign="top", align="left",
            ) + "spacingTop=8;",
        )
        mf.text(x + 12, hitl_y + 34, 220, 22,
                f'<b style="font-size:13px;color:{PALETTE["L3"]["accent"]};">'
                f'{zh}</b>'
                f' <span style="font-size:10px;color:{MUTED};">{en}</span>',
                size=12, color=PALETTE["L3"]["accent"])
        mf.text(x + 12, hitl_y + 56, 220, 56,
                f'<span style="font-size:11px;color:{SLATE_INK};'
                'line-height:1.55;">' + html.escape(detail) + '</span>',
                size=11, color=SLATE_INK, align="left")
        hitl_ids.append(cid)

    # --- 5 Stage cards ---
    stages = [
        ("01", "Ingestion · 摄取", "L1",
         "从文档 / 流数据 / 传感器抽取原始信号，建立来源可追溯与元数据绑定",
         ["DocStore", "Kafka", "Metadata"]),
        ("02", "Curation · 策展", "L2",
         "Feature Store 版本化管理；通过 Active Learning 选出高价值样本",
         ["Feature Store", "Active Learning", "DVC"]),
        ("03", "Labeling · 标注", "L3",
         "Model-Assisted 预标注 + 专家校验，规模化生产高质量标签",
         ["Model-Assisted", "Expert Review", "Schema"]),
        ("04", "QC · 质控", "L4",
         "模式一致性 + 偏见检测 + 金牌数据集盲测，作为流水线门禁",
         ["Gold Set", "Bias Check", "Schema-Lint"]),
        ("05", "Iteration · 迭代", "L5",
         "采集生产漂移信号触发闭环再训练，确保现实环境鲁棒性",
         ["Drift Detect", "Webhook", "Auto-Retrain"]),
    ]

    stage_y, stage_h = 290, 320
    card_w, gap = 240, 30
    x0 = (mf.width - (5 * card_w + 4 * gap)) // 2  # 60
    stage_ids = []
    stage_mid_xs = []
    for i, (idx, name, lvl, detail, chips) in enumerate(stages):
        cx = x0 + i * (card_w + gap)
        stage_mid_xs.append(cx + card_w // 2)
        pal = PALETTE[lvl]

        # outer band
        band_id = mf.rect(
            cx, stage_y, card_w, stage_h, "",
            style_rounded(pal["bg"], pal["border"], stroke_width=1.4,
                          arc=10, opacity=95),
        )
        stage_ids.append(band_id)

        # numbered strip
        mf.rect(
            cx + 12, stage_y + 12, 56, 44,
            f'<b style="font-size:22px;">{idx}</b>',
            f"rounded=1;whiteSpace=wrap;html=1;fillColor={pal['accent']};"
            f"strokeColor=none;arcSize=20;fontColor=#ffffff;"
            f"fontFamily={FONT};verticalAlign=middle;align=center;",
        )
        # stage name
        mf.text(cx + 78, stage_y + 12, card_w - 90, 24,
                f'<b>{html.escape(name.split(" · ")[0])}</b>',
                size=14, bold=True, color=pal["accent"])
        mf.text(cx + 78, stage_y + 36, card_w - 90, 20,
                f'<span style="font-style:italic;color:{MUTED};">'
                f'{html.escape(name.split(" · ")[1] if " · " in name else "")}'
                f'</span>',
                size=10.5, color=MUTED)

        # divider line via thin rect
        mf.rect(cx + 14, stage_y + 70, card_w - 28, 1, "",
                f"strokeColor={pal['border']};fillColor={pal['border']};"
                "html=1;")

        # description card (white)
        mf.rect(
            cx + 14, stage_y + 84, card_w - 28, 150,
            f'<span style="font-size:11.5px;color:{SLATE_INK};'
            f'line-height:1.65;">{html.escape(detail)}</span>',
            style_card(
                "#ffffff", pal["border"], stroke_width=1.0, font_size=11,
                valign="top", align="left",
            ) + "spacingTop=8;",
        )

        # chip section label
        mf.text(cx + 14, stage_y + 244, card_w - 28, 18,
                '<b>关键工艺</b>', size=10, bold=True, color=MUTED)

        # chips (1 per row x up to 3)
        chip_y = stage_y + 264
        for j, chip in enumerate(chips):
            chip_w = 6 * len(chip) + 36 if any(c.isascii() for c in chip) \
                else 12 * len(chip) + 22
            chip_x = cx + 14 + (j % 3) * 0  # stacked simpler: row layout
            # do single horizontal row, flex
        # Render chips horizontally with simple packing
        cur_x = cx + 14
        max_x = cx + card_w - 14
        for chip in chips:
            est_w = (len(chip) * 6.5 + 26) if all(ord(c) < 128 for c in chip) \
                else (len(chip) * 13 + 16)
            est_w = int(est_w)
            if cur_x + est_w > max_x:
                cur_x = cx + 14
                chip_y += 28
            mf.rect(cur_x, chip_y, est_w, 22, html.escape(chip),
                    style_chip(pal["accent"]))
            cur_x += est_w + 6

    # --- inter-stage primary arrows ---
    edge_style_primary = style_edge(
        "#2563eb", stroke_width=2.2,
        edge_style="straight",
    )
    for i in range(4):
        src = stage_ids[i]
        tgt = stage_ids[i + 1]
        mf.edge(src, tgt, edge_style_primary)

    # --- HITL → stage 3/4/5 (downward dashed) ---
    edge_hitl = style_edge(
        PALETTE["L3"]["accent"], stroke_width=1.6, dashed=True,
        edge_style="straight",
    )
    for k, target_id in enumerate(stage_ids[2:], start=0):
        src_id = hitl_ids[k]
        mf.edge(src_id, target_id, edge_hitl,
                value=f'<span style="background:#fff;padding:1px 4px;">'
                      f'{["专家校验","金牌盲测","RLHF 反馈"][k]}</span>')

    # --- loop-back arrow from Stage5 bottom back to Stage1 bottom ---
    loop_y_below = stage_y + stage_h + 50  # 660
    last_mid = stage_mid_xs[-1]
    first_mid = stage_mid_xs[0]
    mf.edge(
        None, None,
        style_edge("#7c3aed", stroke_width=2.0, dashed=True,
                   edge_style="orthogonalEdgeStyle"),
        value='<span style="background:#fff;padding:2px 6px;font-weight:bold;">'
              'Drift Signal · 触发再训练（闭环反馈）</span>',
        source_xy=(last_mid, stage_y + stage_h),
        target_xy=(first_mid, stage_y + stage_h),
        waypoints=[(last_mid, loop_y_below), (first_mid, loop_y_below)],
    )

    # --- key takeaway band ---
    takeaway_y = 760
    mf.rect(60, takeaway_y, 1320, 110, "",
            style_rounded("#f8fafc", "#94a3b8", stroke_width=1.2, arc=8,
                          opacity=95))
    mf.text(78, takeaway_y + 12, 400, 22,
            '<b>核心要点 · 数据资产观</b>',
            size=13, bold=True, color=INK)
    takeaways = [
        ("数据 = 瓶颈",
         "数据质量而非算法成为限制 Human-per-model ratio 的关键"),
        ("反馈即工程",
         "成功架构必须实现「数据 → 模型 → 反馈」工程化闭环"),
        ("自动化 ⨉ 专家",
         "Model-Assisted 标注必须与 HITL 校验深度耦合"),
    ]
    for i, (head, body) in enumerate(takeaways):
        tx = 78 + i * 435
        mf.text(tx, takeaway_y + 42, 30, 22, "●",
                size=18, color=PALETTE["L3"]["accent"], bold=True)
        mf.text(tx + 22, takeaway_y + 42, 410, 22,
                f'<b style="font-size:12px;color:{INK};">{head}</b>',
                size=12, bold=True, color=INK)
        mf.text(tx + 22, takeaway_y + 64, 410, 36,
                f'<span style="font-size:11px;color:{SLATE_SOFT};'
                f'line-height:1.5;">{html.escape(body)}</span>',
                size=11, color=SLATE_SOFT, align="left")

    # --- glossary chips footer ---
    gloss_y = 905
    mf.text(60, gloss_y, 200, 22,
            '<b>关键术语</b>', size=11, bold=True, color=MUTED)
    glossary = [
        ("Data Engine", "L1"),
        ("HITL", "L3"),
        ("Active Learning", "L2"),
        ("Drift Signal", "L4"),
        ("Model-Assisted Labeling", "L5"),
    ]
    gx = 168
    for name, lvl in glossary:
        w = int(len(name) * 7.6 + 32)
        mf.rect(gx, gloss_y - 2, w, 28, html.escape(name),
                style_chip_outlined(PALETTE[lvl]["accent"],
                                    font_color=PALETTE[lvl]["accent"],
                                    font_size=11))
        gx += w + 12

    footer_note(
        mf,
        "Source · Scale AI Data Engine 框架 · Truefoundry / Volvo / Red Hat MLOps 借鉴",
    )
    return mf


# ========================================================================
# Diagram 2 — Gold-Dataset QC Pipeline + Manual vs Automated comparison
# ========================================================================
def build_diagram_2() -> Mxfile:
    mf = Mxfile("Scale AI · 金牌数据集 QC 流水线 + 人机对比")
    title_block(
        mf,
        "数据质量管理 · 金牌数据集 QC 门禁流水线",
        "Gold Dataset 作为 CI/CD Gates · 控制模型晋升与灰度发布",
    )

    # --- Left half: gates pipeline ---
    pipe_x, pipe_y, pipe_w, pipe_h = 60, 120, 820, 780
    mf.rect(pipe_x, pipe_y, pipe_w, pipe_h, "",
            style_rounded("#fefce8", "#fde68a", stroke_width=1.3, arc=8,
                          opacity=92))
    mf.text(pipe_x + 20, pipe_y + 14, 380, 22,
            '<b>CI/CD Gates · 模型晋升流水线</b>',
            size=14, bold=True, color="#92400e")
    mf.text(pipe_x + 20, pipe_y + 38, 540, 18,
            f'<span style="color:{MUTED};font-style:italic;">'
            'New Model → Schema Gate → Bias Gate → Gold Set Blind Test → '
            'Latency Gate → Canary Release</span>',
            size=10.5, color=MUTED)

    # Source node
    src_x, src_y = pipe_x + 40, pipe_y + 88
    src_w, src_h = 200, 64
    src_id = mf.rect(src_x, src_y, src_w, src_h,
                     '<b style="font-size:14px;color:#fff;">New Model Version</b>'
                     '<br/><span style="font-size:10.5px;color:#fef3c7;">'
                     'v_next 提交进入晋升通道</span>',
                     style_card("#0f766e", "#0f766e", accent="#fff",
                                stroke_width=1.4) + "fontColor=#ffffff;")

    # Five gates
    gates = [
        ("01", "Schema Gate", "模式一致性",
         "训练/推理特征定义 100% 匹配", "L1"),
        ("02", "Bias Gate", "偏见检测",
         "受保护属性偏差 < 0.05", "L2"),
        ("03", "Gold Set Gate", "金牌盲测",
         "Top-K Accuracy ≥ 业务阈值", "L3"),
        ("04", "Latency Gate", "性能回归",
         "P99 延迟 ≤ SLO（如 200 ms）", "L4"),
        ("05", "Canary Gate", "灰度门禁",
         "影子流量 24 h 漂移指标 OK", "L5"),
    ]
    gate_ids = []
    gate_w, gate_h = 720, 80
    gate_x = pipe_x + 50
    gate_y0 = src_y + src_h + 32  # 240
    gate_gap = 26
    for i, (idx, en, zh, metric, lvl) in enumerate(gates):
        pal = PALETTE[lvl]
        gy = gate_y0 + i * (gate_h + gate_gap)

        gid = mf.rect(gate_x, gy, gate_w, gate_h, "",
                      style_rounded(pal["bg"], pal["border"],
                                    stroke_width=1.4, arc=10, opacity=95))
        # number badge
        mf.rect(gate_x + 14, gy + 14, 52, 52,
                f'<b style="font-size:20px;">{idx}</b>',
                f"rounded=1;whiteSpace=wrap;html=1;fillColor={pal['accent']};"
                f"strokeColor=none;arcSize=50;fontColor=#ffffff;"
                f"fontFamily={FONT};verticalAlign=middle;align=center;")
        # gate icon emoji-like text
        mf.text(gate_x + 80, gy + 12, 200, 22,
                f'<b>{en}</b> '
                f'<span style="font-weight:normal;color:{MUTED};">· {zh}</span>',
                size=13, bold=True, color=pal["accent"])
        mf.text(gate_x + 80, gy + 38, 480, 20,
                f'<span style="color:{SLATE_INK};font-size:11.5px;">'
                f'{html.escape(metric)}</span>',
                size=11.5, color=SLATE_INK)
        # pass/fail chips on right
        mf.rect(gate_x + gate_w - 168, gy + 18, 70, 22, "PASS",
                style_chip("#15803d"))
        mf.rect(gate_x + gate_w - 88, gy + 18, 70, 22, "FAIL → block",
                style_chip("#dc2626"))
        mf.text(gate_x + gate_w - 168, gy + 44, 158, 18,
                f'<span style="color:{MUTED};font-size:10.5px;">'
                '阻断或回滚到上一稳定版</span>',
                size=10, color=MUTED, align="center")
        gate_ids.append(gid)

    # connect source → gate1, gate_i → gate_{i+1}
    edge_pipe = style_edge("#92400e", stroke_width=2.0,
                           edge_style="orthogonalEdgeStyle")
    mf.edge(src_id, gate_ids[0], edge_pipe)
    for i in range(len(gate_ids) - 1):
        mf.edge(gate_ids[i], gate_ids[i + 1], edge_pipe)

    # production node
    prod_y = gate_y0 + len(gates) * (gate_h + gate_gap) + 4
    prod_id = mf.rect(gate_x, prod_y, gate_w, 56,
                      '<b style="font-size:14px;color:#fff;">'
                      'Production Promotion · 进入生产 + 持续监控</b>',
                      style_card("#1d4ed8", "#1d4ed8", accent="#fff",
                                 stroke_width=1.4) + "fontColor=#ffffff;")
    mf.edge(gate_ids[-1], prod_id, edge_pipe)

    # --- Right half: comparison table ---
    cmp_x, cmp_y, cmp_w = 920, 120, 460
    mf.text(cmp_x, cmp_y, cmp_w, 24,
            '<b>人工审核 vs 自动化 QC</b>',
            size=14, bold=True, color=INK)
    mf.text(cmp_x, cmp_y + 26, cmp_w, 18,
            f'<span style="color:{MUTED};font-style:italic;">'
            '4 维度对照 · 决定何时引入 HITL</span>',
            size=10.5, color=MUTED)

    # header row
    hdr_y = cmp_y + 56
    col_widths = [108, 168, 184]  # dimension, manual, auto
    col_x = [cmp_x, cmp_x + col_widths[0], cmp_x + col_widths[0] + col_widths[1]]
    headers = [("维度", "L5"), ("人工审核 · Manual", "L3"),
               ("自动化 QC · Automated", "L4")]
    row_h = 76
    for i, (label, lvl) in enumerate(headers):
        mf.rect(col_x[i], hdr_y, col_widths[i], 38,
                f'<b style="color:#fff;font-size:12.5px;">{label}</b>',
                style_card(PALETTE[lvl]["accent"], PALETTE[lvl]["accent"],
                           stroke_width=0, arc=4) + "fontColor=#ffffff;")

    rows = [
        ("成本",
         "高 · 随数据量线性增长",
         "低 · 初始投入后边际成本极低"),
        ("速度",
         "慢 · 受人力响应速率限制",
         "极快 · 毫秒级 / 实时"),
        ("一致性",
         "受主观偏见与疲劳影响",
         "极高 · 严格执行预设逻辑"),
        ("适用场景",
         "伦理判断 · 长尾案例定义 · RLHF",
         "模式一致性 · 异常检测 · 性能回归"),
    ]
    for r, (dim, manual, auto) in enumerate(rows):
        ry = hdr_y + 38 + r * row_h
        bg = "#ffffff" if r % 2 == 0 else "#f8fafc"
        # dimension cell
        mf.rect(col_x[0], ry, col_widths[0], row_h,
                f'<b style="font-size:12px;color:{INK};">{dim}</b>',
                style_card(bg, "#cbd5e1", stroke_width=0.8, arc=2))
        # manual cell
        mf.rect(col_x[1], ry, col_widths[1], row_h,
                f'<span style="font-size:11.5px;color:{SLATE_INK};'
                f'line-height:1.55;">{html.escape(manual)}</span>',
                style_card(bg, "#cbd5e1", stroke_width=0.8, arc=2,
                           font_size=11, align="left") +
                "spacingTop=8;spacingLeft=10;spacingRight=10;")
        # auto cell
        mf.rect(col_x[2], ry, col_widths[2], row_h,
                f'<span style="font-size:11.5px;color:{SLATE_INK};'
                f'line-height:1.55;">{html.escape(auto)}</span>',
                style_card(bg, "#cbd5e1", stroke_width=0.8, arc=2,
                           font_size=11, align="left") +
                "spacingTop=8;spacingLeft=10;spacingRight=10;")

    # --- 4 quality dimensions chip strip ---
    qd_y = 560
    mf.text(cmp_x, qd_y, cmp_w, 22,
            '<b>高质量数据 · 4 项检核指标</b>',
            size=13, bold=True, color=INK)
    quality = [
        ("一致性", "Consistency", "L1"),
        ("覆盖度", "Coverage", "L2"),
        ("无偏性", "Unbiasedness", "L3"),
        ("可审计性", "Auditability", "L4"),
    ]
    qx, qy = cmp_x, qd_y + 30
    for zh, en, lvl in quality:
        pal = PALETTE[lvl]
        mf.rect(qx, qy, cmp_w, 56,
                f'<table style="width:100%;border:0;color:{SLATE_INK};">'
                f'<tr><td style="font-size:13px;color:{pal["accent"]};'
                f'font-weight:bold;width:90px;">{zh}</td>'
                f'<td style="font-size:10.5px;color:{MUTED};font-style:italic;'
                f'padding-left:6px;">{en}</td></tr></table>',
                style_card(pal["bg"], pal["border"], stroke_width=1.1, arc=8,
                           font_size=11, align="left") +
                "spacingTop=6;spacingLeft=12;spacingRight=12;")
        qy += 60

    # --- Ground truth callout (full-width below pipeline + table) ---
    gt_y = 910
    mf.rect(60, gt_y, mf.width - 120, 82, "",
            style_rounded(PALETTE["L3"]["bg"],
                          PALETTE["L3"]["border"],
                          stroke_width=1.4, arc=10, opacity=95))
    mf.text(78, gt_y + 10, 300, 22,
            '<b>Ground Truth · 基准事实</b>',
            size=13, bold=True, color=PALETTE["L3"]["accent"])
    mf.text(78, gt_y + 36, mf.width - 156, 40,
            f'<span style="color:{SLATE_INK};font-size:11.5px;'
            f'line-height:1.55;">'
            '识别极端天气感知、罕见欺诈模式等 <b>长尾边缘案例</b>，'
            '注入金牌数据集后强制门禁；'
            '任何预测结果必须能回溯到对应的 <b>数据快照 + 代码版本</b>，'
            '满足合规审计与模型可解释性要求。'
            '</span>',
            size=11.5, color=SLATE_INK, align="left")

    footer_note(
        mf,
        "金牌数据集 = 多重专家验证 · 业务核心标准 · 长期不可篡改基线",
    )
    return mf


# ========================================================================
# Diagram 3 — Deployment Matrix 3×8 + risk callout
# ========================================================================
def build_diagram_3() -> Mxfile:
    mf = Mxfile("Scale AI · 三种部署模式 8 维度对比")
    title_block(
        mf,
        "企业级部署模式 · 8 维度对比矩阵",
        "Standard Cloud · Private VPC · Sovereign AI / On-Prem · 决定合规边界与成本结构",
    )

    # column / row layout
    table_x, table_y = 60, 130
    col_widths = [220, 360, 360, 360]
    col_x = [table_x]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w)
    table_w = sum(col_widths)

    # header
    hdr_h = 64
    headers = [
        ("对比维度", None),
        ("标准云<br/>Standard Cloud", "L1"),
        ("私有 VPC<br/>Private VPC", "L4"),
        ("主权 AI / On-Prem<br/>Sovereign AI", "L5"),
    ]
    for i, (label, lvl) in enumerate(headers):
        if lvl is None:
            fill, stroke, text_color = "#0f172a", "#0f172a", "#ffffff"
        else:
            fill = PALETTE[lvl]["accent"]
            stroke = PALETTE[lvl]["accent"]
            text_color = "#ffffff"
        mf.rect(col_x[i], table_y, col_widths[i], hdr_h,
                f'<b style="color:{text_color};font-size:14px;">{label}</b>',
                style_card(fill, stroke, accent=text_color,
                           stroke_width=0, arc=4) +
                f"fontColor={text_color};")

    rows = [
        ("数据驻留", "供应商控制", "企业选择特定区域",
         "企业物理机房 / 本地控制"),
        ("租户隔离", "Multi-tenant · 共享", "Dedicated · 独占租户",
         "完全物理隔离"),
        ("合规成本", "高 · 需严格三方审计",
         "中 · 复用企业云安全架构", "最低 · 全栈自主受控"),
        ("基础设施成本", "$15-40 / 人 / 月",
         "$25-60 / 人 / 月", "$80-150 / 人 / 月"),
        ("维护负担", "最低 · 厂商负责",
         "中低 · 云商负责硬件", "最高 · 企业团队全责"),
        ("设置时间", "数天", "2-4 周", "2-6 个月"),
        ("气隙隔离", "✗  不支持",
         "部分 · 逻辑隔离", "✓  支持 · 完全物理隔离"),
        ("典型场景", "非管制业务 · 敏捷实验",
         "金融 · 医疗 · 受监管企业", "国防 · 机密政府 · 主权安全"),
    ]
    row_h = 56
    cell_y = table_y + hdr_h
    for r, row in enumerate(rows):
        bg = "#ffffff" if r % 2 == 0 else "#f8fafc"
        for i, value in enumerate(row):
            stroke_color = "#e2e8f0"
            if i == 0:
                cell_value = (f'<b style="font-size:13px;color:{INK};">'
                              f'{value}</b>')
                content_style = (
                    style_card(bg, stroke_color, stroke_width=0.8, arc=0,
                               font_size=12, align="left") +
                    "spacingLeft=14;spacingRight=10;"
                )
            else:
                # detect special marks
                color = SLATE_INK
                weight = "normal"
                disp = html.escape(value)
                if value.startswith("✓"):
                    color = PALETTE["L4"]["accent"]
                    weight = "bold"
                elif value.startswith("✗"):
                    color = "#dc2626"
                    weight = "bold"
                elif value.startswith("部分"):
                    color = "#c2410c"
                    weight = "bold"
                cell_value = (
                    f'<span style="font-size:12px;color:{color};'
                    f'font-weight:{weight};line-height:1.5;">{disp}</span>'
                )
                content_style = (
                    style_card(bg, stroke_color, stroke_width=0.8, arc=0,
                               font_size=11.5, align="center") +
                    "spacingLeft=8;spacingRight=8;"
                )
            mf.rect(col_x[i], cell_y, col_widths[i], row_h,
                    cell_value, content_style)
        cell_y += row_h

    # --- Risk callout under matrix ---
    risk_y = cell_y + 28
    mf.rect(table_x, risk_y, table_w, 130, "",
            style_rounded("#fef2f2", "#fca5a5", stroke_width=1.4, arc=10,
                          opacity=95))
    mf.text(table_x + 22, risk_y + 14, 700, 24,
            '<b>⚠ 供应商稳定性风险 · 解耦数据与算力</b>',
            size=14, bold=True, color="#b91c1c")
    risk_bullets = [
        ("2025-06", "Meta 注资 143 亿美元，取得 Scale AI 49% 股权 · "
                    "可能改变供应商中立性"),
        ("2026-Q1", "原 CFO Dennis Cinelli 离职 · 财务治理稳定性出现波动"),
        ("策略",    "将「数据管理」与「底层算力」解耦 · "
                    "保留 Snorkel / Surge / Labelbox 等备份方案"),
    ]
    for i, (tag, body) in enumerate(risk_bullets):
        ry = risk_y + 46 + i * 24
        mf.rect(table_x + 22, ry, 70, 20, html.escape(tag),
                style_chip("#dc2626", font_size=10))
        mf.text(table_x + 100, ry, table_w - 120, 20,
                f'<span style="color:{SLATE_INK};font-size:11.5px;">'
                f'{html.escape(body)}</span>',
                size=11.5, color=SLATE_INK)

    # --- glossary right-side strip ---
    gloss_y = risk_y + 150
    mf.text(table_x, gloss_y, 200, 22,
            '<b>关键术语</b>', size=12, bold=True, color=MUTED)
    glossary = [
        ("Private VPC", "L1"),
        ("Sovereign AI", "L5"),
        ("Shadow AI", "L3"),
        ("Data Residency", "L4"),
        ("RBAC", "L2"),
    ]
    gx = 168
    for name, lvl in glossary:
        w = int(len(name) * 7.8 + 32)
        mf.rect(gx, gloss_y - 2, w, 28, html.escape(name),
                style_chip_outlined(PALETTE[lvl]["accent"],
                                    font_color=PALETTE[lvl]["accent"],
                                    font_size=11))
        gx += w + 12

    footer_note(
        mf,
        "Truefoundry Gateway Plane · Red Hat Secure Model Factory · "
        "Sigstore 签名 + Tekton Chains",
    )
    return mf


# ========================================================================
# Diagram 4 — Industry case study 2×2 grid
# ========================================================================
def build_diagram_4() -> Mxfile:
    mf = Mxfile("Scale AI · 四行业落地案例全景")
    title_block(
        mf,
        "行业落地案例 · 从实验到生产",
        "金融 · 保险 · 医疗 · 自动驾驶 · 每张卡片三段式（挑战 → 方案 → 成果）",
    )

    cases = [
        {
            "lvl": "L1",
            "industry": "金融 · 欺诈检测与审计",
            "en": "Finance · Fraud + Audit",
            "icon": "￥",
            "challenge": "交易欺诈模式演变极快，审计要求模型必须可解释，"
                         "每笔预测均能回溯依据",
            "solution": "实时 Feature Store 实现在线特征服务；"
                        "TrustyAI 持续在线漂移监测",
            "outcome": "亚秒级响应延迟；每笔预测可追溯到原始特征快照，"
                       "满足银保监审计要求",
            "chips": ["Feature Store", "TrustyAI", "Online Inference",
                      "XAI"],
        },
        {
            "lvl": "L2",
            "industry": "保险 · 理赔自动化",
            "en": "Insurance · Claims Automation",
            "icon": "▣",
            "challenge": "理赔涉及大量非结构化图像 / 手写文本；"
                         "多租户云存在数据泄露风险",
            "solution": "私有 VPC 部署 OCR + NLP 流水线；"
                        "争议案件强制 HITL 人工审核",
            "outcome": "理赔周期由 5 天压缩到 2 小时；数据始终留存企业边界内",
            "chips": ["Private VPC", "OCR", "Clinical NLP", "HITL Review"],
        },
        {
            "lvl": "L3",
            "industry": "医疗 · 临床记录处理",
            "en": "Healthcare · Clinical NLP",
            "icon": "✚",
            "challenge": "PHI 隐私极敏感，必须 100% 符合 HIPAA；"
                         "明文病历不可外泄",
            "solution": "「设计即隐私」网关架构；对病历执行 PII / PHI "
                        "自动掩码后再下发模型",
            "outcome": "在不触碰隐私明文的情况下完成临床报告结构化与质控",
            "chips": ["Gateway", "PHI Masking", "HIPAA", "Audit Log"],
        },
        {
            "lvl": "L4",
            "industry": "自动驾驶 · 感知模型迭代",
            "en": "Autonomous Driving · Perception",
            "icon": "▲",
            "challenge": "雨雪雾等长尾边缘案例标注成本极高，"
                         "迭代周期长，事故风险大",
            "solution": "Scale AI 标注引擎挖掘长尾样本；"
                        "Pipeline-centric 策略加速再训练",
            "outcome": "感知误差率大幅下降；从路测数据到模型上车的反馈循环显著缩短",
            "chips": ["Scale Engine", "Active Learning", "Pipeline-centric",
                      "Edge Cases"],
        },
    ]

    # 2x2 layout: 1440 wide, 1050 tall
    card_w, card_h = 640, 360
    margin_x = (mf.width - card_w * 2 - 40) // 2  # gap 40
    margin_y = 130
    grid_gap_y = 30

    for i, case in enumerate(cases):
        row, col = divmod(i, 2)
        cx = margin_x + col * (card_w + 40)
        cy = margin_y + row * (card_h + grid_gap_y)
        pal = PALETTE[case["lvl"]]

        # outer card
        mf.rect(cx, cy, card_w, card_h, "",
                style_rounded("#ffffff", pal["border"], stroke_width=1.5,
                              arc=14, opacity=100, shadow=True))

        # header strip
        mf.rect(cx, cy, card_w, 64,
                f'<table style="width:100%;color:#fff;">'
                f'<tr><td style="font-size:26px;font-weight:bold;width:60px;'
                f'text-align:center;">{html.escape(case["icon"])}</td>'
                f'<td><b style="font-size:16px;color:#fff;">'
                f'{html.escape(case["industry"])}</b><br/>'
                f'<span style="font-size:10.5px;color:#fef9c3;'
                f'font-style:italic;">{html.escape(case["en"])}</span></td>'
                f'</tr></table>',
                f"rounded=1;whiteSpace=wrap;html=1;fillColor={pal['accent']};"
                f"strokeColor=none;arcSize=14;fontColor=#ffffff;"
                f"fontFamily={FONT};verticalAlign=middle;align=left;"
                f"spacingLeft=14;spacingRight=14;")

        # 3 segments
        seg_titles = [("挑战", "Challenge", "#dc2626"),
                      ("方案", "Solution", pal["accent"]),
                      ("成果", "Outcome", PALETTE["L4"]["accent"])]
        seg_bodies = [case["challenge"], case["solution"], case["outcome"]]
        seg_y = cy + 80
        seg_h = 76
        for s in range(3):
            zh, en, color = seg_titles[s]
            sy = seg_y + s * (seg_h + 8)
            # left chip with seg title
            mf.rect(cx + 16, sy + 8, 84, 56,
                    f'<b style="color:#fff;font-size:13px;">{zh}</b><br/>'
                    f'<span style="color:#fff;font-size:9.5px;'
                    f'font-style:italic;">{en}</span>',
                    f"rounded=1;whiteSpace=wrap;html=1;fillColor={color};"
                    f"strokeColor=none;arcSize=12;fontColor=#ffffff;"
                    f"fontFamily={FONT};verticalAlign=middle;align=center;")
            # body box
            mf.rect(cx + 108, sy + 8, card_w - 124, 56,
                    f'<span style="font-size:12px;color:{SLATE_INK};'
                    f'line-height:1.6;">{html.escape(seg_bodies[s])}</span>',
                    style_card("#f8fafc", "#e2e8f0", stroke_width=0.8, arc=8,
                               font_size=11.5, align="left") +
                    "spacingTop=8;spacingLeft=12;spacingRight=12;")

        # chips row
        chip_y = cy + card_h - 40
        cur_x = cx + 16
        for chip in case["chips"]:
            est_w = int(len(chip) * 6.6 + 22) if all(ord(c) < 128 for c in chip) \
                else int(len(chip) * 13 + 16)
            mf.rect(cur_x, chip_y, est_w, 24, html.escape(chip),
                    style_chip(pal["accent"], font_size=10.5))
            cur_x += est_w + 8

    # --- glossary footer ---
    gl_y = margin_y + 2 * (card_h + grid_gap_y) - 4
    if gl_y + 60 > mf.height - 50:
        gl_y = mf.height - 80
    mf.text(60, gl_y, 200, 22,
            '<b>关键术语</b>', size=12, bold=True, color=MUTED)
    glossary = [
        ("Claims Automation", "L2"),
        ("Clinical NLP", "L3"),
        ("Fraud Pattern Recognition", "L1"),
        ("Active Learning", "L4"),
    ]
    gx = 168
    for name, lvl in glossary:
        w = int(len(name) * 7.6 + 32)
        mf.rect(gx, gl_y - 2, w, 28, html.escape(name),
                style_chip_outlined(PALETTE[lvl]["accent"],
                                    font_color=PALETTE[lvl]["accent"],
                                    font_size=11))
        gx += w + 12

    footer_note(
        mf,
        "案例参考 · Volvo 工业级蓝图 · Red Hat OpenShift AI · Truefoundry Gateway",
    )
    return mf


# ========================================================================
# Diagram 5 — Hub-and-Spoke + 5 Repository fan-out
# ========================================================================
def build_diagram_5() -> Mxfile:
    mf = Mxfile("Scale AI · Hub-and-Spoke + 5 仓库工业化结构")
    title_block(
        mf,
        "千级模型挑战 · Hub-and-Spoke 工业化结构 + 5 核心仓库",
        "中心 = 平台 / 基础设施 · 辐射 = 业务数据科学团队 · 5 仓库构建 Secure Model Factory",
    )

    cx_center, cy_center = mf.width // 2, 470

    # --- Central HUB ---
    hub_w, hub_h = 320, 200
    hub_x, hub_y = cx_center - hub_w // 2, cy_center - hub_h // 2
    pal_hub = PALETTE["L3"]
    hub_id = mf.rect(hub_x, hub_y, hub_w, hub_h, "",
                     style_rounded(pal_hub["bg"], pal_hub["accent"],
                                   stroke_width=2.4, arc=18, opacity=100,
                                   shadow=True))
    # title
    mf.text(hub_x, hub_y + 14, hub_w, 26,
            '<b>Platform Hub · 平台中心</b>',
            size=15, bold=True, color=pal_hub["accent"], align="center")
    mf.text(hub_x, hub_y + 40, hub_w, 18,
            f'<span style="color:{MUTED};font-style:italic;">'
            'Infrastructure & ML Platform Team</span>',
            size=10.5, color=MUTED, align="center")
    # MLOps toolchain chips
    tool_chips = ["Feature Store", "Model Registry", "Policy Engine"]
    tx, ty = hub_x + 14, hub_y + 68
    for chip in tool_chips:
        w = int(len(chip) * 7 + 20)
        mf.rect(tx, ty, w, 22, html.escape(chip),
                style_chip(pal_hub["accent"], font_size=10.5))
        tx += w + 8
        if tx + 120 > hub_x + hub_w:
            tx = hub_x + 14
            ty += 28
    # responsibilities text
    mf.text(hub_x + 14, hub_y + 124, hub_w - 28, 68,
            f'<span style="font-size:11px;color:{SLATE_INK};line-height:1.55;">'
            '统一 MLOps 工具链 · 安全策略 · 流水线模板<br/>'
            '负责吸收复杂性 · 解除人力耦合 · Policy-as-Code'
            '</span>',
            size=11, color=SLATE_INK, align="left")

    # --- 5 Core repositories arranged radially around HUB ---
    import math
    repos = [
        ("Model Config Repo", "JSON / YAML · 模型所有者 + SLO + 超参数", "L1"),
        ("Pipelines Repo", "Tekton / CI-CD 胶水代码 · 协调全流程", "L2"),
        ("Training Pipelines Repo",
         "Kubeflow Pipelines · 具体算法逻辑", "L3"),
        ("Data Management Repo",
         "DVC · 数据元数据 + 版本哈希 · 时间旅行", "L4"),
        ("Deployment Repo",
         "ArgoCD / GitOps · 生产环境状态镜像", "L5"),
    ]
    repo_w, repo_h = 280, 130
    # 5 angles, start from top, going clockwise, evenly distributed
    angles = [-90, -22, 38, 142, -142]  # degrees
    radius_x, radius_y = 460, 290
    repo_ids = []
    repo_anchors = []
    for i, (name, detail, lvl) in enumerate(repos):
        ang = math.radians(angles[i])
        rcx = cx_center + radius_x * math.cos(ang)
        rcy = cy_center + radius_y * math.sin(ang)
        rx, ry = rcx - repo_w / 2, rcy - repo_h / 2
        pal = PALETTE[lvl]
        rid = mf.rect(rx, ry, repo_w, repo_h, "",
                      style_rounded(pal["bg"], pal["accent"],
                                    stroke_width=1.6, arc=12, opacity=100,
                                    shadow=True))
        repo_ids.append(rid)
        repo_anchors.append((rx, ry, repo_w, repo_h, pal))
        # number badge
        mf.rect(rx + 12, ry + 12, 38, 26,
                f'<b style="font-size:13px;color:#fff;">0{i+1}</b>',
                f"rounded=1;whiteSpace=wrap;html=1;fillColor={pal['accent']};"
                f"strokeColor=none;arcSize=18;fontColor=#ffffff;"
                f"fontFamily={FONT};verticalAlign=middle;align=center;")
        mf.text(rx + 58, ry + 12, repo_w - 70, 28,
                f'<b style="font-size:13.5px;color:{pal["accent"]};">'
                f'{html.escape(name)}</b>',
                size=13, bold=True, color=pal["accent"], align="left")
        mf.text(rx + 14, ry + 50, repo_w - 28, 70,
                f'<span style="font-size:11px;color:{SLATE_INK};'
                f'line-height:1.55;">{html.escape(detail)}</span>',
                size=11, color=SLATE_INK, align="left")

    # connect hub ↔ repos with bidirectional lines
    edge_repo = style_edge(
        pal_hub["accent"], stroke_width=2.0, dashed=False,
        end_arrow="classic", start_arrow="classic",
        edge_style="none",
    )
    for rid in repo_ids:
        mf.edge(hub_id, rid, edge_repo)

    # --- 3 Spoke business teams (positioned in the gaps between repos) ---
    # Right slot: between Repo 02 (top-right) and Repo 03 (bottom-right)
    # Left  slot: between Repo 05 (top-left)  and Repo 04 (bottom-left)
    # Bottom slot: below the bottom-row repos (and above the checklist bar)
    spoke_w, spoke_h = 210, 78
    spoke_center_y = cy_center + 35  # 505 — middle of the right-side gap
    spokes = [
        ("业务 A · 营销画像", "Marketing Insights",
         mf.width - spoke_w - 18, spoke_center_y - spoke_h // 2),  # right
        ("业务 B · 风险控制", "Risk Control",
         18, spoke_center_y - spoke_h // 2),                       # left
        ("业务 C · 智能运营", "Smart Ops Loop",
         cx_center - spoke_w // 2, cy_center + 320),               # bottom
    ]
    spoke_ids = []
    for label, en, sx, sy in spokes:
        sid = mf.rect(
            sx, sy, spoke_w, spoke_h,
            f'<table style="width:100%;color:{INK};">'
            f'<tr><td style="font-size:13px;font-weight:bold;'
            f'color:{INK};">{html.escape(label)}</td></tr>'
            f'<tr><td style="font-size:10.5px;color:{MUTED};'
            f'font-style:italic;padding-top:4px;">'
            f'{html.escape(en)} · 业务 Spoke 团队</td></tr>'
            f'<tr><td style="font-size:10px;color:{SLATE_SOFT};'
            f'padding-top:6px;">通过装配线消费 Hub 平台能力</td></tr>'
            f'</table>',
            style_card("#ffffff", "#64748b", stroke_width=1.6, arc=10,
                       font_size=11, align="left") +
            "spacingTop=6;spacingLeft=10;spacingRight=10;",
        )
        spoke_ids.append(sid)

    # Spoke ↔ Hub edges (orthogonal routing with explicit waypoints to skip
    # the radial repo ring).
    spoke_edge_style = style_edge(
        "#475569", stroke_width=1.6, dashed=True,
        edge_style="orthogonalEdgeStyle",
        end_arrow="classic", start_arrow="classic",
    )
    # right spoke → hub (horizontal at cy_center)
    mf.edge(
        spoke_ids[0], hub_id, spoke_edge_style,
        value='<span style="background:#fff;padding:1px 4px;">装配线服务</span>',
    )
    # left spoke → hub (horizontal at cy_center)
    mf.edge(
        spoke_ids[1], hub_id, spoke_edge_style,
        value='<span style="background:#fff;padding:1px 4px;">装配线服务</span>',
    )
    # bottom spoke → hub (vertical at cx_center, through the gap between
    # bottom-left & bottom-right repos)
    mf.edge(
        spoke_ids[2], hub_id, spoke_edge_style,
        value='<span style="background:#fff;padding:1px 4px;">装配线服务</span>',
    )

    # --- Bottom strip · 5 checklist chips ---
    chk_y = mf.height - 110
    mf.rect(60, chk_y, mf.width - 120, 64, "",
            style_rounded("#f8fafc", "#94a3b8", stroke_width=1.2, arc=8,
                          opacity=95))
    mf.text(78, chk_y + 10, 260, 22,
            '<b>大规模实施 · 5 项检核</b>',
            size=13, bold=True, color=INK)
    checks = [
        ("权责明确", "L1"), ("可重复流水线", "L2"),
        ("Policy-as-Code", "L3"), ("资产化数据", "L4"),
        ("L3 闭环监控", "L5"),
    ]
    chx = 220
    for name, lvl in checks:
        w = int(len(name) * 12 + 28)
        mf.rect(chx, chk_y + 18, w, 28,
                f'✓ {html.escape(name)}',
                style_chip(PALETTE[lvl]["accent"], font_size=11.5))
        chx += w + 14

    # --- glossary footer ---
    gl_y = mf.height - 38
    mf.text(0, gl_y, mf.width, 22,
            f'<span style="color:{MUTED};font-style:italic;">'
            'Hub-and-Spoke · MVPS · MLOps Fleet Management · Feature Store · '
            'Data Contract</span>',
            size=10.5, color=MUTED, italic=True, align="center")
    return mf


# ========================================================================
# Main entry
# ========================================================================
def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    diagrams = [
        ("scale-ai-01-data-engine-loop.drawio", build_diagram_1),
        ("scale-ai-02-gold-dataset-qc.drawio", build_diagram_2),
        ("scale-ai-03-deployment-matrix.drawio", build_diagram_3),
        ("scale-ai-04-industry-cases.drawio", build_diagram_4),
        ("scale-ai-05-hub-spoke-repos.drawio", build_diagram_5),
    ]
    for name, builder in diagrams:
        target = OUTPUT_DIR / name
        mf = builder()
        mf.write(target)
        print(f"✓ wrote {target} ({len(mf.cells)} mxCell lines)")


if __name__ == "__main__":
    main()
