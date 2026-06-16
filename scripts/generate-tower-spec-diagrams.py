#!/usr/bin/env python3
"""Generate 3 drawio diagrams for the China Tower work-order QC spec.

Usage:
    python3 scripts/generate-tower-spec-diagrams.py

Outputs (under output/):
    tower-spec-01-system-architecture.drawio
    tower-spec-02-functional-flow.drawio
    tower-spec-03-network-topology.drawio

Visual style: README Style 1 — Flat Icon (white bg, colored accents).
Pure-Python, stdlib only.
"""
from __future__ import annotations

import html
import math
from pathlib import Path
from typing import Optional

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"

# --- Style 1 Flat Icon palette (references/style-1-flat-icon.md) -----------
S1 = {
    "bg":       "#ffffff",
    "box_fill": "#ffffff",
    "box_stroke": "#d1d5db",
    "text_primary": "#111827",
    "text_secondary": "#6b7280",
    "flow_a": "#2563eb",   # main / blue
    "flow_b": "#ea580c",   # alt / orange (used for auth / core)
    "flow_c": "#16a34a",   # data / green
    "flow_d": "#9333ea",   # async / purple
    "red":    "#dc2626",
}
TINTS = {
    "blue":   {"bg": "#eff6ff", "border": "#93c5fd", "accent": "#1d4ed8"},
    "indigo": {"bg": "#eef2ff", "border": "#a5b4fc", "accent": "#4338ca"},
    "orange": {"bg": "#fff7ed", "border": "#fdba74", "accent": "#c2410c"},
    "green":  {"bg": "#f0fdf4", "border": "#6ee7b7", "accent": "#047857"},
    "slate":  {"bg": "#f8fafc", "border": "#94a3b8", "accent": "#334155"},
    "red":    {"bg": "#fef2f2", "border": "#fca5a5", "accent": "#b91c1c"},
}
FONT = "Helvetica Neue,Helvetica,PingFang SC,Microsoft YaHei,sans-serif"
INK = S1["text_primary"]
MUTED = S1["text_secondary"]


# --- XML attribute escaper (drawio stores HTML in value="…") ---------------
def _xa(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


# --- Mxfile builder (copied from generate-scale-ai-diagrams.py) -----------
class Mxfile:
    def __init__(self, title: str, w: int = 1440, h: int = 1050):
        self.title, self.width, self.height = title, w, h
        self.cells: list[str] = []
        self._c = 1000

    def nid(self) -> str:
        self._c += 1; return f"c{self._c}"

    def rect(self, x, y, w, h, value, style, cid=None):
        cid = cid or self.nid()
        self.cells += [
            f'        <mxCell id="{cid}" parent="1" style="{style}" value="{_xa(value)}" vertex="1">',
            f'          <mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry" />',
            '        </mxCell>']
        return cid

    def text(self, x, y, w, h, value, *, size=12, color=INK, bold=False,
             italic=False, align="left"):
        bits = ["text","html=1",f"align={align}","verticalAlign=middle",
                "strokeColor=none","fillColor=none",
                f"fontSize={size}",f"fontColor={color}",f"fontFamily={FONT}"]
        if bold: bits.append("fontStyle=1")
        elif italic: bits.append("fontStyle=2")
        return self.rect(x,y,w,h,value,";".join(bits)+";")

    def edge(self, src, tgt, style, value="", waypoints=None,
             source_xy=None, target_xy=None):
        cid = self.nid()
        attrs = [f'id="{cid}"','parent="1"',f'style="{style}"',
                 f'value="{_xa(value)}"','edge="1"']
        if src: attrs.append(f'source="{src}"')
        if tgt: attrs.append(f'target="{tgt}"')
        self.cells.append(f'        <mxCell {" ".join(attrs)}>')
        self.cells.append('          <mxGeometry relative="1" as="geometry">')
        if source_xy:
            self.cells.append(f'            <mxPoint x="{source_xy[0]}" y="{source_xy[1]}" as="sourcePoint" />')
        if target_xy:
            self.cells.append(f'            <mxPoint x="{target_xy[0]}" y="{target_xy[1]}" as="targetPoint" />')
        if waypoints:
            self.cells.append('            <Array as="points">')
            for wx, wy in waypoints:
                self.cells.append(f'              <mxPoint x="{wx}" y="{wy}" />')
            self.cells.append('            </Array>')
        self.cells.append('          </mxGeometry>')
        self.cells.append('        </mxCell>')
        return cid

    def render(self):
        return "\n".join([
            '<mxfile host="Electron" agent="scripts/generate-tower-spec-diagrams.py" version="29.6.1">',
            f'  <diagram name="{html.escape(self.title)}" id="tower-spec">',
            f'    <mxGraphModel dx="1566" dy="1071" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="{self.width}" pageHeight="{self.height}" math="0" shadow="0" background="#ffffff">',
            '      <root>','        <mxCell id="0" />','        <mxCell id="1" parent="0" />',
            *self.cells,
            '      </root>','    </mxGraphModel>','  </diagram>','</mxfile>',
        ]) + "\n"

    def write(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(self.render(), encoding="utf-8")


# --- Style helpers (Style 1 Flat Icon) -------------------------------------
def sty_band(tint: str, *, dashed=False):
    t = TINTS[tint]
    d = "dashed=1;dashPattern=6 4;" if dashed else ""
    return (f"rounded=1;whiteSpace=wrap;html=1;fillColor={t['bg']};"
            f"strokeColor={t['border']};strokeWidth=1.3;arcSize=4;opacity=92;{d}")

def sty_card(tint_or_stroke: str = "#d1d5db", *, fill="#ffffff", arc=8,
             sw=1.4, fs=12, va="middle", al="center"):
    return (f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill};"
            f"strokeColor={tint_or_stroke};strokeWidth={sw};arcSize={arc};"
            f"fontSize={fs};fontColor={INK};fontFamily={FONT};"
            f"verticalAlign={va};align={al};spacingTop=4;spacingLeft=8;spacingRight=8;")

def sty_chip(fill, *, fc="#ffffff", fs=10.5):
    return (f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill};"
            f"strokeColor=none;fontColor={fc};fontSize={fs};"
            f"fontStyle=1;fontFamily={FONT};arcSize=50;verticalAlign=middle;align=center;")

def sty_chip_o(stroke, *, fc=None, fs=10.5):
    c = fc or stroke
    return (f"rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;"
            f"strokeColor={stroke};strokeWidth=1.2;fontColor={c};"
            f"fontSize={fs};fontStyle=1;fontFamily={FONT};arcSize=50;"
            f"verticalAlign=middle;align=center;")

def sty_edge(color, *, sw=1.6, dashed=False, end="classic", start="none",
             es="orthogonalEdgeStyle"):
    bits = [f"edgeStyle={es}","rounded=1","orthogonalLoop=1","jettySize=auto",
            "html=1",f"strokeColor={color}",f"strokeWidth={sw}",
            f"endArrow={end}","endFill=1","endSize=8",f"startArrow={start}",
            f"fontColor={MUTED}",f"fontFamily={FONT}","fontSize=10"]
    if dashed: bits += ["dashed=1","dashPattern=6 4"]
    return ";".join(bits)+";"

def sty_strip(accent):
    return (f"rounded=1;whiteSpace=wrap;html=1;fillColor={accent};"
            f"strokeColor=none;arcSize=20;fontColor=#ffffff;"
            f"fontFamily={FONT};verticalAlign=middle;align=center;")

def title_block(mf, title, subtitle, y=18):
    mf.text(0,y,mf.width,32,html.escape(title),size=20,bold=True,color=INK,align="center")
    mf.text(0,y+32,mf.width,20,
            f'<span style="font-style:italic;">{html.escape(subtitle)}</span>',
            size=11.5,color=MUTED,align="center")

def footer_note(mf, note, y=None):
    y = y if y is not None else mf.height-28
    mf.text(0,y,mf.width,20,html.escape(note),size=10,color=MUTED,italic=True,align="center")

def add_legend(mf, items, x, y):
    """items = [(color, label, dashed_bool), …]"""
    mf.rect(x, y, 260, 20+len(items)*22, "",
            f"rounded=1;whiteSpace=wrap;html=1;fillColor=#f9fafb;strokeColor=#e5e7eb;strokeWidth=1;arcSize=6;")
    mf.text(x+8, y+2, 200, 18, '<b>图例 Legend</b>', size=10, bold=True, color=MUTED)
    for i,(c,lbl,dash) in enumerate(items):
        ly = y+22+i*22
        d = "dashed=1;dashPattern=6 4;" if dash else ""
        mf.rect(x+10, ly+8, 30, 0, "",
                f"shape=line;html=1;strokeColor={c};strokeWidth=2;{d}fillColor=none;")
        mf.text(x+46, ly, 200, 20, html.escape(lbl), size=10, color=INK)

def chip_row(mf, chips, x, y, accent, *, h=22):
    cx = x
    for chip in chips:
        w = int(len(chip)*6.6+24) if all(ord(c)<128 for c in chip) else int(len(chip)*12+20)
        mf.rect(cx, y, w, h, html.escape(chip), sty_chip(accent, fs=10))
        cx += w + 8
    return cx


# =========================================================================
# Diagram 1 — System Architecture (6-layer + right sidebar)
# =========================================================================
def build_arch() -> Mxfile:
    mf = Mxfile("中国铁塔 · 工单质检系统架构图", 1440, 1050)
    title_block(mf,
        "工单质检系统 · 总体技术架构",
        "基于铁塔 AI 中台 + 高质量数据集平台 · 图像识别 + 语义识别 + 综合研判 · Style 1 Flat Icon")

    main_w = 1080
    side_x = main_w + 80
    side_w = 300
    mx = 40

    layers = [
        ("01","接入层","Access Layer","blue",
         [("PC Web 端","浏览器 · 响应式"),("移动端 H5","铁塔 App 内嵌"),("第三方 API","运维系统对接")]),
        ("02","应用层","Application Layer","indigo",
         [("数据接入与标注","工单/图片采集 · 多模态标注"),("模型训练与应用","质检模型管理 · 推理调度"),
          ("统计分析与追溯","报表 · 追溯查询 · 异常分析")]),
        ("03","服务层（核心）","Core Service Layer","orange",
         [("图像识别引擎","打卡照片合规检测"),("语义识别引擎","工单文本语义分析"),
          ("综合研判引擎","跨模态关联 · 规则融合"),("缺陷数据生成","自动扩充训练集")]),
        ("04","平台层","Platform Layer","green",
         [("铁塔 AI 中台","塔娃智能体 · 模型训练/推理"),("高质量数据集平台","标注 · 版本化 · 质量管控"),
          ("铁塔数据中台","统一存储 · 数据治理")]),
        ("05","数据层","Data Layer","slate",
         [("工单文本库","故障描述 · 处理结论"),("打卡照片存储","对象存储 · 分辨率归一化"),
          ("多模态数据集","打卡≥5万 · 维护≥5万"),("模型仓库","版本化模型资产")]),
        ("06","基础设施层","Infrastructure","slate",
         [("信创服务器","国产 CPU / NPU"),("K8s + 容器","集群化部署 · 弹性伸缩"),
          ("消息队列","Kafka · 异步解耦"),("关系数据库","信创兼容 · 高可用集群")]),
    ]

    tint_map = {"blue":"blue","indigo":"indigo","orange":"orange","green":"green","slate":"slate"}
    layer_h = 128
    layer_gap = 14
    y0 = 78
    layer_ids = []

    for i,(idx,zh,en,tint,comps) in enumerate(layers):
        ly = y0 + i*(layer_h+layer_gap)
        t = TINTS[tint]
        bid = mf.rect(mx, ly, main_w, layer_h, "", sty_band(tint))
        layer_ids.append(bid)
        mf.rect(mx+10, ly+10, 52, layer_h-20,
                f'<b style="font-size:20px;">{idx}</b>', sty_strip(t["accent"]))
        mf.text(mx+72, ly+10, 200, 22,
                f'<b>{html.escape(zh)}</b>', size=14, bold=True, color=t["accent"])
        mf.text(mx+72, ly+32, 200, 16, en, size=10, italic=True, color=MUTED)

        comp_w = (main_w - 100) // len(comps) - 10
        cx = mx + 72
        for j,(cn,cd) in enumerate(comps):
            mf.rect(cx + j*(comp_w+10), ly+54, comp_w, 62,
                    f'<b style="font-size:12px;color:{t["accent"]};">{html.escape(cn)}</b><br/>'
                    f'<span style="font-size:10px;color:{MUTED};">{html.escape(cd)}</span>',
                    sty_card(t["accent"], va="top", al="left")+"spacingTop=8;")

    # inter-layer edges
    edge_main = sty_edge(S1["flow_a"], sw=2.0)
    for i in range(len(layer_ids)-1):
        mf.edge(layer_ids[i], layer_ids[i+1], edge_main)

    # --- Right sidebar (cross-cutting) ---
    sb_y0, sb_y1 = y0, y0 + len(layers)*(layer_h+layer_gap) - layer_gap
    sb_h = sb_y1 - sb_y0
    mf.rect(side_x, sb_y0, side_w, sb_h, "",
            f"rounded=1;whiteSpace=wrap;html=1;fillColor=#f9fafb;strokeColor=#e5e7eb;"
            f"strokeWidth=1.2;arcSize=4;")
    mf.text(side_x, sb_y0+8, side_w, 22,
            '<b>横切关注点 · Cross-Cutting</b>', size=13, bold=True, color=INK, align="center")

    sidebar = [
        ("4A 统一认证授权","覆盖全栈 · RBAC","blue"),
        ("安全合规","等保 2.0 三级 · 传输加密 · 数据脱敏","red"),
        ("可观测性","IT 网管 · 统一日志 · 监控告警 · SOC","green"),
        ("DevOps & 治理","研发流水线 · ITSM · CMDB","indigo"),
        ("算法治理","模型版本 · 漂移检测 · A/B 验证","orange"),
    ]
    sb_card_h = (sb_h - 50) // len(sidebar) - 6
    for i,(lbl,desc,tint) in enumerate(sidebar):
        sy = sb_y0 + 38 + i*(sb_card_h+6)
        t = TINTS[tint]
        mf.rect(side_x+10, sy, side_w-20, sb_card_h,
                f'<b style="font-size:11.5px;color:{t["accent"]};">{html.escape(lbl)}</b><br/>'
                f'<span style="font-size:9.5px;color:{MUTED};">{html.escape(desc)}</span>',
                sty_card(t["border"], fill=t["bg"], arc=8, sw=1.2, fs=11, va="top", al="left")
                +"spacingTop=6;spacingLeft=8;spacingRight=8;")

    # --- KPI chips bottom ---
    kpi_y = sb_y0 + sb_h + 18
    mf.text(mx, kpi_y, 200, 20, '<b>关键性能指标</b>', size=12, bold=True, color=INK)
    kpis = ["准确率 ≥ 90%","召回率 ≥ 90%","100 并发 / 300 TPS",
            "RT < 500ms","99.5% SLA"]
    chip_row(mf, kpis, mx+140, kpi_y-2, TINTS["orange"]["accent"])

    footer_note(mf, "中国铁塔 2026 · 标段 1 工单质检 · 铁塔 AI 中台 + 高质量数据集平台 + 数据中台")
    return mf


# =========================================================================
# Diagram 2 — Functional Flow (9-stage closed loop)
# =========================================================================
def build_flow() -> Mxfile:
    mf = Mxfile("中国铁塔 · 工单质检功能流程图", 1440, 1050)
    title_block(mf,
        "工单质检 · 全生命周期功能流程图",
        "数据采集 → 标注 → 训练 → 推理 → 质检 → 追溯 → 反馈迭代 · 9 阶段闭环")

    # --- HITL strip top ---
    hitl_y = 78
    mf.rect(40, hitl_y, 1360, 42, "",
            f"rounded=1;whiteSpace=wrap;html=1;fillColor={TINTS['orange']['bg']};"
            f"strokeColor={TINTS['orange']['border']};strokeWidth=1.2;arcSize=4;dashed=1;dashPattern=6 4;")
    mf.text(50, hitl_y+8, 350, 24,
            '<b>HITL 人工质检参与点</b> '
            f'<span style="color:{MUTED};font-size:10px;">阶段 3 标注 · 阶段 7 异常复核</span>',
            size=12, bold=True, color=TINTS["orange"]["accent"])

    # HITL hooks pointing down (will be drawn after stage cards)
    hitl_hook_stages = [2, 6]  # 0-indexed: stage 3 & stage 7

    # --- 9 stages in 2 rows (4 + 5) ---
    stages = [
        ("01","数据采集与接入","Data Ingestion","blue",
         "打卡工单图片 · 故障工单文本 · 传感器/告警数据",
         ["打卡图片","故障文本","告警数据"]),
        ("02","数据标准化预处理","Preprocessing","indigo",
         "文本格式统一 · 图像分辨率归一化 · 异常值过滤",
         ["格式统一","归一化","异常过滤"]),
        ("03","多模态数据标注","Multimodal Labeling","orange",
         "图像文本对生成 · 业务标签打标 · 缺陷数据自动生成",
         ["图文对","业务标签","缺陷生成"]),
        ("04","高质量数据集构建","Dataset Build","green",
         "打卡数据集 ≥5 万 · 维护数据集 ≥5 万 · 版本化管理",
         ["≥5万打卡","≥5万维护","DVC"]),

        ("05","模型训练","Model Training","orange",
         "图像识别 + 语义识别 + 工单属性综合研判 · AI 中台",
         ["图像识别","语义识别","综合研判"]),
        ("06","模型推理与综合研判","Inference","blue",
         "文本质检 · 照片质检 · 属性质检 → 跨模态关联分析",
         ["文本质检","照片质检","跨模态"]),
        ("07","质检结果展示","Result Display","indigo",
         "实时质检结果 · 异常项醒目标识 · 多维度状态展示",
         ["实时展示","异常标识","多维度"]),
        ("08","追溯查询与统计分析","Analytics","green",
         "多条件查询 · 日/周/月/季报 · 异常原因分析报表",
         ["追溯查询","统计报表","异常分析"]),
        ("09","反馈采集与迭代","Feedback Loop","slate",
         "人工复核反馈 → 数据集回填 → 触发模型再训练",
         ["人工复核","数据回填","再训练"]),
    ]

    card_w, card_h = 290, 210
    row1_count, row2_count = 4, 5
    gap = 30
    row1_w = row1_count * card_w + (row1_count-1)*gap
    row2_w = row2_count * card_w + (row2_count-1)*gap
    row1_x0 = (mf.width - row1_w) // 2
    row2_x0 = (mf.width - row2_w) // 2
    row1_y = 145
    row2_y = row1_y + card_h + 60

    stage_ids = []
    stage_centers = []
    for i, (idx,zh,en,tint,desc,chips) in enumerate(stages):
        if i < row1_count:
            cx = row1_x0 + i*(card_w+gap)
            cy = row1_y
        else:
            j = i - row1_count
            cx = row2_x0 + j*(card_w+gap)
            cy = row2_y
        t = TINTS[tint]
        sid = mf.rect(cx, cy, card_w, card_h, "", sty_band(tint))
        stage_ids.append(sid)
        stage_centers.append((cx+card_w//2, cy+card_h//2, cx, cy))

        mf.rect(cx+10, cy+10, 44, 36,
                f'<b style="font-size:18px;">{idx}</b>', sty_strip(t["accent"]))
        mf.text(cx+62, cy+10, card_w-74, 20,
                f'<b>{html.escape(zh)}</b>', size=13, bold=True, color=t["accent"])
        mf.text(cx+62, cy+30, card_w-74, 16, en, size=9.5, italic=True, color=MUTED)

        mf.rect(cx+12, cy+54, card_w-24, 90,
                f'<span style="font-size:11px;color:{INK};line-height:1.6;">'
                f'{html.escape(desc)}</span>',
                sty_card(t["border"], fill="#ffffff", arc=6, sw=1.0, fs=11, va="top", al="left")
                +"spacingTop=8;spacingLeft=10;spacingRight=10;")

        chip_row(mf, chips, cx+12, cy+154, t["accent"])

    # --- Edges: row 1 sequential ---
    e_data = sty_edge(S1["flow_a"], sw=2.0, es="straight")
    for i in range(row1_count-1):
        mf.edge(stage_ids[i], stage_ids[i+1], e_data)

    # row1 last → row2 first (drop down)
    e_train = sty_edge(S1["flow_b"], sw=2.5, es="orthogonalEdgeStyle")
    mf.edge(stage_ids[3], stage_ids[4], e_train,
            value='<span style="background:#fff;padding:1px 4px;font-weight:bold;">训练</span>')

    # row 2 sequential
    for i in range(row1_count, len(stages)-1):
        e = e_data if i != 4 else sty_edge(S1["flow_b"], sw=2.2, es="straight")
        mf.edge(stage_ids[i], stage_ids[i+1], e if i > 4 else e)

    # stage 5→6 orange (推理)
    # already drawn above via sequential, but let's label it:
    # (labels for edges not easily editable after; skip for clarity)

    # --- Feedback loop: stage 9 → stage 3 (purple dashed) ---
    e_loop = sty_edge(S1["flow_d"], sw=2.0, dashed=True, es="orthogonalEdgeStyle")
    last_cx = stage_centers[8]
    first_cx = stage_centers[2]
    loop_y = row2_y + card_h + 50
    mf.edge(None, None, e_loop,
            value='<span style="background:#fff;padding:2px 6px;font-weight:bold;">持续迭代闭环 · 数据回填 + 再训练</span>',
            source_xy=(last_cx[0], row2_y + card_h),
            target_xy=(first_cx[0], row1_y + card_h),
            waypoints=[(last_cx[0], loop_y), (first_cx[0], loop_y)])

    # --- HITL hooks (stage 3 and stage 7) ---
    e_hitl = sty_edge(TINTS["orange"]["accent"], sw=1.4, dashed=True, es="straight")
    for si in hitl_hook_stages:
        scx = stage_centers[si][0]
        scy = stage_centers[si][3]  # top of card
        mf.edge(None, None, e_hitl,
                source_xy=(scx, hitl_y+42), target_xy=(scx, scy))

    # --- Legend (bottom-left) ---
    add_legend(mf, [
        (S1["flow_a"], "数据流 Data Flow", False),
        (S1["flow_b"], "训练 / 推理流 Train/Infer", False),
        (S1["flow_d"], "迭代回流 Feedback Loop", True),
        (TINTS["orange"]["accent"], "HITL 人工参与", True),
    ], 40, mf.height - 140)

    # --- KPI chips (bottom-right) ---
    kpi_y = mf.height - 72
    mf.text(780, kpi_y, 200, 20, '<b>核心验收 KPI</b>', size=12, bold=True, color=INK)
    kpis = ["识别准确率 ≥ 90%","召回率 ≥ 90%","标签覆盖 ≥ 90%","验证集 ≥ 1万条"]
    chip_row(mf, kpis, 940, kpi_y-2, TINTS["orange"]["accent"])

    footer_note(mf, "中国铁塔 2026 · 标段 1 工单质检 · 全流程 9 阶段闭环 · HITL 人机协作")
    return mf


# =========================================================================
# Diagram 3 — Network Topology (logical zones + XC/compliance)
# =========================================================================
def build_topology() -> Mxfile:
    mf = Mxfile("中国铁塔 · 工单质检网络拓扑图", 1440, 1100)
    title_block(mf,
        "工单质检系统 · 网络拓扑与信创部署",
        "逻辑分区 · 等保 2.0 三级安全边界 · 信创 CPU/NPU 全链路适配")

    # --- Overall enterprise boundary (red dashed) ---
    ent_x, ent_y, ent_w, ent_h = 250, 68, 1170, 620
    mf.rect(ent_x, ent_y, ent_w, ent_h, "",
            f"rounded=1;whiteSpace=wrap;html=1;fillColor=none;"
            f"strokeColor={S1['red']};strokeWidth=2;arcSize=2;dashed=1;dashPattern=8 4;")
    mf.text(ent_x+8, ent_y+4, 300, 18,
            f'<b style="color:{S1["red"]};font-size:11px;">等保 2.0 三级 安全边界</b>',
            size=11, bold=True, color=S1["red"])

    # --- External sources (left) ---
    ext_x, ext_y = 20, 120
    ext_w, ext_h = 200, 500
    mf.rect(ext_x, ext_y, ext_w, ext_h, "",
            f"rounded=1;whiteSpace=wrap;html=1;fillColor=#f9fafb;strokeColor=#d1d5db;"
            f"strokeWidth=1.2;arcSize=6;dashed=1;dashPattern=4 3;")
    mf.text(ext_x+8, ext_y+6, ext_w-16, 20,
            '<b>外部源 · External</b>', size=12, bold=True, color=INK)

    externals = [
        ("全国运维用户","100 并发","blue"),
        ("运营商主设备告警","告警数据接入","green"),
        ("铁塔动环告警","动环数据接入","green"),
        ("气象 API","环境数据","slate"),
        ("地理信息 API","GIS 数据","slate"),
    ]
    for i,(lbl,desc,tint) in enumerate(externals):
        ey = ext_y + 36 + i*88
        t = TINTS[tint]
        mf.rect(ext_x+12, ey, ext_w-24, 76,
                f'<b style="font-size:12px;color:{t["accent"]};">{html.escape(lbl)}</b><br/>'
                f'<span style="font-size:10px;color:{MUTED};">{html.escape(desc)}</span>',
                sty_card(t["border"], fill=t["bg"], arc=8, sw=1.2, fs=11, va="top", al="left")
                +"spacingTop=8;spacingLeft=8;")

    # --- Internal zones (left→right inside enterprise boundary) ---
    zone_x0 = ent_x + 16
    zone_y0 = ent_y + 28
    zone_h = ent_h - 48

    zones = [
        ("DMZ 边界区","防火墙 + WAF\n反向代理\n负载均衡（双活）","red", 150),
        ("接入与认证区","4A 统一认证网关\nAPI 网关","blue", 140),
        ("应用区\n信创 K8s 集群","工单质检 Web/App\n微服务集群\n容器管理平台","indigo", 190),
        ("AI / 中台区\n信创 + NPU","铁塔 AI 中台\n(塔娃 + 数据集 + 训练推理)\n铁塔数据中台","orange", 210),
        ("数据存储区","信创关系库集群\n对象存储（图片/数据集）\nKafka · 模型仓库","green", 190),
    ]

    zone_ids = []
    zx = zone_x0
    for lbl, desc, tint, zw in zones:
        t = TINTS[tint]
        zid = mf.rect(zx, zone_y0, zw, zone_h, "",
                f"rounded=1;whiteSpace=wrap;html=1;fillColor={t['bg']};"
                f"strokeColor={t['border']};strokeWidth=1.4;arcSize=4;opacity=92;")
        zone_ids.append(zid)
        mf.text(zx+6, zone_y0+8, zw-12, 40,
                f'<b style="font-size:12px;color:{t["accent"]};">{html.escape(lbl)}</b>',
                size=12, bold=True, color=t["accent"])
        lines = desc.split('\n')
        for li, line in enumerate(lines):
            mf.rect(zx+10, zone_y0+56+li*70, zw-20, 58,
                    f'<span style="font-size:11px;color:{INK};line-height:1.5;">{html.escape(line)}</span>',
                    sty_card(t["border"], fill="#ffffff", arc=6, sw=1.0, fs=11, va="middle", al="center"))
        zx += zw + 14

    # --- Edges between external→DMZ, DMZ→Auth, Auth→App, App→AI, AI→Data ---
    e_main = sty_edge(S1["flow_a"], sw=2.0)
    e_green = sty_edge(S1["flow_c"], sw=1.6)
    e_orange = sty_edge(S1["flow_b"], sw=1.6)

    # external → DMZ (main blue)
    mf.edge(None, zone_ids[0], e_main,
            source_xy=(ext_x+ext_w, ext_y+ext_h//2),
            target_xy=(zone_x0, zone_y0+zone_h//2))

    # DMZ → Auth → App → AI → Data
    for i in range(len(zone_ids)-1):
        e = e_orange if i == 0 else e_main  # DMZ→Auth is orange (4A)
        mf.edge(zone_ids[i], zone_ids[i+1], e,
                value='<span style="background:#fff;padding:1px 3px;font-size:9px;">'
                      + (['4A 鉴权','请求路由','模型调度','数据读写'][i]) + '</span>')

    # --- Ops management zone (bottom strip) ---
    ops_y = ent_y + ent_h + 18
    ops_h = 120
    mf.rect(ent_x, ops_y, ent_w, ops_h, "",
            f"rounded=1;whiteSpace=wrap;html=1;fillColor={TINTS['slate']['bg']};"
            f"strokeColor={TINTS['slate']['border']};strokeWidth=1.2;arcSize=4;")
    mf.text(ent_x+14, ops_y+8, 400, 20,
            '<b>运维管理区 · Operations Management</b>', size=13, bold=True, color=INK)
    ops_items = ["统一监控告警","统一日志平台","ITSM","CMDB","SOC","IT 研发流水线"]
    ox = ent_x + 14
    for item in ops_items:
        w = int(len(item)*12+24) if any(ord(c)>127 for c in item) else int(len(item)*7+24)
        mf.rect(ox, ops_y+36, w, 28, html.escape(item),
                sty_chip(TINTS["slate"]["accent"], fs=10.5))
        ox += w + 10

    mf.text(ent_x+14, ops_y+76, ent_w-28, 18,
            f'<span style="color:{MUTED};font-size:10px;">syslog / Filebeat → 日志转发 · '
            'Prometheus / 自定义 Exporter → 监控采集 · 全量运行日志 / 操作日志 / 审计日志</span>',
            size=10, color=MUTED)

    # dashed lines from app/AI/data zones down to ops
    e_gray = sty_edge("#9ca3af", sw=1.2, dashed=True, es="orthogonalEdgeStyle")
    for zid in zone_ids[2:]:  # app, AI, data
        mf.edge(zid, None, e_gray,
                target_xy=(ent_x+ent_w//2, ops_y))

    # --- Lower half: XC / compliance section ---
    xc_y = ops_y + ops_h + 28
    xc_h = mf.height - xc_y - 38

    # --- XC adaptation table ---
    mf.text(40, xc_y, 300, 22,
            '<b>信创全链路适配</b>', size=14, bold=True, color=INK)
    xc_items = [
        ("CPU","国产 CPU (鲲鹏/飞腾/海光)","✓"),
        ("操作系统","银河麒麟 / 统信 UOS","✓"),
        ("数据库","达梦 / openGauss / OceanBase","✓"),
        ("中间件","东方通 / 宝兰德","✓"),
    ]
    col_w = [100, 250, 60]
    tx = 40
    # header
    for ci,(cw,lbl) in enumerate(zip(col_w, ["组件层","国产方案","状态"])):
        mf.rect(tx, xc_y+28, cw, 28, f'<b style="color:#fff;font-size:11px;">{lbl}</b>',
                sty_card(TINTS["slate"]["accent"], fill=TINTS["slate"]["accent"],
                         arc=2, sw=0)+"fontColor=#ffffff;")
        tx += cw
    # rows
    for ri,(comp,solution,status) in enumerate(xc_items):
        ry = xc_y + 56 + ri * 30
        bg = "#ffffff" if ri%2==0 else "#f8fafc"
        tx = 40
        for ci,(val,cw) in enumerate(zip([comp,solution,status], col_w)):
            color = TINTS["green"]["accent"] if val=="✓" else INK
            weight = "bold" if val == "✓" else "normal"
            mf.rect(tx, ry, cw, 30,
                    f'<span style="font-size:11px;color:{color};font-weight:{weight};">'
                    f'{html.escape(val)}</span>',
                    sty_card("#e5e7eb", fill=bg, arc=0, sw=0.6, fs=11, al="center"))
            tx += cw

    # --- RTO/RPO chips ---
    rto_x = 500
    mf.text(rto_x, xc_y, 300, 22,
            '<b>容灾 / 可用性指标</b>', size=14, bold=True, color=INK)
    rto_chips = ["RTO < 30 min","RPO < 24 h","99.5% SLA","3 次容灾演练"]
    chip_row(mf, rto_chips, rto_x, xc_y+30, TINTS["blue"]["accent"])

    # --- Integration matrix ---
    im_x = 500
    im_y = xc_y + 68
    mf.text(im_x, im_y, 300, 22,
            '<b>内部平台对接矩阵</b>', size=13, bold=True, color=INK)
    integrations = ["4A","ITSM","CMDB","SOC","IT 网管","统一日志","统一监控","统一流程引擎"]
    chip_row(mf, integrations, im_x, im_y+26, TINTS["indigo"]["accent"])

    # --- Legend ---
    add_legend(mf, [
        (S1["flow_a"], "主请求路径 Main Flow", False),
        (S1["flow_b"], "认证拦截 4A Auth", False),
        (S1["flow_c"], "外部数据接入 Data In", False),
        ("#9ca3af", "日志/监控采集 Telemetry", True),
    ], 1140, xc_y)

    footer_note(mf,
        "中国铁塔 2026 · 标段 1 工单质检 · 等保 2.0 三级 · 信创全链路适配 · "
        "RTO<30min RPO<24h 99.5% SLA",
        y=mf.height-28)
    return mf


# =========================================================================
# Main
# =========================================================================
def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for name, builder in [
        ("tower-spec-01-system-architecture.drawio", build_arch),
        ("tower-spec-02-functional-flow.drawio", build_flow),
        ("tower-spec-03-network-topology.drawio", build_topology),
    ]:
        target = OUTPUT_DIR / name
        mf = builder()
        mf.write(target)
        print(f"✓ wrote {target} ({len(mf.cells)} mxCell lines)")

if __name__ == "__main__":
    main()
