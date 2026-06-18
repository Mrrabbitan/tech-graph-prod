#!/usr/bin/env python3
"""Generate a drawio (mxfile) XML of a multi-layer technical architecture diagram.

Takes a JSON config (either from stdin, a file path, or `--from` argument) and emits
an editable `.drawio` XML to the output path.

JSON schema (minimum):

    {
      "title":    "Diagram title",
      "subtitle": "Optional one-line subtitle",
      "page":     {"width": 1440, "height": 1050},
      "footer":   "Optional cross-cutting concerns line",
      "layers": [
        {
          "idx":        "01",
          "name":       "前端展示层",
          "en":         "Presentation Layer",
          "highlights": "...short highlights joined by '  ·  '...",
          "bg":     "#eff6ff",    // soft band fill
          "border": "#93c5fd",    // band border
          "accent": "#1d4ed8",    // primary color (left strip, component border, chip fill)
          "components": [
            {"title": "任务中心", "subtitle": "任务创建 · 编辑 · 启停"},
            ...   // 3-5 recommended
          ],
          "chips": ["Vue 3", "TypeScript", "Element Plus", ...]
        }
      ],
      "arrows": [    // optional; one fewer than layers. Labels for inter-layer bidirectional arrows
        "REST API · HTTPS", "MCP 协议 · 双向调用", ...
      ]
    }

Usage:

    python3 scripts/layered-to-drawio.py \\
        --config fixtures/layered-architecture-example.json \\
        --output output/my-arch.drawio

Or pipe JSON in via stdin:

    cat config.json | python3 scripts/layered-to-drawio.py -o out.drawio
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def xml_escape(s: str) -> str:
    return (
        s.replace('&', '&amp;')
         .replace('<', '&lt;')
         .replace('>', '&gt;')
         .replace('"', '&quot;')
    )


def _chip_width_px(text: str) -> int:
    """Crude pixel-width estimate for 10.5pt PingFang chip labels."""
    w = 22
    for ch in text:
        if ord(ch) > 0x2E80:
            w += 14
        elif ch.isupper():
            w += 8
        elif ch.isdigit():
            w += 7
        else:
            w += 6.5
    return max(int(w), 58)


def build_drawio(cfg: dict) -> str:
    page = cfg.get('page', {})
    PAGE_W = int(page.get('width', 1440))
    PAGE_H = int(page.get('height', 1050))
    title = cfg.get('title', 'Technical Architecture')
    subtitle = cfg.get('subtitle', '')
    footer = cfg.get('footer', '')
    layers = cfg['layers']
    arrows = cfg.get('arrows', [])

    MARGIN_X = 30
    CONTENT_W = PAGE_W - 2 * MARGIN_X
    LAYER_H = int(cfg.get('layer_height', 168))
    LAYER_GAP = int(cfg.get('layer_gap', 14))
    LAYER_START_Y = 86
    LEFT_STRIP_W = 62
    LABEL_X = MARGIN_X + 10 + LEFT_STRIP_W + 10
    LABEL_W = 258
    COMP_START_X = LABEL_X + LABEL_W + 20
    COMP_AREA_W = (MARGIN_X + CONTENT_W) - COMP_START_X
    COMP_GAP = 14
    COMP_H = 62
    CHIP_H = 22
    CHIP_Y_OFFSET = 132

    lines: list[str] = []
    lines.append('<mxfile host="app.diagrams.net" type="device" version="24.0.0">')
    lines.append(f'  <diagram name="{xml_escape(title)}" id="layered-arch">')
    lines.append(
        f'    <mxGraphModel dx="1600" dy="1100" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" '
        f'arrows="1" fold="1" page="1" pageScale="1" pageWidth="{PAGE_W}" pageHeight="{PAGE_H}" math="0" shadow="0">'
    )
    lines.append('      <root>')
    lines.append('        <mxCell id="0"/>')
    lines.append('        <mxCell id="1" parent="0"/>')

    def cell(cid, value, style, x, y, w, h):
        lines.append(
            f'        <mxCell id="{cid}" value="{xml_escape(value)}" style="{style}" parent="1" vertex="1">'
        )
        lines.append(f'          <mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/>')
        lines.append('        </mxCell>')

    # --- Title ---
    cell(
        'title', title,
        'text;html=1;align=center;verticalAlign=middle;strokeColor=none;fillColor=none;'
        'fontSize=22;fontStyle=1;fontFamily=PingFang SC,Helvetica,sans-serif;fontColor=#0f172a;',
        0, 18, PAGE_W, 32,
    )
    if subtitle:
        cell(
            'subtitle', subtitle,
            'text;html=1;align=center;verticalAlign=middle;strokeColor=none;fillColor=none;'
            'fontSize=13;fontColor=#64748b;fontFamily=PingFang SC,Helvetica,sans-serif;',
            0, 50, PAGE_W, 20,
        )

    layer_top_ys = []

    for i, layer in enumerate(layers):
        top = LAYER_START_Y + i * (LAYER_H + LAYER_GAP)
        bottom = top + LAYER_H
        layer_top_ys.append((top, bottom))

        bg = layer.get('bg', '#f8fafc')
        border = layer.get('border', '#cbd5e1')
        accent = layer.get('accent', '#334155')

        # Band
        cell(
            f'band_{i}', '',
            f'rounded=1;whiteSpace=wrap;html=1;fillColor={bg};strokeColor={border};'
            f'strokeWidth=1.3;arcSize=3;opacity=90;',
            MARGIN_X, top, CONTENT_W, LAYER_H,
        )
        # Left index strip
        cell(
            f'strip_{i}', f'<b style="font-size:24px;">{layer.get("idx", str(i+1).zfill(2))}</b>',
            f'rounded=1;whiteSpace=wrap;html=1;fillColor={accent};strokeColor=none;arcSize=6;'
            f'fontColor=#ffffff;fontSize=22;fontStyle=1;fontFamily=PingFang SC,Helvetica,sans-serif;'
            f'verticalAlign=middle;align=center;',
            MARGIN_X + 10, top + 14, LEFT_STRIP_W, LAYER_H - 28,
        )
        # Layer name
        cell(
            f'name_{i}', f'<b>{xml_escape(layer.get("name", ""))}</b>',
            f'text;html=1;align=left;verticalAlign=middle;strokeColor=none;fillColor=none;'
            f'fontSize=17;fontStyle=1;fontColor={accent};fontFamily=PingFang SC,Helvetica,sans-serif;',
            LABEL_X, top + 14, LABEL_W, 24,
        )
        # English subtitle
        if layer.get('en'):
            cell(
                f'en_{i}', layer['en'],
                'text;html=1;align=left;verticalAlign=middle;strokeColor=none;fillColor=none;'
                'fontSize=11;fontColor=#64748b;fontStyle=2;fontFamily=Helvetica,sans-serif;',
                LABEL_X, top + 40, LABEL_W, 16,
            )
        # Highlights box
        highlights_val = (
            f'<b style="font-size:10px;color:#64748b;">技术亮点</b>'
            f'<br/><span style="font-size:11px;color:#1e293b;line-height:1.6;">'
            f'{xml_escape(layer.get("highlights", ""))}</span>'
        )
        cell(
            f'hl_{i}', highlights_val,
            f'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor={border};'
            f'strokeWidth=1;arcSize=8;fontSize=11;fontColor=#1e293b;'
            f'fontFamily=PingFang SC,Helvetica,sans-serif;verticalAlign=top;align=left;'
            f'spacingTop=4;spacingLeft=8;spacingRight=8;',
            LABEL_X, top + 62, LABEL_W, 90,
        )

        # Components row
        components = layer.get('components', [])
        num = max(len(components), 1)
        comp_w = (COMP_AREA_W - (num - 1) * COMP_GAP) / num
        for j, comp in enumerate(components):
            cx = COMP_START_X + j * (comp_w + COMP_GAP)
            cy = top + 20
            val = (
                f'<b style="font-size:13px;color:{accent};">{xml_escape(comp.get("title", ""))}</b>'
                f'<br/><span style="font-size:10.5px;color:#475569;line-height:1.4;">'
                f'{xml_escape(comp.get("subtitle", ""))}</span>'
            )
            cell(
                f'comp_{i}_{j}', val,
                f'rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor={accent};'
                f'strokeWidth=1.4;arcSize=12;fontSize=12;fontColor=#1e293b;'
                f'fontFamily=PingFang SC,Helvetica,sans-serif;verticalAlign=middle;align=center;',
                cx, cy, comp_w, COMP_H,
            )

        # Tech-stack label
        cell(
            f'ts_label_{i}', '<b>技术栈</b>',
            'text;html=1;align=left;verticalAlign=middle;strokeColor=none;fillColor=none;'
            'fontSize=10;fontColor=#64748b;fontStyle=1;fontFamily=PingFang SC,Helvetica,sans-serif;',
            COMP_START_X, top + CHIP_Y_OFFSET - 2, 60, 18,
        )
        # Chips
        chip_x = COMP_START_X + 56
        chip_y = top + CHIP_Y_OFFSET
        for k, chip_text in enumerate(layer.get('chips', [])):
            w = _chip_width_px(chip_text)
            cell(
                f'chip_{i}_{k}', chip_text,
                f'rounded=1;whiteSpace=wrap;html=1;fillColor={accent};strokeColor=none;'
                f'fontColor=#ffffff;fontSize=10.5;fontStyle=1;'
                f'fontFamily=PingFang SC,Helvetica,sans-serif;arcSize=50;'
                f'verticalAlign=middle;align=center;',
                chip_x, chip_y, w, CHIP_H,
            )
            chip_x += w + 8

    # Arrows between layers
    arrow_x_center = PAGE_W / 2
    for idx in range(len(layers) - 1):
        src_bottom = layer_top_ys[idx][1]
        tgt_top = layer_top_ys[idx + 1][0]
        label = arrows[idx] if idx < len(arrows) else ''
        lines.append(
            f'        <mxCell id="arr_{idx}" value="{xml_escape(label)}" '
            f'style="endArrow=classic;startArrow=classic;html=1;strokeColor=#475569;strokeWidth=2;'
            f'fontSize=10.5;fontColor=#334155;fontFamily=PingFang SC,Helvetica,sans-serif;'
            f'labelBackgroundColor=#ffffff;rounded=0;" edge="1" parent="1">'
        )
        lines.append('          <mxGeometry relative="1" as="geometry">')
        lines.append(f'            <mxPoint x="{arrow_x_center}" y="{src_bottom - 1}" as="sourcePoint"/>')
        lines.append(f'            <mxPoint x="{arrow_x_center}" y="{tgt_top + 1}" as="targetPoint"/>')
        lines.append('          </mxGeometry>')
        lines.append('        </mxCell>')

    # Footer
    if footer:
        footer_y = PAGE_H - 44
        cell(
            'footer_bg', '',
            'rounded=1;whiteSpace=wrap;html=1;fillColor=#f1f5f9;strokeColor=#cbd5e1;'
            'strokeWidth=1;arcSize=6;',
            MARGIN_X, footer_y, CONTENT_W, 30,
        )
        cell(
            'footer', footer,
            'text;html=1;align=center;verticalAlign=middle;strokeColor=none;fillColor=none;'
            'fontSize=11;fontColor=#475569;fontFamily=PingFang SC,Helvetica,sans-serif;',
            MARGIN_X, footer_y, CONTENT_W, 30,
        )

    lines.append('      </root>')
    lines.append('    </mxGraphModel>')
    lines.append('  </diagram>')
    lines.append('</mxfile>')
    return '\n'.join(lines)


def load_config(args) -> dict:
    if args.config:
        with open(args.config, 'r', encoding='utf-8') as f:
            return json.load(f)
    data = sys.stdin.read().strip()
    if not data:
        raise SystemExit('No JSON config provided (use --config or stdin).')
    return json.loads(data)


def main() -> int:
    ap = argparse.ArgumentParser(description='Generate layered-architecture drawio XML.')
    ap.add_argument('--config', '-c', help='Path to JSON config file.')
    ap.add_argument('--output', '-o', required=True, help='Output .drawio path.')
    args = ap.parse_args()

    cfg = load_config(args)
    xml = build_drawio(cfg)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(xml, encoding='utf-8')
    print(f'✓ drawio written: {out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
