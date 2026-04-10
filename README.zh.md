# fireworks-tech-graph

> 不用手画图了。用中文描述你的系统，几秒钟得到可直接发布的 SVG + PNG 技术图。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Claude Code Skill](https://img.shields.io/badge/Claude%20Code-Skill-blue)](https://claude.ai/code)
[![5 种视觉风格](https://img.shields.io/badge/风格-5种-purple)]()
[![8 种图类型](https://img.shields.io/badge/图类型-8种-green)]()

## 概述

`fireworks-tech-graph` 将自然语言描述转化为精美的 SVG 技术图，并通过 `rsvg-convert` 导出高分辨率 PNG。内置 **5 种视觉风格**，深度覆盖 AI/Agent 领域常见图类型（RAG、Agentic Search、Mem0、Multi-Agent、Tool Call 流程等）。

```
用户: "画一张 Mem0 的架构图，暗黑风格"
  → Skill 识别：Memory Architecture Diagram，Style 2
  → 生成含泳道、圆柱体、语义箭头的 SVG
  → 导出 1920px PNG
  → 输出路径：mem0-architecture.svg / mem0-architecture.png
```

---

## 效果展示

| 风格 | 示例图 |
|------|--------|
| **风格 1** — 扁平图标风（默认） | ![](assets/samples/sample-style1-flat.png) |
| **风格 2** — 暗黑极客风 | ![](assets/samples/sample-style2-dark.png) |
| **风格 3** — 工程蓝图风 | ![](assets/samples/sample-style3-blueprint.png) |
| **风格 4** — Notion 极简风 | ![](assets/samples/sample-style4-notion.png) |
| **风格 5** — 玻璃态卡片风 | ![](assets/samples/sample-style5-glass.png) |
| **AI 领域** — Mem0 架构图 | ![](assets/samples/sample-mem0.png) |

---

## 功能特性

- **5 种视觉风格** — 从白底极简到暗黑 Neon 再到磨砂玻璃
- **8 种图类型** — 架构图、数据流图、流程图、Agent 图、记忆图、序列图、对比图、思维导图
- **AI/Agent 领域内建知识** — RAG、Agentic Search、Mem0、Multi-Agent、Tool Call 等常见 Pattern 开箱即用
- **语义形状词汇表** — LLM = 双边框圆角矩形，Agent = 六边形，Vector Store = 带内环圆柱
- **语义箭头系统** — 颜色 + 虚线样式编码含义（写入/读取/异步/循环）
- **产品图标库** — 40+ 产品品牌色：OpenAI、Anthropic、Pinecone、Weaviate、Kafka、PostgreSQL……
- **泳道分组** — 自动为复杂架构添加层级标签
- **SVG + PNG 双输出** — SVG 可编辑，1920px PNG 可直接嵌入文章
- **rsvg-convert 兼容** — 纯内联 SVG，不依赖外部字体，渲染稳定

---

## 安装

```bash
claude skills install fireworks-tech-graph
```

或直接克隆：

```bash
git clone https://github.com/yizhiyanhua-ai/fireworks-tech-graph.git ~/.claude/skills/fireworks-tech-graph
```

---

## 安装依赖

```bash
# macOS
brew install librsvg

# Ubuntu/Debian
sudo apt install librsvg2-bin

# 验证安装
rsvg-convert --version
```

---

## 使用方式

### 触发词

以下关键词会自动触发 Skill：

```
画图 / 帮我画 / 生成图 / 做个图 / 架构图 / 流程图 / 可视化一下 / 出图
generate diagram / draw diagram / create chart / visualize
```

### 基本用法

```
画一张 RAG 流程图
```

```
生成一张 Agentic Search 架构图
```

### 指定风格

```
画一张微服务架构图，风格2（暗黑极客风）
```

```
生成 Multi-Agent 协作图，玻璃态风格
```

### 指定输出路径

```
生成 Mem0 架构图，输出到 ~/Desktop/
```

```
画一张 Tool Call 流程图 --output /tmp/diagrams/
```

---

## 5 种风格

| # | 名称 | 背景色 | 字体 | 适用场景 |
|---|------|--------|------|----------|
| 1 | **扁平图标风** *(默认)* | `#ffffff` | Helvetica | 博客、幻灯片、技术文档 |
| 2 | **暗黑极客风** | `#0f0f1a` | SF Mono / Fira Code | GitHub README、开发者文章 |
| 3 | **工程蓝图风** | `#0a1628` | Courier New | 架构设计文档、工程规范 |
| 4 | **Notion 极简风** | `#ffffff` | system-ui | Notion、Confluence、内部 Wiki |
| 5 | **玻璃态卡片风** | `#0d1117` 渐变 | Inter | 产品官网、演讲 Keynote |

每种风格在 `references/` 目录下都有专属参考文件，包含精确的颜色 Token、SVG 模板和使用规范。

---

## 支持的图类型

| 类型 | 描述 | 关键布局规则 |
|------|------|-------------|
| **架构图** | 服务、组件、云基础设施 | 水平分层，自上而下 |
| **数据流图** | 数据在系统中的流向 | 每条箭头标注数据类型 |
| **流程图** | 决策树、流程步骤 | 菱形 = 决策，自上而下 |
| **Agent 架构图** | LLM + 工具 + 记忆 | 五层模型：输入/Agent/记忆/工具/输出 |
| **记忆架构图** | Mem0、MemGPT 风格 | 读/写路径分离，记忆层级分明 |
| **序列图** | API 调用链、时序交互 | 垂直生命线，水平消息箭头 |
| **对比图** | 功能矩阵、方案比较 | 列 = 系统，行 = 属性 |
| **思维导图** | 概念地图、发散思维 | 中心节点，贝塞尔曲线分支 |

---

## AI/Agent 领域内建 Pattern

Skill 内置以下领域知识，可直接描述场景生成：

```
RAG Pipeline         → Query → Embed → VectorSearch → Retrieve → LLM → Response
Agentic RAG          → 在 RAG 基础上加入 Agent 循环 + 工具调用
Agentic Search       → Query → Planner → [Search/Calc/Code] → Synthesizer
Mem0 记忆层          → Input → Memory Manager → [VectorDB + GraphDB] → Context
Agent 记忆类型       → 感知记忆 → 工作记忆 → 情景记忆 → 语义记忆 → 程序记忆
Multi-Agent          → Orchestrator → [SubAgent×N] → Aggregator → Output
Tool Call 流程       → LLM → Tool Selector → Execution → Parser → LLM (循环)
```

---

## 形状词汇表

形状在所有风格中保持一致的语义：

| 概念 | 形状 |
|------|------|
| 用户 / 人类 | 圆形 + 身体路径 |
| LLM / 模型 | 圆角矩形，双边框，⚡ |
| Agent / 编排器 | 六边形 |
| 短期记忆 | 虚线边框圆角矩形 |
| 长期记忆 | 实线圆柱体 |
| Vector Store | 带内环圆柱 |
| Graph DB | 三圆簇 |
| 工具 / 函数 | 带 ⚙ 的矩形 |
| API / 网关 | 六边形（单边框） |
| 消息队列 / 流 | 横向管道 |
| 文档 / 文件 | 折角矩形 |
| 浏览器 / UI | 带三点标题栏的矩形 |
| 决策节点 | 菱形 |
| 外部服务 | 虚线边框矩形 |

---

## 箭头语义

| 流类型 | 线宽 | 虚线 | 含义 |
|--------|------|------|------|
| 主数据流 | 2px 实线 | — | 主要请求/响应路径 |
| 控制 / 触发 | 1.5px 实线 | — | 系统 A 触发 B |
| 记忆读取 | 1.5px 实线 | — | 从存储检索 |
| 记忆写入 | 1.5px | `5,3` | 写入/存储操作 |
| 异步 / 事件 | 1.5px | `4,2` | 非阻塞 |
| 反馈 / 循环 | 1.5px 曲线 | — | 迭代推理 |

---

## 文件结构

```
fireworks-tech-graph/
├── SKILL.md                      # 主 Skill 文件 — 图类型、布局规则、形状词汇
├── README.md                     # 英文文档
├── README.zh.md                  # 本文件（中文）
├── references/
│   ├── style-1-flat-icon.md      # 白底风格 — 彩色强调色
│   ├── style-2-dark-terminal.md  # 暗黑风格 — Neon 配色，等宽字体
│   ├── style-3-blueprint.md      # 蓝图风格 — 网格底纹，青色线条
│   ├── style-4-notion-clean.md   # 极简风格 — 白底，单色箭头
│   ├── style-5-glassmorphism.md  # 玻璃态风格 — 深色渐变，磨砂卡片
│   └── icons.md                  # 40+ 产品图标 + 语义形状模板
└── assets/
    └── samples/                  # 示例图 PNG
```

---

## 产品图标覆盖范围

**AI/ML 模型：** OpenAI、Anthropic/Claude、Google Gemini、Meta LLaMA、Mistral、Cohere、Groq、Hugging Face

**AI 框架：** Mem0、LangChain、LlamaIndex、LangGraph、CrewAI、AutoGen、DSPy、Haystack

**向量数据库：** Pinecone、Weaviate、Qdrant、Chroma、Milvus、pgvector、Faiss

**关系型/NoSQL 数据库：** PostgreSQL、MySQL、MongoDB、Redis、Elasticsearch、Neo4j、Cassandra

**消息队列：** Kafka、RabbitMQ、NATS、Pulsar

**云服务 & 基础设施：** AWS、GCP、Azure、Cloudflare、Vercel、Docker、Kubernetes

**可观测性：** Grafana、Prometheus、Datadog、LangSmith、Langfuse、Arize

---

## License

MIT © 2025 fireworks-tech-graph contributors
