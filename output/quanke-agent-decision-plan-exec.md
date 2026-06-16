# 全客存量客户智能策略生成 — 决策 / 规划 / 执行 智能体设计材料

**依据材料**：《A计划存量客户智能策略生成》工作方案、`业务流程(3).pptx`（智能客户洞察 / 智能策略策划 / 智能策略生成 / 触点推荐流程）。

**产出说明**：本文将原方案中的「洞察 / 策划 / 生成」能力，抽象为统一的 **决策（Decision）→ 规划（Planning）→ 执行（Execution）** 三智能体分工；配套 **五层记忆**（感知 / 工作 / 情景 / 语义 / 程序）与 **语义化数据流箭头**，用于评审与迭代。

---

## 1. 全客场景背景与目标

### 1.1 业务痛点（摘录）

| 维度 | 数据表现 / 现象 | 痛点 |
|------|-----------------|------|
| 客群圈选 | 全客客户群新增量大，总部+省分标签 2000+ | 标签组合门槛高、圈选慢、难支撑多样化场景 |
| 策略策划 | 策划涉及客群/政策/产商品/触点/话术等，最少操作约 61 步 | 多系统、多环节、依赖个人经验，响应业务慢 |
| 营销转化 | 触达用户规模大，策略覆盖率与营业员推荐率有提升空间 | 部分用户无策略覆盖，或策略与用户意图不匹配，转化受限 |

### 1.2 业务与牵引目标（KPI 摘要）

**基础目标**

| 类别 | 指标 | 定义要点 | 2026 目标 |
|------|------|----------|-----------|
| 执行效率 | 重点场景智能策略策划占比 | 智能策划创建数 / 重点场景月创建数 | ≥10% |
| 业务效果 | 中国联通 app 重点场景策略订单转化率提升 | (2026−2025)/2025 | 提升 20% |
| 业务效果 | 中国联通 app 产能提升 | 累计产能差（收入增量规则见方案） | +1500 万元 |

**牵引目标（节选）**

| 指标 | 目标方向 |
|------|----------|
| 智能客群生成时间 | 较人工点选缩短 ≥40% |
| 智能圈选客群占比 | ≥10%（试点省） |
| 智能推荐标签使用占比 | ≥10% |
| 策略配置时长 | ≤3 分钟（对话→完整方案） |
| 人工操作步骤 | ≤5 步 |
| 策略转化率提升（AI vs 人工） | ≥10% |
| 进厅用户策略覆盖率（CB 智慧推荐） | 提升 ≥10% |

---

## 2. 智能体重组：决策 / 规划 / 执行 与原文档能力映射

| 新分工 | 职责摘要 | 对应原能力（pptx / 工作方案） |
|--------|----------|------------------------------|
| **决策智能体** | 意图识别、场景判定（L1/L2）、槽位抽取、澄清与路由 | 「智能客户洞察」中的语义解析与场景识别；策略策划「意图识别与场景理解」；动态策略「无意图/有意图」分流 |
| **规划智能体** | RAG 检索历史最优、知识图谱推理、客群/产品/触点/话术组装、档位与约束校验 | 「智能策略策划」策略组装与 MCP 工具编排；「智能策略生成」图谱召回与排序；标签推荐与组合优化 |
| **执行智能体** | MCP 工具调用、客群创建、策略配置写入全客、触点下发、效果数据采集 | 工具执行层 + 原子能力层（查询/写入/话术模板）；CB 弹窗 / 10010 等触点闭环 |

**与 pptx 流程对齐**：客户触点 →（覆盖/匹配判断）→ 策略生成 + 话术推荐 → 办理 → 运营复盘；其中「策略与意图未匹配」路径触发 **决策 + 规划** 的实时或会话内重规划。

---

## 3. 三智能体输入 / 输出契约

### 3.1 总览

| 智能体 | 主要输入 | 主要输出 | 下游 |
|--------|----------|----------|------|
| 决策 | 自然语言 / 语音文本 / 弹窗事件、画像快照、会话状态 | 意图、置信度、场景、槽位、路由 | 规划 或 澄清 或 直达执行（极少） |
| 规划 | 决策输出、RAG 片段、Neo4j 子图、规则约束 | 结构化 `StrategyPlan`（可映射策略原子字段） | 执行 |
| 执行 | `StrategyPlan`、环境凭证 | `strategy_id`、发布状态、推送记录、效果指标 | 情景记忆 / 效果库 |

### 3.2 决策智能体 — JSON Schema（示例）

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "DecisionAgentOutput",
  "type": "object",
  "required": ["intent", "confidence", "route"],
  "properties": {
    "intent": {
      "type": "string",
      "enum": ["策略策划", "客户洞察", "策略生成", "无意图"]
    },
    "confidence": { "type": "number", "minimum": 0, "maximum": 1 },
    "scene_l1": { "type": "string", "description": "一级场景 e.g. 宽带感知提升" },
    "scene_l2": { "type": "string", "description": "二级场景 e.g. 千兆升级" },
    "slots": {
      "type": "object",
      "properties": {
        "startDate": { "type": "string", "format": "date" },
        "endDate": { "type": "string", "format": "date" },
        "businessScene": { "type": "string" },
        "disturbType": { "type": "string" },
        "netType": { "type": "string" },
        "channelCode": { "type": "string" },
        "custgroupHint": { "type": "string" }
      },
      "additionalProperties": true
    },
    "route": {
      "type": "string",
      "enum": ["planning", "execution", "clarify"]
    },
    "clarify_question": { "type": "string" }
  }
}
```

### 3.3 规划智能体 — JSON Schema（示例）

```json
{
  "title": "PlanningAgentOutput",
  "type": "object",
  "required": ["strategy_plan_version", "elements"],
  "properties": {
    "strategy_plan_version": { "type": "string" },
    "elements": {
      "type": "object",
      "properties": {
        "custgroup": { "type": "object", "description": "客群条件或客群ID" },
        "products": { "type": "array", "items": { "type": "object" } },
        "channel": { "type": "object" },
        "sellingTitle": { "type": "string" },
        "sellingPoint": { "type": "string" },
        "scriptText": { "type": "string" },
        "executionParams": {
          "type": "object",
          "description": "对齐全客原子配置项：busiType、strategyType、policy、触点字段等"
        }
      }
    },
    "evidence": {
      "type": "object",
      "properties": {
        "rag_hits": { "type": "array" },
        "kg_path": { "type": "string" }
      }
    }
  }
}
```

### 3.4 执行智能体 — JSON Schema（示例）

```json
{
  "title": "ExecutionAgentOutput",
  "type": "object",
  "properties": {
    "strategy_id": { "type": "string" },
    "publish_status": { "type": "string" },
    "push_records": { "type": "array" },
    "outcome": {
      "type": "object",
      "properties": {
        "workorder_count": { "type": "integer" },
        "contact_count": { "type": "integer" },
        "order_count": { "type": "integer" },
        "conversion_rate": { "type": "number" }
      }
    }
  }
}
```

---

## 4. 提升策略有效性的数据：省分采集 + 全客历史（规则 / 模型 / 知识库）

### 4.1 分类视图

| 类型 | 内容 | 典型来源 | 用途 |
|------|------|----------|------|
| **规则** | 场景—标签规则、档位/激励/约束、免打扰与产品互斥、保底推荐规则 | 省分业务侧、全客策略配置约束 | 硬校验、排序、过滤违规推荐 |
| **模型** | 意图与实体识别、工具调用(Function Calling)、场景分类(LightGBM/XGBoost)、Embedding | 共享算力模型 + 省分样本 | 解析、召回、排序、预测 |
| **知识库** | 标签元数据与场景映射、历史策略切片向量库、场景最优策略库、Neo4j 产品图谱、话术模板 | 标签库、全客、省分语料 | RAG、图谱多跳、话术生成 |

### 4.2 省分需配合提供（与场景清单对齐，节选）

| 场景方向 | 省分配合资料（示例） |
|----------|----------------------|
| 控流失 / 套餐迁转等 | 产品特征文档、场景预设产品与触点、热门场景对话标注、热门话术 |
| 宽带感知 / 单转融等 | 场景标签规则、档位/激励/约束规则；动态策略业务与数据梳理 |
| 通用 | 热点/保底产品清单、意图评审语料（如来话意图打标）、画像标签加工与上架 |

### 4.3 全客与标签侧历史数据（节选）

| 数据 | 说明 |
|------|------|
| 历史策略配置 | 按场景分类的策略 JSON / 原子字段，支持向量化切片 |
| 策略运营效果 | 触达、接触、订购及时间字段，用于 TOP-K 最优策略与 10 档效能评估 |
| 标签元数据 | 编码、口径、分布、检索热度；支撑标签推荐与治理 |
| 对话与采纳 | 用户提问、模型回答、采纳/修改/原因；用于知识库迭代与反例 |

### 4.4 合规与字段建议

- 最小必要原则采集；敏感字段脱敏与权限分级；血缘记录（省分版本、生效账期）。
- 统一主键：策略 ID、产品 ID、标签编码、触点编码；时间字段与账期对齐。

---

## 5. 输出验证与准确度优化

### 5.1 验证方式（五类）

| 方式 | 做法 | 参考目标 / 备注 |
|------|------|-----------------|
| 离线自动评估 | 意图分类、实体抽取、工具调用命中率 | 方案目标：意图 ≥90%，工具调用 ≥85% |
| 抽样专家评测 | 抽 10% 复杂样本，业务评审可执行性与合规 | 与自动化指标互补 |
| A/B 与网格搜索 | 温度、TopK、提示词版本、阈值 | 多版本方案对比选优 |
| 线上对照实验 | 同场景 AI 策略 vs 人工策略 | 转化率提升目标 ≥10% |
| 业务回流 | W/C/O 与 D_CREATE 等漏斗进入 10 档效能口径 | 驱动劣化策略下线与 RAG 负样本 |

### 5.2 准确度优化手段

- **提示词与知识协同**：意图定义库注入 + 少样本动态分类；Schema 约束的 JSON 抽取。
- **RAG**：原子配置项切片、场景最优库单独索引；召回阈值（如置信度 0.8）与重排序。
- **图谱**：意图—场景—特征—产品—约束路径可解释输出，减少幻觉办理。
- **小模型 + 标签治理**：`feature_importances_` 反哺标签口径；对话未采纳回流纠错。
- **人在环路（HITL）**：发布前二次确认、高风险触点强制审核。

### 5.3 策略执行效能分档（10 档 — 用于数据验证与筛选 TOP 策略）

> 下列为工作方案中的**自动化归类思路**摘要，用于「观察 / 异常 / 无效 / 低效 / 高效 / 最优」等分桶；线上实现时以正式口径表为准。

| 优先级 | 一级 | 二级 | 规则要点（自然语言） |
|--------|------|------|----------------------|
| 1 | 观察 | 新建观察 | 创建时间短，尚处冷启动观察期 |
| 2 | 异常 | 数据异常 | 接触、订购、工单等字段逻辑矛盾 |
| 3 | 无效 | 无效 | 长周期无接触无转化等 |
| 4 | 低效 | 零转化 | 有接触无订购等 |
| 5 | 低效 | 微转化 | 极低订购/接触比等 |
| 6 | 低效 | 双低 | 转化与接触占比双低 |
| 7 | 高效 | 高转化 | 大规模样本下高 O/C 或高订购量 |
| 8 | 最优 | 最优 | 高效策略中按 O/C 等排序 TOP5 |
| 9 | 正常 | 常规 | 稳定期达标的常规表现 |
| 10 | 其他 | — | 不满足以上分类 |

（原始文档含 `D_CREATE`、`W`、`C`、`O` 等字段表达式，落地时由数据团队固化 SQL/规则引擎。）

---

## 6. 后续迭代方向与方案

| 维度 | 方向 | 方案要点 |
|------|------|----------|
| 数据 | 多省语料与反例 | 扩充 100+ 热门对话/省；沉淀「未采纳原因」负样本 |
| 模型 | 降本与在线学习 | 意图蒸馏小模型；话术 bandit / 多臂实验 |
| 知识 | KG 与标签治理 | 增量构建管道；跨省冲突检测与版本合并 |
| 平台 | 工具与可观测 | MCP 工具注册中心；全链路 trace；HITL 埋点 |
| 治理 | 策略生命周期 | 按 10 档自动归档；低效自动下线；免打扰合规审计 |

---

## 7. 五智能体上下游流程图（1 决策 + 1 规划 + 3 执行）

为支撑「执行环节」的高并行 + 强治理，将原「执行智能体」拆为 3 个职能内聚、可独立扩缩容、具备各自 SLO 的子智能体，形成 **决策 → 规划 → 三执行** 的标准上下游：

| 智能体 | 主要职责 | 上游 | 下游 / 写入 |
|--------|----------|------|-------------|
| AGENT-1 决策 | 意图分类 / 场景判定 / 槽位抽取 / 路由（planning · execution · clarify） | 触点输入（CB弹窗 / 10010 ASR / 全客 NLU / 画像快照） | 规划（默认）；clarify 回路反馈到输入 |
| AGENT-2 规划 | RAG + Neo4j 召回 → 客群+产品+触点+话术 组装 → 档位排序 + 约束校验 | 决策的结构化意图 + L4 语义 + L3 情景 | 三执行（StrategyPlan 扇出） |
| AGENT-3a 客群圈选执行 | MCP 圈选 / 客群生成 / 规模预估；写客户群表 | 规划 | 全客客户群、L3 客群快照；输出 `custgroup_id` 给 3b |
| AGENT-3b 策略写入与触点下发 | 策略 JSON 写入全客；CB 弹窗 / 10010 / 外呼 / 工单 触点下发；免打扰/互斥校验 | 规划 + 3a (`custgroup_id`) | 触点系统、L3 推送记录；输出 `strategy_id` / `push_records` 给 3c |
| AGENT-3c 效果回流与治理 | 采集 W/C/O 漏斗 + D_CREATE → 10 档自动归类、TOP-K 入选；标签 / Prompt 治理回流；A/B 写 L3 | 3b 推送结果 + 全客效果库 | L3 情景、L5 程序（治理）；闭环触发决策的「重规划 / 下次会话个性化」 |

**关键箭头语义（与图例一致）**：

- 紫色：控制流（决策 → 规划 → 三执行的扇出）
- 黄色：决策路由（diamond `route?`）
- 青色：检索（智能体读 L1–L5 / 数据底座）
- 绿色虚线：写记忆 / 治理（3c → L3 情景，3c → L5 程序）
- 橙色：重规划 / 澄清回路（clarify 回到输入 + 3c 闭环回到决策）
- 白色：原始数据 ingest（触点 → 感知）

**调用时序示例（成功路径）**：

```text
USER → DECISION:
   {intent: "策略策划", scene_l1:"宽带感知提升",
    slots:{startDate, endDate, businessScene, channelCode, ...},
    route:"planning"}

DECISION → PLANNING:
   structured_intent + portrait_snapshot

PLANNING → (RAG@L3 + KG@L4 + RuleEngine@L5)
PLANNING → StrategyPlan {custgroup_cond, products[], channel, scriptText, executionParams}

StrategyPlan ─┬─▶ 3a 客群圈选执行  → custgroup_id
              └─▶ 3b 策略写入与触点下发 (depends_on=custgroup_id)
                                       → strategy_id, push_records
                  ─▶ 3c 效果回流与治理   → W/C/O, conversion_rate, 10档归类
                                       → 写 L3 情景 + L5 治理 → DECISION 重规划
```

**异常 / 澄清路径**：

- `clarify`：决策智能体置信度不足或槽位缺失 → 回写到工作记忆 L2 → 回路到输入层等待用户补全；
- `execution-only`：用户已给出完整方案 → 跳过规划直达 3b（受决策置信度与权限策略门控）；
- 3c 监测劣化（10 档 = 无效 / 低效）→ 触发自动下线 + 反例语料回流。

**图例**：

- 矢量：[`output/quanke-five-agent-flow.svg`](output/quanke-five-agent-flow.svg)
- 高分 PNG：[`output/quanke-five-agent-flow.png`](output/quanke-five-agent-flow.png)
- 可编辑 drawio：[`output/quanke-five-agent-flow.drawio`](output/quanke-five-agent-flow.drawio)（由 [`fixtures/quanke-five-agent-flow.json`](fixtures/quanke-five-agent-flow.json) 通过 `layered-to-drawio.py` 生成）

![全客存量 · 五智能体上下游流程](output/quanke-five-agent-flow.png)

---

## 8. 总架构图（Blueprint：底色 + 语义箭头 + 五层记忆）

下列 PNG 与矢量文件路径：

- `output/quanke-three-agent-memory.svg`（蓝图矢量）
- `output/quanke-three-agent-memory.png`（Chrome headless 2880×2100 导出）
- `output/quanke-three-agent-memory.drawio`（分层可编辑，由 `fixtures/quanke-three-agent-memory.json` 生成）

![全客存量 · 三智能体 × 五层记忆 · 总架构](output/quanke-three-agent-memory.png)

**图例（语义箭头）**：紫色控制流、青色检索、绿色虚线效果写入、橙色重规划回路、白色触点原始数据 — 详见图内 LEGEND。

---

## 附录 A：分层架构 JSON 配置（drawio / 评审）

同源文件：[`fixtures/quanke-three-agent-memory.json`](../fixtures/quanke-three-agent-memory.json)

其中 `_meta` 字段包含：`arrow_semantics`、`memory_tiers`、`agents` 与产出物路径，便于与架构图、需求基线一并版本管理。

---

## 附录 B：制图与导出命令（工程复现）

```bash
# 总架构图（三智能体 × 五层记忆，蓝图 SVG）
python3 scripts/generate-quanke-blueprint-svg.py

# 五智能体上下游流程图（1 决策 + 1 规划 + 3 执行，蓝图 SVG）
python3 scripts/generate-quanke-five-agent-flow-svg.py

# 可编辑 drawio（两张图分别对应两份 JSON）
python3 scripts/layered-to-drawio.py \
  -c fixtures/quanke-three-agent-memory.json \
  -o output/quanke-three-agent-memory.drawio
python3 scripts/layered-to-drawio.py \
  -c fixtures/quanke-five-agent-flow.json \
  -o output/quanke-five-agent-flow.drawio

# PNG（需本机 Chrome；将 object 嵌入 HTML 以保留纵横比）
# 见 README 与 output/wrap.html 模板或自行用 draw.io CLI 导出
```

---

**文档版本**：v1.0 · 与《A计划存量客户智能策略生成》及业务流程 pptx 对齐，供评审后按省分试点迭代。
