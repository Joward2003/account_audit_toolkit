# Account Audit Toolkit 使用说明

> 端到端账号审阅自动化流水线。
>
> **从 WP Excel 输入，到最终审阅结论 Excel 输出，一键完成。**

---

## 1. 这个工具现在能做什么？

当前版本已完成完整流水线：

```text
WP / Lead Sheet (test.xlsx)
        ↓
Scope Engine 范围判断（哪些节点要查）
        ↓
System-Folder Mapper（映射到 supporting 文件夹）
        ↓
Execution Plan Builder（账号审阅导出 → expected accounts + 手工上传标识判断）
        ↓
Evidence Checker（OCR 截图比对，仅 手工上传标识=Y 时）
        ↓
最终审阅结论 Excel（是 / N/A / 否-xxx）
```

**核心测试逻辑（来自 `execution_plan_builder.decide_evidence_plan`）：**

| 手工上传标识 | 测试方式 |
|---|---|
| **N** | 只需确认账号审阅导出中存在对应账号 → 通过即「是」 |
| **Y** | 账号存在 **且** 截图 OCR 比对通过 → 「是」 |
| 派生节点 | 开发人员清单文件存在 / 开发人员无读写权限 / 应用层与DB管理员无 SOD 重合 |

同时还有独立脚本：

| 脚本 | 用途 |
|---|---|
| `main.py` | 端到端流水线，生成审阅结论 Excel |
| `developer_permission_check.py`（项目根目录） | 开发人员权限 & 管理员重合检查 |
| `execution_plan_builder.py` | 独立导出 execution_plan.xlsx |
| `evidence_checker.py` | 独立运行截图 OCR 比对 |

---

## 2. 项目结构怎么看？

```text
account_audit_toolkit/
├── main.py                         # 🔥 端到端流水线入口（一键生成审阅结论）
├── config.py                       # WP 列名配置
├── decisions.py                    # 是 / N/A / Refer to 的判断对象
├── text_tools.py                   # 文本清洗工具
├── field_translator.py             # 把 WP 文字翻译成程序能懂的事实
├── rules_release_tool.py           # 发版工具前台 / 后台规则
├── audit_nodes.py                  # 10 个检查节点列表
├── rules_scope.py                  # 总体审计范围判断规则
├── pipeline.py                     # 批量处理 WP（scope engine 管线）
├── evidence_task_router.py         # 把范围判断转成后续凭证任务
├── exporter.py                     # 导出 Excel / JSON / JSON Schema
│
├── execution_plan_builder.py       # 从账号审阅导出构建 expected accounts + 证据计划
├── evidence_checker.py             # OCR 截图比对（EvidenceLoader + OCRExtractor + EvidenceChecker）
├── system_folder_mapper.py         # System → supporting 文件夹映射
│
├── configs/                        # 配置表 / 映射表
│   ├── README.md                   # 配置表说明书，维护 Excel 前先看它
│   ├── system_folder_mapping.xlsx
│   ├── account_alias_mapping.xlsx
│   ├── supporting_sheet_mapping.xlsx
│   └── manual_override.xlsx
│
├── output/                         # 程序输出文件
│   ├── account_audit_conclusion.xlsx   # 最终审阅结论
│   ├── execution_plan.xlsx             # 执行计划（中间产物）
│   └── ...
│
├── 工作日志-2026-5-20.md
├── 工作日志-2026-5-21.md
├── 工作日志-2026-5-25.md
├── 工作日志-2026-5-26.md
├── README.md
├── MAINTAIN_GUIDE.md               # 维护指南（改哪里最快）
└── requirements.txt                # Python 依赖

../developer_permission_check.py    # 独立脚本：开发人员权限 & 管理员重合检查
```

平时你最常用的是：

| 你要做什么 | 看哪里 |
|---|---|
| 运行端到端流水线 | `main.py`（直接 `python main.py`） |
| 看最终审阅结论 | `output/account_audit_conclusion.xlsx` |
| 单独检查开发人员权限/SOD | `../developer_permission_check.py` |
| 单独导出执行计划 | `execution_plan_builder.py`（CLI） |
| 改 WP 输入列名 | `config.py` |
| 改规则代码 | `rules_release_tool.py` / `rules_scope.py` |
| 维护系统文件夹、账号别名、sheet 命名、人工覆盖 | `configs/`，尤其先看 `configs/README.md` |

---

## 3. 第一次怎么运行？

### 3.1 进入项目文件夹

```bash
cd account_audit_toolkit
```

### 3.2 安装依赖

```bash
pip install -r requirements.txt
```

### 3.3 运行端到端流水线

默认使用项目根目录的 `test.xlsx`：

```bash
python main.py
```

指定 WP 文件：

```bash
python main.py --input "../你的WP文件.xlsx" --sheet Sheet1
```

### 3.4 运行独立脚本

开发人员权限 & SOD 检查：

```bash
cd ..
python developer_permission_check.py
```

---

## 4. 跑完会生成什么？

### 4.1 `account_audit_conclusion.xlsx`（最终审阅结论）

这是最重要的输出文件，包含两张 sheet：

#### Sheet「审阅结论」

格式对齐 `test.xlsx` 模板，WP 原列 + 10 个审计节点结论 + 通过率。

每个节点结论为：
- **是** — 测试通过
- **N/A / N/A-xxx** — Scope 判断无需检查
- **否-账号缺失** — 账号审阅导出中无对应账号
- **否-截图检查未通过** — 手工上传标识=Y 但截图 OCR 未通过
- **待确认** — 需要人工判断

#### Sheet「Scope明细」

完整的 Pipeline 中间结果，包含每个节点的判断原因、warning 等，用于 debug。

---

### 4.2 历史输出（`exporter.py` 单独运行时）

`exporter.py` 单独运行时会生成：

| 文件 | 用途 |
|---|---|
| `账号审阅自动化todo.xlsx` | todo_list / program_template / detail 三张 sheet |
| `账号审阅自动化todo.json` | 给 Dify / LangGraph 等后续节点消费 |
| `账号审阅自动化todo.schema.json` | JSON 结构说明书 |

---

### 4.3 `execution_plan.xlsx`（中间产物）

`execution_plan_builder.py` 单独运行时生成：

| Sheet | 内容 |
|---|---|
| `expected_accounts` | 从账号审阅导出中按 审阅对象 筛出的预期账号 |
| `evidence_plan` | 每个节点的证据计划（手工上传标识 / ParseMethod / 是否需要截图） |
| `compare_result` | 比对结果（expected vs actual accounts） |

---

## 5. 后续配置表在哪里改？

配置表都在：

```text
configs/
```

请先看：

```text
configs/README.md
```

那里单独解释了：

- 每个 Excel 表是干嘛的；
- 每个字段代表什么意思；
- 哪些列后续代码会读取；
- 平时应该改哪一列；
- 哪些情况才需要改配置表。

这里主 README 不再展开配置表细节，避免两边内容重复。

---

## 6. 结果不对时，先看哪里？

按这个顺序排查：

| 问题 | 先看哪里 |
|---|---|
| 程序读不到 WP 列 | `config.py` |
| Remark 里的新写法识别不出来 | `field_translator.py` |
| 发版工具组合没有覆盖 | `rules_release_tool.py` |
| 数据库 / SOD / 开发人员规则不对 | `rules_scope.py` |
| 某个系统就是特殊 | `configs/manual_override.xlsx`，具体看 `configs/README.md` |
| 后续找不到文件夹 | `configs/system_folder_mapping.xlsx`，具体看 `configs/README.md` |
| 后续账号默认清洗后仍对不上 | `configs/account_alias_mapping.xlsx`，具体看 `configs/README.md` |
| 后续不知道去哪张 supporting sheet 找 | `configs/supporting_sheet_mapping.xlsx`，具体看 `configs/README.md` |
| 输出 Excel / JSON 想多一列或少一列 | `exporter.py` |

最重要原则：

```text
能改配置表，就先不要改代码。
只有通用规则真的变了，才改 Python。
```

---

## 7. 推荐推进顺序

当前已完成全部阶段。后续维护按需进行：

```text
✅ 第一步：跑通 WP → Scope Engine → todo_list / program_template
✅ 第二步：System-Folder Mapper（映射系统到 supporting 文件夹）
✅ 第三步：Execution Plan Builder（账号审阅导出 + 手工上传标识判断）
✅ 第四步：Evidence Checker（OCR 截图比对）
✅ 第五步：端到端 main.py（一键生成审阅结论 Excel）
✅ 第六步：developer_permission_check.py（开发人员权限 & SOD 独立脚本）
```

一句话：

```text
先把“该查什么”跑通，再解决“去哪查”，最后才做“自动检查截图”。
```

---

## 8. 后续 workflow 长什么样？

```text
WP / Lead Sheet
        ↓
Scope Engine
生成 program_template + todo_list + JSON
        ↓
System-Folder Mapper
匹配 System 和系统文件夹
        ↓
Evidence Loader
读取 supporting.xlsx / 图片 / PDF
        ↓
日常维护只需关注：

```text
- WP 列名变了 → config.py
- 规则变了 → rules_scope.py / rules_release_tool.py
- 文件夹/账号映射 → configs/
- 输出格式 → exporter.py
```

---

## 8. 完整流水线架构

```text
WP / Lead Sheet (test.xlsx)
        ↓
Scope Engine (rules_scope.py)
生成 program_template + todo_list + JSON
        ↓
System-Folder Mapper (system_folder_mapper.py)
匹配 System 和 supporting 文件夹
        ↓
Execution Plan Builder (execution_plan_builder.py)
从账号审阅导出构建 expected accounts
判断 手工上传标识（Y/N）→ 决定是否需要 OCR
        ↓
Evidence Checker (evidence_checker.py)
OCR 截图比对（仅 手工上传标识=Y 时）
        ↓
最终审阅结论 Excel (account_audit_conclusion.xlsx)
是 / N/A / 否-xxx
```

---

## 9. 最宝宝版记忆

```text
main.py                      = 一键启动，生成审阅结论
execution_plan_builder.py    = 账号比对逻辑（手工上传标识 Y/N → 要不要 OCR）
evidence_checker.py          = 截图 OCR 比对
developer_permission_check.py = 开发人员权限 & SOD（项目根目录）
rules_scope.py               = 判断要不要查
exporter.py                  = 把结果吐出来
output/                      = 放结果
configs/                     = 小字典，系统/文件夹/账号映射
```

你不需要一次看懂全部。

先跑 `python main.py`，看 `output/account_audit_conclusion.xlsx`，再按需维护 configs。
