# Account Audit Toolkit（Windows 版）

端到端账号审阅自动化流水线：从 WP Excel 输入，到最终审阅结论 Excel 输出，一键完成。

```text
WP / Lead Sheet
  → Scope Engine（范围判断）→ System-Folder Mapper（映射 supporting 文件夹）
  → Execution Plan Builder（账号审阅导出 → expected accounts，手工上传标识判断）
  → Evidence Checker（OCR 截图比对，仅手工上传标识=Y 时）
  → 最终审阅结论 Excel（是 / N/A / 否-xxx）
```

核心测试逻辑（`execution_plan_builder.decide_evidence_plan`）：

| 手工上传标识 | 测试方式 |
|---|---|
| N | 确认账号审阅导出中存在对应账号即通过 |
| Y | 账号存在 **且** 截图 OCR 比对通过 |
| 派生节点 | 开发人员清单 / 读写权限 / 应用层与 DB 管理员 SOD / 应用层管理员角色覆盖 |

## 项目结构

```text
├── main.py                    # 端到端流水线入口（一键生成审阅结论）
├── web_app.py                 # Web 界面
├── 01_install.bat / 02_start.bat  # Windows 一键安装 / 启动
├── config.py                  # WP 列名配置
├── decisions.py / text_tools.py / field_translator.py
├── rules_release_tool.py / rules_scope.py / audit_nodes.py
├── pipeline.py / evidence_task_router.py / exporter.py
├── execution_plan_builder.py  # expected accounts + 证据计划
├── evidence_checker.py        # OCR 截图比对
├── system_folder_mapper.py    # SID → supporting 文件夹映射
├── application_admin_role_config.py / sync_*.py / rules_*.py / sid_utils.py
├── configs/                   # 配置表（先看 configs/README.md）
├── templates/ / uploads/ / tests/
├── output/                    # 程序输出（不入库）
├── User_Manual_CN.md / Windows_User_Guide_CN.txt
└── requirements.txt
```

| 你要做什么 | 看哪里 |
|---|---|
| 运行端到端流水线 | `python main.py` |
| 看最终审阅结论 | `output/account_audit_conclusion.xlsx` |
| 改 WP 输入列名 | `config.py` |
| 改规则代码 | `rules_release_tool.py` / `rules_scope.py` |
| 维护映射表 | `configs/`（先看 `configs/README.md`） |

## 快速开始

```bash
pip install -r requirements.txt
python main.py                                  # 默认读取 test.xlsx
python main.py --input "你的WP文件.xlsx" --sheet Sheet1
```

Windows：双击 `01_install.bat` 安装依赖，再双击 `02_start.bat` 启动。

## 输出

- `output/account_audit_conclusion.xlsx`：测试范围 / 测试结果 / 明细-截图核对 / 明细-导出核对 / 明细-开发人员 / 明细-SOD / 明细-应用层管理员角色
- `execution_plan.xlsx`（中间产物）：expected_accounts / evidence_plan / compare_result
- `exporter.py` 单独运行时：todo Excel / JSON / JSON Schema

## 排查

| 问题 | 先看哪里 |
|---|---|
| 读不到 WP 列 | `config.py` |
| Remark 新写法识别不出 | `field_translator.py` |
| 发版工具组合未覆盖 | `rules_release_tool.py` |
| DB / SOD / 开发人员规则不对 | `rules_scope.py` |
| 特殊系统 | `configs/系统规则特殊覆盖表.xlsx` |
| 文件夹映射 / 账号别名 / sheet 映射 | `configs/` 对应映射表 |
| 输出格式 | `exporter.py` |

原则：能改配置表就不改代码；只有通用规则变了才改 Python。
