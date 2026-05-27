# 后续维护指南

你以后不要从头读所有代码，只按下面几个入口改。

## 1. WP 列名变了
改 `config.py` 里的 `AuditConfig`。

## 2. Remark 识别规则变了
改 `field_translator.py`。例如新增“容器化”“无应用层”“无数据库”等关键词。

## 3. 发版工具规则变了
改 `rules_release_tool.py`。重点看：

- `SYSTEM_RELEASE_OVERRIDES`：系统级特殊规则
- `_map_one_tool()`：PMS流程 + 发版工具 的组合规则

## 4. 数据库 / SOD / 开发人员规则变了
改 `rules_scope.py`。重点看：

- `SYSTEM_DB_OVERRIDES`：数据库相关系统特例
- `_apply_database_rules()`：数据库服务器 / 数据库账号规则
- `_apply_sod_rules()`：前后台管理员 SOD 规则

## 5. 文件夹名和 System 不一致
不要改代码，改 `configs/system_folder_mapping.xlsx`。

## 6. 英文账号、中文名、AD账号对不上
不要改代码，改 `configs/account_alias_mapping.xlsx`。

## 7. 某个检查节点应该去哪张 supporting sheet 找
不要改代码，改 `configs/supporting_sheet_mapping.xlsx`。

## 8. 某个系统需要人工覆盖判断
优先改 `configs/manual_override.xlsx`。后续如果这个特例很稳定，再考虑写入代码。

## 9. 输出格式想改
改 `exporter.py`。

## 推荐运行方式

```bash
cd account_audit_toolkit
pip install -r requirements.txt
python main.py
```

指定 WP 文件：

```bash
python main.py --input "../你的WP文件.xlsx" --sheet Sheet1
```

开发人员权限 & SOD 独立检查：

```bash
cd ..
python developer_permission_check.py
```
