# configs 配置表宝宝说明书

> 这个文件夹里的 Excel 不是原始审计数据，而是给后续程序看的“小字典”。
>
> 你的主程序负责先生成：`program_template`、`todo_list`、`JSON`。
> 这些配置表主要给下一阶段用：找文件夹、找 supporting sheet、统一账号名、处理人工特例。

---

## 0. 先记住一句话

```text
普通情况：程序走默认规则。
特殊情况：才来这里改 Excel。
```

也就是说，这些配置表不是让你把所有系统、所有账号都填进去。
它们只负责处理那些“默认规则处理不了”的情况。

---

## 1. 这四张表分别管什么？

| 表格 | 一句话解释 | 什么时候改 |
|---|---|---|
| `system_folder_mapping.xlsx` | System 名称和实际 supporting 文件夹名称的对应关系 | 程序找不到文件夹时改 |
| `supporting_sheet_mapping.xlsx` | 检查节点应该去 supporting.xlsx 哪张 sheet 找证据 | 想统一 / 修改 sheet 命名时改 |
| `account_alias_mapping.xlsx` | 少数特殊账号别名映射成标准账号 | 默认标准化后账号还是对不上时改 |
| `manual_override.xlsx` | 人工覆盖程序判断结果 | mentor 说某个系统特殊时改 |

最宝宝版：

```text
system_folder_mapping = 去哪个文件夹找
supporting_sheet_mapping = 去哪张 sheet 找
account_alias_mapping = 这个账号其实是谁
manual_override = 这个系统特殊，听人工的
```

---

# 2. system_folder_mapping.xlsx

## 2.1 它在解决什么问题？

解决：

> Excel / JSON 里的 System 名称，和实际 supporting 文件夹名称不一致。

比如：

| WP / todo_list 里的 TargetSystem | 实际文件夹名称 |
|---|---|
| `PMS/VMS` | `PMS_VMS公共凭证` |
| `Infra Jenkins` | `Infra_Jenkins` |
| `码云` | `MaYun` |

程序后续会先看 `todo_list` 里的 `TargetSystem`，再通过这张表找到真实文件夹。

---

## 2.2 每一列是什么意思？

| 列名 | 是否参与后续代码 | 这列是什么意思 | 维护时怎么改 |
|---|---:|---|---|
| `system` | 是 | WP / todo_list / JSON 里的系统名称，也就是查找 key | 尽量不要乱改，必须和 `System` 或 `TargetSystem` 对得上 |
| `actual_folder_name` | 是 | 电脑里真实的 supporting 文件夹名称 | **最常改这一列**，文件夹实际叫什么就填什么 |
| `alias` | 后续可用 | 系统可能出现的其他叫法 | 可以补充，用分号隔开，例如 `PMS&VMS; PMS_VMS` |
| `note` | 否 | 给人看的备注 | 随便写，说明为什么这样映射 |

---

## 2.3 你平时改哪列？

最常改：

```text
actual_folder_name
alias
note
```

尽量少改：

```text
system
```

因为 `system` 是程序用来查找的 key。key 改错，程序就找不到了。

---

## 2.4 例子

如果 `todo_list` 里是：

```text
TargetSystem = PMS/VMS
```

但是实际文件夹叫：

```text
PMS_VMS公共凭证
```

那就维护成：

| system | actual_folder_name | alias | note |
|---|---|---|---|
| PMS/VMS | PMS_VMS公共凭证 | PMS&VMS; PMS_VMS | 公共平台，Refer to 时使用 |

程序后续逻辑可以理解成：

```python
folder_name = system_folder_map.get(target_system, target_system)
```

意思是：

```text
如果映射表里有，就用 actual_folder_name。
如果映射表里没有，就默认文件夹名 = TargetSystem。
```

---

# 3. supporting_sheet_mapping.xlsx

## 3.1 它在解决什么问题？

解决：

> 某个检查节点，应该去 supporting.xlsx 的哪张 sheet 找证据。

比如：

| 检查节点 | 应该去 supporting.xlsx 哪张 sheet |
|---|---|
| 需审阅应用层管理员账号 | 应用层管理员账号 |
| 需上传开发人员清单 | 开发人员清单 |
| 无前后台管理员SOD问题 | SOD检查 / 或 derived_check |

---

## 3.2 每一列是什么意思？

| 列名 | 是否参与后续代码 | 这列是什么意思 | 维护时怎么改 |
|---|---:|---|---|
| `node` | 是 | 检查节点名称，必须和 `todo_list` 里的 `检查节点` 一致 | **尽量不要改**，除非代码里的节点名也改了 |
| `expected_sheet_name` | 是 | supporting.xlsx 里推荐 / 约定的 sheet 名 | **最常改这一列**，你希望证据放哪张 sheet 就填什么 |
| `evidence_type` | 是 | 证据类型，告诉程序怎么读 | 谨慎改，常见值见下表 |
| `note` | 否 | 给人看的备注 | 随便写，说明这个 sheet 放什么 |

---

## 3.3 evidence_type 怎么理解？

| evidence_type | 意思 | 举例 |
|---|---|---|
| `screenshot` | 主要读截图，需要 OCR | 审批流截图、权限页面截图 |
| `screenshot_or_excel` | 截图或 Excel 导出都可能有 | 应用层管理员账号、服务器账号 |
| `excel_or_txt` | 主要读 Excel / TXT 清单 | 开发人员清单 |
| `derived_check` | 不是直接读一张图，而是用前面抽出来的数据计算 | SOD、开发人员是否有只读以上权限 |

---

## 3.4 你平时改哪列？

最常改：

```text
expected_sheet_name
note
```

谨慎改：

```text
evidence_type
```

尽量不改：

```text
node
```

---

## 3.5 例子

如果程序输出的检查节点是：

```text
需审阅应用层管理员账号
```

但你和 mentor 约定 supporting.xlsx 里这张 sheet 叫：

```text
应用管理员截图
```

那就改成：

| node | expected_sheet_name | evidence_type | note |
|---|---|---|---|
| 需审阅应用层管理员账号 | 应用管理员截图 | screenshot_or_excel | 放应用层管理员账号截图或清单 |

不要改 `node`，只改 `expected_sheet_name`。

---

# 4. account_alias_mapping.xlsx

## 4.1 它在解决什么问题？

解决：

> 少数账号默认标准化后还是对不上。

注意：这张表不是账号主数据表，不要把所有账号都填进去。

大多数账号应该走默认规则：

```text
去空格 → 统一大小写 → 去掉常见分隔符 → 再匹配
```

比如这些不需要维护进 alias 表：

| 原始账号 | 默认标准化后 |
|---|---|
| ` SHMS9037 ` | `SHMS9037` |
| `shms9037` | `SHMS9037` |
| `TCGX179 ` | `TCGX179` |

只有下面这种特殊情况，才需要维护：

| 原始写法 | 默认标准化后 | 实际应该识别成 |
|---|---|---|
| `江丽` | `江丽` | `TMAC` |
| `Jiang Li` | `JIANGLI` | `TMAC` |
| `SH Message 037` | `SHMESSAGE037` | `SHMS9037` |

---

## 4.2 默认规则 + 特殊映射应该怎么走？

后续账号标准化逻辑应该是：

```text
原始账号
  ↓
默认标准化
  ↓
查 account_alias_mapping.xlsx
  ↓
如果有特殊映射：用 standard_account
如果没有特殊映射：用默认标准化结果
```

也就是：

```python
standard = alias_map.get(default_normalized_name, default_normalized_name)
```

---

## 4.3 每一列是什么意思？

| 列名 | 是否参与后续代码 | 这列是什么意思 | 维护时怎么改 |
|---|---:|---|---|
| `raw_name` | 是 | 特殊原始写法，例如中文名、英文名、OCR 错读名 | **只填特殊情况**，不要填所有账号 |
| `standard_account` | 是 | 最终统一成的标准账号 | **最关键**，后续 set 匹配用它 |
| `chinese_name` | 后续可用 | 中文名 | 知道就填，不知道可以空着 |
| `source` | 否 / 可选 | 来源，比如 mentor / OCR / system_export | 建议写，方便追溯 |
| `note` | 否 | 备注 | 随便写 |

如果后续你加 `enabled` 列，也可以作为开关使用：

| 列名 | 是否参与后续代码 | 意思 |
|---|---:|---|
| `enabled` | 可用 | TRUE 表示启用，FALSE 表示暂时不用这条映射 |

---

## 4.4 你平时改哪列？

只在特殊账号出现时，新增或修改：

```text
raw_name
standard_account
chinese_name
source
note
```

最重要的是：

```text
raw_name → standard_account
```

---

## 4.5 例子

| raw_name | standard_account | chinese_name | source | note |
|---|---|---|---|---|
| 江丽 | TMAC | 江丽 | mentor | 中文名对应 AD 账号 |
| Jiang Li | TMAC | 江丽 | mentor | 英文名对应 AD 账号 |
| SH Message 037 | SHMS9037 |  | OCR | OCR 可能误读 |

意思是：

```text
截图 / OCR 里看到 江丽 或 Jiang Li，都统一成 TMAC。
截图 / OCR 里看到 SH Message 037，统一成 SHMS9037。
```

---

# 5. manual_override.xlsx

## 5.1 它在解决什么问题？

解决：

> 程序按照通用规则判断出来的结果，和 mentor / WP 确认的特殊结果不一致。

这种情况不要急着改主代码。

先在 `manual_override.xlsx` 里加一行，让人工结论覆盖程序判断。

---

## 5.2 每一列是什么意思？

| 列名 | 是否参与后续代码 | 这列是什么意思 | 维护时怎么改 |
|---|---:|---|---|
| `system` | 是 | 哪个系统要覆盖 | 必须和 WP / todo_list 里的 System 对得上 |
| `node` | 是 | 覆盖哪个检查节点 | 必须和程序里的检查节点名称一致 |
| `override_label` | 是 | 人工覆盖后的结果 | **最常改**，例如 `是` / `N/A` / `Refer to XXX` |
| `requires_evidence` | 是 | 覆盖后是否仍需要凭证 | TRUE / FALSE，必须认真填 |
| `refer_to` | 是 | 如果覆盖结果要引用公共系统，填这里 | 例如 `PMS/VMS`、`码云`、`Infra Jenkins` |
| `reason` | 建议参与 / 给人看 | 为什么覆盖 | 建议认真写，方便审计留痕 |
| `enabled` | 是 | 这条人工覆盖是否启用 | TRUE 启用，FALSE 暂停 |

---

## 5.3 override_label 应该怎么填？

| override_label | requires_evidence | refer_to | 含义 |
|---|---:|---|---|
| `是` | TRUE | 空 | 本系统需要检查凭证 |
| `N/A` | FALSE | 空 | 不适用，不需要凭证 |
| `Refer to PMS/VMS` | TRUE | PMS/VMS | 去 PMS/VMS 查凭证 |
| `Refer to 码云` | TRUE | 码云 | 去码云查凭证 |
| `N/A-Refer to PMS/VMS` | TRUE | PMS/VMS | 本节点不直接审阅，但仍引用 PMS/VMS 凭证 |
| `待确认` | FALSE | 空 | 暂不自动判断，留给人工确认 |

重点：

```text
Refer to XXX 虽然不是查本系统，但仍然需要凭证，所以 requires_evidence = TRUE。
```

---

## 5.4 你平时改哪列？

新增规则时必须填：

```text
system
node
override_label
requires_evidence
enabled
```

有 Refer 时填：

```text
refer_to
```

建议认真写：

```text
reason
```

---

## 5.5 例子

| system | node | override_label | requires_evidence | refer_to | reason | enabled |
|---|---|---|---|---|---|---|
| JDE | 需审阅发版工具发版账号 | 是 | TRUE |  | mentor/WP 确认的系统特例 | TRUE |
| ABC系统 | 需审阅数据库账号 | N/A | FALSE |  | mentor 确认该系统无数据库账号 | TRUE |
| VMall | 需审阅发版工具后台账号 | Refer to PMS/VMS | TRUE | PMS/VMS | 后台由公共 PMS/VMS 管理 | TRUE |

---

# 6. 维护时按问题找表

| 你遇到的问题 | 改哪张表 | 改哪列 |
|---|---|---|
| 程序找不到系统文件夹 | `system_folder_mapping.xlsx` | `actual_folder_name`，必要时补 `alias` |
| TargetSystem 和文件夹名称不一致 | `system_folder_mapping.xlsx` | `actual_folder_name` |
| 不知道某个检查点去哪张 sheet 找 | `supporting_sheet_mapping.xlsx` | `expected_sheet_name` |
| supporting sheet 改名了 | `supporting_sheet_mapping.xlsx` | `expected_sheet_name` |
| OCR 账号和系统账号默认清洗后仍对不上 | `account_alias_mapping.xlsx` | 新增 `raw_name` → `standard_account` |
| 中文名 / 英文名 / AD 账号需要对应 | `account_alias_mapping.xlsx` | 新增或修改 `raw_name`、`standard_account`、`chinese_name` |
| mentor 说某个系统特殊 | `manual_override.xlsx` | 新增 `system` + `node` + `override_label` |
| 某条人工覆盖先不用 | `manual_override.xlsx` | 把 `enabled` 改成 FALSE |

---

# 7. 推荐维护顺序

每次出问题，按这个顺序想：

```text
1. 是找不到文件夹吗？
   → 改 system_folder_mapping.xlsx

2. 是找不到 sheet 吗？
   → 改 supporting_sheet_mapping.xlsx

3. 是账号名字对不上吗？
   → 先看默认标准化能不能解决；不行再改 account_alias_mapping.xlsx

4. 是某个系统规则特殊吗？
   → 改 manual_override.xlsx

5. 是所有系统通用规则都变了吗？
   → 再考虑改 Python 代码
```

最重要原则：

```text
能改配置表，就先不要改代码。
```

---

# 8. 和后续 workflow 的关系

这四张表不是孤立的，它们会接在主程序输出后面：

```text
WP / Lead Sheet
        ↓
Scope Engine
生成 program_template + todo_list + JSON
        ↓
System-Folder Mapper
读取 system_folder_mapping.xlsx
        ↓
Evidence Loader
读取 supporting_sheet_mapping.xlsx
        ↓
OCR Extractor
抽取账号文字
        ↓
Account Normalizer
默认标准化 + account_alias_mapping.xlsx 特殊映射
        ↓
Evidence Checker
set 匹配 / SOD intersection / 权限等级判断
        ↓
Manual Override
读取 manual_override.xlsx
        ↓
Result Reporter
输出 PASS / FAIL / REVIEW / 缺失凭证
```

---



