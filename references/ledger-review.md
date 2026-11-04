# 台账独立评审（ledger-quality-v1）

评审者须未参与起草，阅读完整台账、原始来源及本次纪要。先独立判断记录是否能用于项目诊断与后续核实，再核查证据；不能只审生成器的摘要或检查关键词。缺少独立评审者时交付待评审稿，不自称独立通过。

## 逐条验收

1. **whole_record**：业务对象和具体差距清楚，标题与描述、性质、影响、诉求、状态、结论一致；完整连贯表达，影响有证据或明确为潜在风险；没有用“待核实”掩盖确定性断言。
2. **source_and_time**：引用回查原文，区分客户陈述、顾问问题/建议、历史事件、当前要求和执行效果。数量、单位、范围、频次及主体无推定补全。证据备注也受同样约束。
3. **verification**：原因假设有依据或留空；核实问题具体，材料与完成判据足以作出确认、排除或继续核实判断；责任人与日期未编造，历史确认信息未被覆盖。
4. **relations_and_counting**：与全表可能重复项比较对象、场景、差距和证据，记录合并或保留理由；综合议题列明子项并核查实际统计排除，合并关系可追溯。

首次采用本规则全表评审；日常增量对新增、修改、关联受影响记录重新评审。未变记录可以沿用已有评审，但四项basis均须引用已通过评审文件及相应记录，并核实记录、证据和关联未改变。不得仅写“同上”“无问题”。空台账允许空record_checks，仍须说明联审范围。

## 纪要与台账联审

内部逐项列出共同事实：业务主体、时点、现状/要求/假设、数量、单位、范围和重要限制，记录Word位置及台账编号。两份文件均与原始证据核对，不以彼此一致代替证据正确。正文、影响、原因、证据备注、核实结论均纳入检查。无共同事实也须记录检查范围及判断依据，不能跳过联审。

以下为虚构反例，不得作为项目事实：

| 不通过的写法 | 返修要求 |
|---|---|
| 原文已说明包装由销售管理，仅顾问建议补充成本测算，标题却写“包装职责归口不清” | 保留已知分工，成本测算的职责与流程列为核实线索，不扩大为组织缺陷 |
| 原文说过去规格漏传、现在已要求销售通知，标题写“规格未及时传递” | 区分历史案例与当前要求，核实执行效果后再判断现状 |
| 纪要写“300吨”，台账写“300、单位待确认”，证据备注又称“按上下文是吨” | 回查原始来源；无法确定则各处统一保留不确定性，不补单位 |
| 两段分别说“一款”和“三款”，汇总成“一至三款” | 先核对时间与口径，未消除矛盾前不得构造新区间 |
| 综合议题仅备注“不重复计数”，未列子项且统计仍计入具体问题 | 明确性质、列全相关子项并核对实际计算结果 |
| 顾问问是否有责任空档，对方没有回答，台账写“职责存在空档” | 问句不是事实，改为有明确核实问题和判据的线索 |

## 评审记录与定稿门

评审JSON必含：

- rules_version：ledger-quality-v1；reviewer、reviewed_at、independent:true。
- data_sha256、ledger_sha256；联审时还含minutes_sha256。使用脚本digest或SHA256工具计算当前文件的真实哈希。
- record_checks：覆盖工作数据中每个record.id，不可重复。每项含id、verdict，以及whole_record、source_and_time、verification、relations_and_counting四个对象，各含assessment（具体判断）、basis（字段、来源位置及理由）。
- cross_document：联审时含verdict及checks，每项含claim、minutes_location、record_ids、assessment、basis；没有纪要时verdict为NOT_APPLICABLE并有reason说明实际范围，不宣称完成联审。
- verdict：PASS或REVISE；issues数组逐条记录问题、位置和修订要求，闭合时status为resolved/not_applicable并填写resolution。空数组表示没有发现问题。

任何不支持的断言、时间状态混淆、跨文件矛盾、未闭合问题均为REVISE。修改后重新生成文件并评审受影响内容、重算哈希；不手改哈希冒充复核。原始证据必须实际读过。程序只核查评审覆盖、记录完整和文件绑定，不能证明评审判断正确或评审者真正独立。

标准格式最终检查：

```bash
python3 <SKILL_DIR>/scripts/check_deliverables.py --ledger <台账.xlsx> \
  --expected-data <normalized-data.json> --ledger-review <台账评审.json> --require-calculated
```

同时定稿纪要时补充 `--minutes <纪要.docx> --final-review <纪要评审.json>`，台账评审同时绑定该Word。
旧格式仅做内容评审时，先按maintenance.md建立完整逐条JSON快照（顶层records，各条保留id及所有原字段），单独执行：

```bash
python3 <SKILL_DIR>/scripts/check_ledger_review.py --review <台账评审.json> \
  --data <旧表完整快照.json> --ledger <原台账.xlsx> --minutes <联审纪要.docx>
```

此命令不验证快照与Excel一致性，也不替代旧表映射逐行复核或标准格式校验；不得宣称格式验收通过。
