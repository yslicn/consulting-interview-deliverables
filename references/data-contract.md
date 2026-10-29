# 工作数据合同

完整可运行例子见`examples/interview.json`，原始虚构文本为`examples/transcript.txt`。字段名采用JSON中的英文键；输出文件使用中文列名。工作JSON和审查包不默认交付客户。

## 顶层

| 字段 | 用途 |
|---|---|
| project | 必填：name、ASCII code、groups数组（每项name和可选scope）；组数可变 |
| interview | 本次访谈；生成纪要时必填；自动纳入项目访谈登记 |
| interviews | 可选：其他需追加的访谈登记；同一ID元数据不能冲突 |
| minutes | 本次纪要内容；不写入项目台账存储 |
| records | 新记录，或与已有数据相同、仅扩充来源/关联的记录 |
| evidence | 本次新增证据；允许无关联记录的背景或行动证据 |
| record_updates | 显式修订既有记录，必须提供id、reason、changes |

`interview`可包括id、stage（可空）、level、object、date（ISO日期）、place、duration_hours（数字）、participants数组、topic、has_outline布尔值、source_file、minutes_file。会议时间或地点缺失可为空，但正式归档前人工补核。

## 记录

每条`records`需要id、group、nature、domain、title、description、interview_ids，以及verification中的focus、materials、completion_criteria。

可选字段：collaborating_groups、impact、management_request、departments、status、related_ids、historical_id。性质和状态取值见台账规范。无依据的字段用null，不用0或虚拟姓名填满。

`verification`其他字段：owner、due_date、cause_hypothesis、priority、conclusion、confirmer、confirmed_date。日期为`YYYY-MM-DD`；优先级为高、中、低或待评估。确认及处置状态要求有结论、确认人和日期。

## 证据

每条需要id、record_ids数组、interview_id、speaker_role（client/consultant/unknown）、speaker、source_file、locator、quote、utterance_type、purpose。可选outline_question_id、note及source_sha256。

`utterance_type`取值：现状陈述、管理诉求、顾问建议、提问、明确决定、明确安排、其他。提纲缺省时题号为“不适用”。同一轮发言可分段标记，不能把整段含希望、疑问的发言全部当作现状。

`source_file`是相对于`--sources-dir`的UTF-8文本路径，不允许越出该目录。Word/PDF转写材料应先用可用工具提取为可核对的UTF-8文本副本，保留原始文件与对应关系。音频不是该参数的文本来源。

引文匹配忽略排版空白，其他文字必须出现在给定原文中。生成器计算来源摘要，保存在台账与工作报告中。该检查不验证定位、发言归属或业务推断。

## 纪要

`minutes`包括title、overview、sections、可选actions及open_items。

每个section有title、items及可选subsections。item有label、text、kind（fact/request/suggestion/decision）和evidence_ids数组。来源引用只用于工作校核，**不会写入客户版Word**。

action包括item、content、owner、when、agreed布尔值及evidence_ids。未明确约定的行动不能包含指派人或期限；agreed为true时须有“明确决定”或“明确安排”类证据。脚本不能判断这些标签是否真实，仍需复核原文。

## 增量输入与显式修订

新增访谈可只给project、新interview及新evidence，已有问题只通过新证据的record_ids关联，无需重复提供全部旧记录。需要新增问题时提供新项目级ID。

修改既有记录示例：

```json
{
  "record_updates": [
    {
      "id": "DEMO-ISS-0001",
      "reason": "项目组核对后调整归口，保留原确认信息",
      "changes": {"group": "流程组", "collaborating_groups": ["采购组"]}
    }
  ]
}
```

合并变更可同时填写status、related_ids和verification中的结论及确认信息。不要又在records里发送同ID的旧字段，否则生成器会识别冲突。

已有来源引文和访谈元数据冲突时拒绝静默覆盖。原始转写纠错、证据重定位或访谈身份变更应保留历史并人工迁移，不通过改标签或删除校验绕过。
