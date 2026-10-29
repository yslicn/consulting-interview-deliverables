# Consulting Interview Deliverables · 咨询访谈交付 Skill

[English](#english) | [中文说明](#中文说明)

## English

A Claude Skill for consulting teams: create structured Word meeting minutes from original interview transcripts and incrementally maintain one shared Excel issue ledger per project. Current release: **1.2.0**. Project names, phases, workstreams and the number of teams are configurable.

**This is a drafting and review aid. Project staff must review deliverables before formal archiving.** Rules and code checks cannot replace business judgment or guarantee consistent quality across models.

### Outputs and workflow

- One minutes document per interview, organized into themes, subtopics and numbered points. Internal source identifiers stay in the working files rather than the client-facing narrative.
- One issue ledger per project, with five worksheets: issues, verification follow-up, evidence, interview registry, and definitions/configuration. The generated workbook uses Chinese worksheet and column names.
- Separate problems, leads requiring verification, cross-cutting topics, management requests and implementation risks. Configure one lead team and optional collaborating teams per record.
- Append interviews and evidence using stable IDs while preserving existing verification and confirmation information. Changes to existing business fields require an explicit reason.
- Fixed generators produce editable Office files, back up replaced files, and create structural checks and semantic-review materials.

The public package includes neutral styling and entirely fictional examples. **It contains no real client materials, corporate logos, proprietary templates or font files.** Use a private style profile for authorized project branding. The operating instructions and reference documents are currently in Chinese; a bilingual README does not imply that every bundled document or output label is translated.

### Installation

#### Claude Code

Requires Git and Python 3.10 or later:

```bash
git clone https://github.com/yslicn/consulting-interview-deliverables.git ~/.claude/skills/consulting-interview-deliverables
python -m pip install -r ~/.claude/skills/consulting-interview-deliverables/requirements.txt
```

Install dependencies using your team's Python environment practices. If the destination already exists, back it up first; do not overwrite older project resources blindly. Start a new Claude Code session after installation.

#### Custom Skills in Claude web / desktop

Download the ZIP from [Releases](https://github.com/yslicn/consulting-interview-deliverables/releases) and upload it in an environment that supports custom Skills. Code execution and file creation must be enabled; availability depends on account and organization settings. See the [official Claude instructions](https://support.claude.com/en/articles/12512180-use-skills-in-claude).

The ZIP contains one top-level skill directory. If the environment cannot execute Python or install the dependencies, loading the instructions alone will not make the generators operational.

#### Optional tools and fonts

`--recalculate` and `--preview` require LibreOffice (`soffice` or `libreoffice`) on PATH. Without it, you can generate a review draft, but must not claim that formula caches or print layout were verified.

The default font declarations are Noto Sans CJK SC for Chinese and Arial for Latin text. Install suitable fonts locally; substitution may change pagination. See [Style and environment](references/style-and-tools.md). Keep client logos, real data and private profiles out of the public repository.

### Prompt for Claude

```text
Use the consulting-interview-deliverables Skill to prepare this interview's deliverables from the original transcript and any audio you can actually process.
Project name: [name]; project code: [e.g. ABC]; phase: [phase].
Teams and scope: [list each team].
Interviewee, date, location, meeting duration and participants: [details; leave unknown facts unresolved].
Original materials: [paths]; interview outline: [optional, read-only].
Existing project ledger: [full path, or create a new one]; output directory: [full path].
Private style profile: [optional path].
Produce one minutes document for this interview and maintain one shared ledger for the project. Work independently from original materials, without using prior AI deliverables.
Preserve existing confirmation information. Check duplicates and team assignments before appending evidence.
Distinguish current-state statements, requests, suggestions and questions. Do not turn aspirations into proven deficiencies or combine separate remarks into an assigned owner or deadline.
Review semantic support, check the files, and perform available recalculation/rendering checks. Disclose the checks that could not be completed.
```

For audio-only inputs, first establish whether the environment can process the recording. Otherwise obtain a readable transcript; do not claim to have listened to audio or produced a verbatim transcript when that did not happen.

### Local quick check

Run from the cloned repository. The example is a fictional procurement interview and refers to no real organization or person.

```bash
python -m unittest discover -s tests -v
python scripts/build_deliverables.py \
  --data examples/interview.json --sources-dir examples \
  --minutes-out outputs/demo_minutes.docx --ledger-out outputs/demo_ledger.xlsx \
  --work-dir .interview-work/demo --recalculate --preview
python scripts/check_deliverables.py \
  --minutes outputs/demo_minutes.docx --ledger outputs/demo_ledger.xlsx \
  --expected-data .interview-work/demo/normalized-data.json --require-calculated
```

Without LibreOffice, omit `--recalculate`, `--preview` and the checker's `--require-calculated`; retain the resulting verification warnings. Existing minutes are protected from overwriting by default: use a new versioned filename, or explicitly pass `--replace-minutes`. Writing to an existing ledger path reads and merges the project ledger first. See the [Data contract](references/data-contract.md).

Use `semantic-review.json` in the working directory to review assertions, speaker attribution and action commitments against the source. `generation-report.json` records changes and the scope of checks. Do not automatically include these internal working files with client-facing minutes.

### Limitations and conditions of use

1. **Business facts are not automatically verified.** Source statements may themselves be inaccurate. Quote matching checks that text exists; it does not verify speaker attribution, location, context or whether a conclusion is justified. An evidence-type label is not factual confirmation.
2. **No built-in audio recognition guarantee.** This package includes no transcription service, speaker identification or audio-listening capability. The source checker accepts UTF-8 text only. Extract other formats with available tools and review critical content manually.
3. **Human acceptance is required.** Review deficiency claims, negative statements, responsibilities, dates, quantities, product stages and management requests. Passing structural checks does not authorize unattended final archiving. The semantic-review flag always remains required.
4. **Existing ledgers must match the supported schema.** The generator reads and writes this version's five-sheet structure. Macro-enabled workbooks are unsupported, and custom formatting is not guaranteed to survive. Back up and explicitly map local v1.1 files, extra worksheets/columns and other team formats before migration; do not overwrite them directly. See [Maintenance and migration](references/maintenance.md).
5. **Use a single writer, sequentially.** File fingerprints and backups reduce overwrite risk but provide no collaboration lock. The two output files are not a cross-file transaction. Do not let multiple users/processes write the same ledger concurrently or retry an interrupted replacement without inspection.
6. **Results depend on the environment.** Python dependencies, installed fonts and Office tools affect output. LibreOffice recalculation and preview are not validation in a particular Microsoft Word/Excel version. Inspect pagination and long text before print archiving. Cells estimated to exceed readable height are rejected; split the evidence where appropriate.
7. **Users are responsible for permissions.** Handle recordings, personal information, client materials and external-model uploads under project agreements and organizational policies. The MIT license covers this repository's code and original documentation, not third-party brands or materials.
8. **No cross-model quality or runtime guarantee.** Rules, fixed generators and fictional tests are not comprehensive certification across models, industries or long interviews. Before team rollout, independently trial representative original materials and review the results.

### Maintenance and validation scope

Start with [SKILL.md](SKILL.md). The 24 automated tests cover the data contract, quote checks, evidence-type gates, incremental updates, preservation of confirmation data, explicit change auditing, file generation and structural checks. GitHub Actions runs them on Python 3.10 and 3.12; it performs no real-client semantic assessment and installs no Office layout engine.

Use incremental checks for routine additions. After script, formula or schema changes, rerun full tests and actual export checks. Report issues using sanitized or fictional minimal examples; never upload client information to public issues.

Release status: suitable for team trials and human-reviewed delivery. See [CHANGELOG.md](CHANGELOG.md).

### License

[MIT](LICENSE).

---

## 中文说明

面向咨询项目组的 Claude Skill：基于原始转写材料，生成结构化 Word 会议纪要，并增量维护整个项目共用的 Excel 问题台账。当前版本 **1.2.0**，支持不同客户、阶段及任意数量的小组。

**定位：辅助制作和评审的工具，正式归档必须经过项目人员复核。** 规则和代码检查不能替代业务判断，也不能保证不同模型输出相同质量。

### 输出与工作方式

- 每场访谈一份会议纪要，按主题、子议题和编号要点组织；内部来源编号留在工作区，不进入客户正文。
- 每个项目一个问题台账，包含问题台账、核实跟进、证据明细、访谈登记、口径与配置五个工作表。
- 问题、待核实线索、综合议题、管理诉求、实施风险分开记录；归口小组与协同小组可配置。
- 按稳定编号追加访谈和证据，保留既有核实及确认信息；修改已有业务字段需明确提交修改理由。
- 固定脚本生成可编辑 Office 文件；自动备份被替换文件，形成结构检查报告和语义复核材料。

公开包使用中性样式和完全虚构的示例，**不包含真实客户材料、企业 Logo、专有模板或字体文件**。可通过私有样式配置复用本项目的品牌资源。Skill 操作规范及参考文档目前以中文提供，生成的 Excel 表名和字段也为中文；双语 README 不代表包内文档及输出界面均已翻译。

### 安装

#### Claude Code

需要 Git、Python 3.10 或以上版本。在终端执行：

```bash
git clone https://github.com/yslicn/consulting-interview-deliverables.git ~/.claude/skills/consulting-interview-deliverables
python -m pip install -r ~/.claude/skills/consulting-interview-deliverables/requirements.txt
```

建议按所在团队的 Python 环境规范安装依赖。若目录已存在，先备份，勿直接覆盖旧版项目资源。重启或新开 Claude Code 会话后使用该 Skill。

#### Claude 网页端 / 桌面端的自定义 Skill

从 [Releases](https://github.com/yslicn/consulting-interview-deliverables/releases) 下载 ZIP，在支持自定义 Skill 的环境中上传。需要允许文件生成和代码执行；功能可用性受账号和组织设置影响。上传方式见 [Claude 官方说明](https://support.claude.com/en/articles/12512180-use-skills-in-claude)。

ZIP 包含一个顶层技能目录。若当前运行环境不允许安装依赖或执行 Python，单纯加载说明不会使生成脚本可运行，应先解决环境条件。

#### 可选依赖与字体

`--recalculate` 和 `--preview` 需要 PATH 中存在 LibreOffice（`soffice` 或 `libreoffice`）。未安装时可生成待评审稿，但不能声称已验证公式缓存或打印版式。

默认声明中文字体 Noto Sans CJK SC、英文字体 Arial；机器需自行安装对应字体，否则可能发生替代和分页变化。品牌配置见 [样式与环境](references/style-and-tools.md)。不要将客户 Logo、真实数据或私有配置提交到公开仓库。

### 给 Claude 的使用提示词

```text
使用 consulting-interview-deliverables Skill，依据我提供的原始转写稿及可用录音完成本次访谈交付。
项目名称：[项目名称]；项目代号：[如 ABC]；阶段：[阶段]。
项目小组及范围：[逐组列出]。
访谈对象、时间、地点、会议时长、参会人：[填写；未知的保持待明确]。
材料位置：[原始材料]；访谈提纲：[可选，只读]。
现有项目台账：[完整路径；没有则新建]；输出目录：[完整路径]。
私有样式：[可选配置路径]。
每场一份纪要，整个项目共用一份台账。独立依据原始材料，不使用历史AI成品。
保留已有台账的确认信息，先核对去重与归口，再追加本次证据。
区分现状、诉求、建议和提问，不把期望写成缺陷，不拼接责任或期限。
生成后完成语义复核、文件检查和可用的重算/渲染检查；说明未完成的验证范围。
```

只有录音时，先确认当前环境能否处理录音。不能处理时需提供可读转写稿，不得假称复听或生成逐字稿。

### 本地快速验证

以下路径均相对于克隆目录。示例是虚构的采购访谈，不涉及真实组织或个人。

```bash
python -m unittest discover -s tests -v
python scripts/build_deliverables.py \
  --data examples/interview.json --sources-dir examples \
  --minutes-out outputs/demo_minutes.docx --ledger-out outputs/demo_ledger.xlsx \
  --work-dir .interview-work/demo --recalculate --preview
python scripts/check_deliverables.py \
  --minutes outputs/demo_minutes.docx --ledger outputs/demo_ledger.xlsx \
  --expected-data .interview-work/demo/normalized-data.json --require-calculated
```

没有 LibreOffice 时省略两个可选参数，并省略检查命令的 `--require-calculated`；此时报告中的未验证警告应保留。已有纪要默认拒绝覆盖，修订应换版本名，或明确使用 `--replace-minutes`。再次写入同一台账路径会先读取并增量合并，详见 [数据合同](references/data-contract.md)。

工作区的 `semantic-review.json` 用于逐条人工核对断言、发言归属和行动安排；`generation-report.json` 记录检查范围及变更。工作区材料不要作为客户纪要附件自动交付。

### 使用限制与适用边界

1. **不能自动确认业务事实。** 原文中的说法仍可能不准确。引文匹配仅检查文本是否存在，不验证说话人、定位、上下文含义或推论是否成立；填写证据性质标签也不意味着事实已核实。
2. **不能保证录音识别质量。** 本包不内置语音识别服务、说话人识别或音频复听功能。原文校验脚本只接受 UTF-8 文本文件；其他格式先用当前环境支持的工具提取，并人工核对关键内容。
3. **必须人工验收。** 尤其复核缺陷判断、否定表述、职责、时间、数量、产品阶段和管理层诉求。结构检查通过不代表成品可直接免审归档。语义复核标记始终保持需要复核。
4. **已有台账有结构要求。** 脚本只读写本版约定的五表结构；不支持带宏文件，也不保证保留自定义格式。旧版 v1.1、额外工作表/列或其他团队格式需先备份并显式映射迁移，不能直接覆盖。见 [维护与迁移](references/maintenance.md)。
5. **单人串行维护。** 文件指纹检查和备份降低覆盖风险，但不提供多人协同锁；两个输出文件也不是跨文件事务。不要让多人或多个进程同时写同一台账，勿在异常后不经检查重复覆盖。
6. **环境影响结果。** 需要 Python 依赖、可用字体及相应 Office 工具。LibreOffice 重算与预览不等于已在目标版本的 Microsoft Word/Excel 验证。打印归档前检查实际分页和长文本；过长单元格会拒绝生成，需拆分证据。
7. **权限由使用者负责。** 处理录音、个人信息、客户资料及上传外部模型须遵守项目约定和组织政策。MIT 许可证只覆盖本仓库代码与原创说明，不授予任何第三方品牌或资料的使用权。
8. **没有跨模型质量或耗时保证。** 当前包含规则、固定生成器及虚构数据测试，不是对所有模型、行业和长访谈的全面效果认证。首次在团队环境部署，应选取有代表性的原始材料，独立试运行后验收。

### 维护与测试范围

完整规则入口：[SKILL.md](SKILL.md)。24 项自动化测试覆盖数据合同、引文检查、事实类型门槛、增量维护、保留确认信息、显式修改审计、文件生成与结构检查等。GitHub Actions 在 Python 3.10 / 3.12 上运行这些测试；不执行真实客户语义评审，也不安装 Office 排版引擎。

日常追加采用增量检查；脚本、公式或结构变化后重新执行完整测试与实际导出检查。团队反馈问题时请使用脱敏或虚构的最小复现材料，不在公开 Issue 上传客户信息。

本版本的发布状态是“可供团队试用并经人工复核交付”。升级说明见 [CHANGELOG.md](CHANGELOG.md)。

### 许可证

[MIT](LICENSE)。
