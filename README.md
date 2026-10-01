# client-fund-health-report · 客户基金资产配置健诊报告

根据客户基金持仓、可选产品池与调仓约束，研究最新市场并生成固定版式的中文 Word（DOCX）客户基金资产配置健诊报告。报告供银行代销渠道理财经理内部参考，只形成分析与建议，不执行交易，不替代适当性评估。

- 版本：1.1.5
- 运行要求：Python 3.10+，依赖 `lxml`（`pip install -r requirements.txt`）
- 入口：[`SKILL.md`](SKILL.md)

## v1.1.5 主要能力

- **大类资产穿透**：逐只基金取报告日之前最新已披露季报的股票、债券、现金及其他占净值比例，按客户当前持仓市值加权，穿透出组合真实的权益、债券、现金及其他/未识别结构。
- **权益行业归因**：基于各基金前十大重仓股统一行业分类、按市值穿透汇总，同时给出"占总资产""占已识别前十大敞口"两个口径，并标注覆盖度，识别买重复、风格过度集中的隐性问题。
- **固定期间规则**：第一部分复盘报告日所在月之前的最近完整自然月，第二部分展望当月；指数默认展示最近完整月及最近两完整月累计，随报告日滚动。
- **规范版式**：深蓝表头白字、列居中、表内字体统一，红涨绿跌、保留正负号，输出可编辑 DOCX。

## 一键安装

**豆包类 / Codex / WorkBuddy / 扣子 / OpenClaw 各平台的一键安装口令见 [INSTALL.md](INSTALL.md)。**

通用 Skill 目录下执行：

```bash
git clone https://github.com/liuyx120/client-fund-health-report.git
```

安装后**新建一个会话**以重新扫描技能。升级：进入该目录执行 `git pull`。

> 不同平台的技能目录位置不同（例如 Codex 为 `~/.codex/skills/`），以所在平台说明为准；安装后应能直接看到 `client-fund-health-report/SKILL.md`，注意不要多套一层文件夹。

## 安装包

[`dist/`](dist/) 目录提供单一通用安装 ZIP、《安装与使用说明》、SHA256 校验值与快速使用模板：

- `客户基金资产配置健诊报告Skill_v1.1.5.zip`：通用安装包，适用于豆包等通用 AI 对话、扣子、Codex、WorkBuddy、OpenClaw 等平台。

## 目录结构

```
client-fund-health-report/
├── SKILL.md                 # 技能入口（Agent 首先读取）
├── requirements.txt         # Python 依赖
├── references/              # 报告版式、研究校验、输入规范
├── examples/                # 结构化输入示例（虚构演示数据）
├── scripts/build_report.py  # Word 生成程序
├── assets/、templates/      # 脱敏参考模板与预览图
├── agents/、artifact-template.json  # 跨平台兼容文件
└── dist/                    # 通用安装包与说明
```

## 使用方式

在新会话中提供：客户持仓（截图/表格/PDF/文字）、可选产品池与销售档期、本轮调仓约束，然后参考 `dist/快速使用模板.txt` 下达指令，例如：

> 使用 client-fund-health-report，根据我上传的客户持仓、可选产品和调整要求生成 Word 报告。

## 隐私与免责

- 模板与示例均已脱敏，产品名称、代码、金额为虚构演示数据，不得直接作为投资建议。
- 客户姓名、账号、持仓金额与截图只能在本地处理，不得上传到公开检索服务。
- 市场数据与产品销售状态需在生成时核验，无法核验时输出标注缺口的草稿。
