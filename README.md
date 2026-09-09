# client-fund-health-report · 客户基金资产配置健诊报告

根据客户基金持仓、可选产品池与调仓约束，研究最新市场并生成固定版式的中文 Word（DOCX）客户基金资产配置健诊报告。只形成分析与建议，不执行交易，不替代适当性评估。

- 版本：1.0.0
- 运行要求：Python 3.10+，依赖 `lxml`（`pip install -r requirements.txt`）
- 入口：[`SKILL.md`](SKILL.md)

## 一键安装（通用 Skill 目录）

在 Agent 对应的用户技能目录下执行：

```bash
git clone https://github.com/<你的GitHub用户名>/client-fund-health-report.git
```

安装后**新建一个会话**以重新扫描技能。升级：进入该目录执行 `git pull`。

> 不同平台的技能目录位置不同（例如 Codex 为 `~/.codex/skills/`），以所在平台说明为准；安装后应能直接看到 `client-fund-health-report/SKILL.md`，注意不要多套一层文件夹。

## 其他平台安装包

[`dist/`](dist/) 目录内置三个平台的原始安装 ZIP、《安装与使用说明》、SHA256 校验值与快速使用模板：

- `..._Codex_v1.0.0.zip`：Codex
- `..._WorkBuddy_v1.0.0.zip`：WorkBuddy（界面上传）
- `..._豆包扣子_v1.0.0.zip`：扣子编程导入

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
└── dist/                    # 各平台安装包与说明
```

## 使用方式

在新会话中提供：客户持仓（截图/表格/PDF/文字）、可选产品池与销售档期、本轮调仓约束，然后参考 `dist/快速使用模板.txt` 下达指令，例如：

> 使用 client-fund-health-report，根据我上传的客户持仓、可选产品和调整要求生成 Word 报告。

## 隐私与免责

- 模板与示例均已脱敏，产品名称、代码、金额为虚构演示数据，不得直接作为投资建议。
- 客户姓名、账号、持仓金额与截图只能在本地处理，不得上传到公开检索服务。
- 市场数据与产品销售状态需在生成时核验，无法核验时输出标注缺口的草稿。
