# 各平台一键安装口令

仓库地址：<https://github.com/liuyx120/client-fund-health-report>

## 豆包类 / 通用 AI 对话（直接发送给 AI）

```
按照 https://github.com/liuyx120/client-fund-health-report 安装 client-fund-health-report 技能：完整克隆整个仓库到我的用户技能目录（保留 scripts、references、assets、examples、templates 全部文件），再执行 pip install -r requirements.txt。
```

## Codex（终端执行）

```bash
mkdir -p ~/.codex/skills && cd ~/.codex/skills && git clone https://github.com/liuyx120/client-fund-health-report.git && cd client-fund-health-report && pip install -r requirements.txt
```

## WorkBuddy（对话框直接发送）

```
帮我从 Git 仓库安装技能：https://github.com/liuyx120/client-fund-health-report ，完整保留目录结构，安装到用户级技能目录，并安装 requirements.txt 中的依赖。
```

## 扣子 / 扣子编程（在"上传技能包"处导入此 zip）

```
https://github.com/liuyx120/client-fund-health-report/raw/main/dist/客户基金资产配置健诊报告Skill_豆包扣子_v1.0.0.zip
```

## OpenClaw（终端执行）

```bash
openclaw skill install --github https://github.com/liuyx120/client-fund-health-report
```

---

装完**新开会话**生效；上传客户持仓和可选产品池后，说"**使用 client-fund-health-report 生成 Word 报告**"即可触发。
