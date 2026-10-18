# 求职投递助手 · Rebacca

个人求职工作流：分析当次岗位、推荐本地简历、起草投递材料，经确认后登记本地原始追踪表。仅显式调用，每次处理一个岗位。

## 换电脑使用

1. 在新电脑登录有权限的 GitHub 账号，克隆或下载本仓库。
2. 自行迁移简历文件和追踪表原文件。本仓库不携带这些文件，也不会创建替代追踪表。
3. 在仓库目录用 PowerShell 执行，将下列三个示例路径换成新电脑实际路径：

```powershell
.\install.ps1 -ProjectRoot 'C:\JobSearch\Applications' -ResumeRoot 'C:\JobSearch\Resumes' -TrackerPath 'C:\JobSearch\Applications\暑期实习.xlsx'
```

安装器将技能复制到 `$CODEX_HOME/skills/rebacca`（未设置时为 `~/.codex/skills/rebacca`），保存本机路径配置，并在指定项目的 AGENTS.md 添加显式调用入口。保留现有 AGENTS.md 内容，遇到已有技能会停止而不会覆盖。它不修改或上传简历、追踪表内容。

4. 在 Codex 打开该项目，确认技能列表可找到“求职投递助手”。文件安装成功不等于已验证应用发现；若未出现，重启 Codex 再检查。
5. 输入 `$Rebacca` 加本次 JD，或“调用求职投递助手 skill”。底层名称为小写 `rebacca`，也可通过技能选择器选择 `$rebacca`。

也可让新电脑的 Codex 使用 skill-installer，从本仓库的 `skills/rebacca` 安装。安装后仍需按 `local-settings.example.json` 配置真实路径；若需要大小写口令路由，应在项目 AGENTS.md 添加显式入口。

## 必须保留的规则

- 单独 JD、岗位咨询或写邮件不会自动启动完整流程。
- 登记前必须确认是否已投递及目标 sheet；起草不表示授权发送邮件。
- 文字 JD 全文原样入表；图片 JD 实际嵌入对应 JD 单元格，不能用 OCR 或摘要替代。
- 直接编辑指定原文件，保留现有记录、格式、超链接、WPS 图片和 DISPIMG；锁定或无权限时停止，不另存更新版。
- 当前附带脚本支持文字 JD 追加和只读检查；图片嵌入需由支持该能力的编辑工具执行，未实现时不得声称完成。
- 既有简历之间存在的事实冲突保留待核实，不编造毕业日期、到岗承诺或业绩。

仓库包含个人背景摘要，建议保持私有。`local-settings.json`、源文件、投递日志及账号密钥不提交。

## 运行环境

安装器：Windows PowerShell 5.1 或 PowerShell 7。文字追踪表脚本：Python 3.9+，仅标准库，无需整表转换库。简历读取和图片嵌入由当前 Codex 环境的可用工具完成。
