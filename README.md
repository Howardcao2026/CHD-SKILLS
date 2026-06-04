# CHD-SKILLS

Howard 的个人 AI 技能仓库，用于 WorkBuddy / Claude Code 等 Agent 环境的技能管理。

## 技能列表

| 技能名 | 版本 | 说明 | 触发词 |
|--------|------|------|--------|
| [shandong-meeting-minutes](./shandong-meeting-minutes/) | v1.2.0 | 山东公司生产月度会议纪要生成器 | 会议纪要、生产会纪要、生成纪要 |
| [monthly-power-report](./monthly-power-report/) | v2.0 | 山东公司月度电量生产分析报告生成器 | 月报分析、电量报告、生成月报、限电率分析 |
| [ai-pastor](./ai-pastor/) | - | AI 基督教牧养助手（讲道稿、圣经讲解、灵修） | 讲道稿、灵修、圣经讲解、ai-pastor |
| [cgnew-energy-news](./cgnew-energy-news/) | v1.0 | 中广核新能源公众号新闻稿生成器（4型模板+联网数据补充） | 写新闻稿、会见新闻稿、拜会新闻稿 |

## 快速开始

### 安装单个技能

将技能目录复制到 WorkBuddy 技能目录：

```bash
# Windows
xcopy /E /I "shandong-meeting-minutes" "%USERPROFILE%\.workbuddy\skills\shandong-meeting-minutes"

# 或手动复制到 ~/.workbuddy/skills/{skill-name}/
```

### 使用方法

安装后，在 WorkBuddy 对话中直接说触发词即可：

```
用户：帮我把这个钉钉会议文字稿生成会议纪要
[附件：会议逐字稿.docx]
→ Agent 自动调用 shandong-meeting-minutes 技能
```

## 技能详情

### shandong-meeting-minutes

**功能**：从钉钉导出的会议逐字稿自动生成符合山东公司公文规范的 Word (.docx) 会议纪要。

**特点**：
- 两段式结构（工作完成情况 + 后续安排）
- 40+ 条术语自动修正表（场站名、项目名、专业术语）
- 严格遵循《山东公司日常文稿规范2026》排版
- Darwin 评分：85.9/100

**文件结构**：
```
shandong-meeting-minutes/
├── SKILL.md              # 完整规范和业务知识
├── scripts/
│   └── generate.py       # .docx 生成器（可独立运行）
└── test-prompts.json     # 测试场景
```

### monthly-power-report

**功能**：从3个Excel文件一键生成含8章节10个ECharts图表的完整HTML月度电量分析报告。

**特点**：
- 自动化流水线：Excel → JSON → HTML（6步流程，含CHECKPOINT确认）
- 8大章节：上网电量、增量vs存量、超欠发明细、限电率、关键场站、风光资源、场用电率、发电小时数
- Apple 简约风格 HTML 输出，自包含单文件
- 3个显式 🔴CHECKPOINT 防止月度适配错误
- 8条反例黑名单 + 穷举式改动清单（Step 2 + Step 4）
- Darwin 评分：81.7/100（优化前 67.1）

**文件结构**：
```
monthly-power-report/
├── SKILL.md              # 完整规范和6步流程
├── scripts/
│   ├── generate_data.py  # 数据提取与计算
│   └── gen_report.py     # HTML 报告生成
├── references/
│   └── report_spec.md    # 报告结构规格
├── evals/
│   └── evals.json        # 评估用例
└── test-prompts.json     # 测试场景
```

**使用方式**：
```
用户：/monthly-power-report 给我报告
[附件：3个Excel文件]
→ Agent 自动运行完整6步流水线
```

### ai-pastor

**功能**：AI 基督教牧养助手，用于讲道稿、圣经讲解、灵修、家庭敬拜等场景。

**特点**：
- 基于 2025 年讲道 60+ 场讲章经验
- 支持大纲生成、逐字稿扩展、经文深度解读
- 内置讲道风格指南（金句开头、生活化例子、呼召结尾）

### cgnew-energy-news

**功能**：根据会议/会见原始文字记录或要点，按企业公众号标准文风自动生成新闻稿。

**特点**：
- 4 种稿件类型全覆盖（集团拜会省领导 / 集团会见伙伴 / 公司拜会市县 / 公司会见伙伴）
- 内置完整高层次词汇库、格式规范和 4 套填空模板
- ⭐ 自动搜索互联网最新数据补充稿件说服力（装机量、政策目标等）
- 所有模板均已脱敏，可直接部署到公开仓库

**文件结构**：
```
cgnew-energy-news/
├── SKILL.md              # 完整工作流程和数据补充规则
└── references/
    └── template.md        # V2.0 表达范式模板（4型+词库+填空模板）
```

**使用方式**：
```
用户：公司董事长今天去拜访了某市市长，谈了十五五新能源布局和储能建设，写一篇新闻稿
→ Agent 自动判断类型 → 搜索互联网补充数据 → 按模板生成稿件
```

## 自动同步脚本

在 PowerShell 中添加以下函数，实现一键推送更新：

```powershell
function skill-push {
    param([string]$Message = "Update skills")
    $repo = "E:\WORKBUDDY2026.5\2026-06-01-15-03-50\CHD-SKILLS-deploy"
    $skills = @(
        "$env:USERPROFILE\.workbuddy\skills\shandong-meeting-minutes",
        "$env:USERPROFILE\.workbuddy\skills\monthly-power-report",
        "$env:USERPROFILE\.workbuddy\skills\ai-pastor",
        "$env:USERPROFILE\.workbuddy\skills\cgnew-energy-news"
    )

    foreach ($src in $skills) {
        $name = Split-Path $src -Leaf
        Copy-Item -Recurse -Force "$src\*" "$repo\$name\"
    }

    Set-Location $repo
    git add . && git commit -m $Message && git push origin main
}
```

## 贡献

个人技能仓库，欢迎参考学习。

## 许可证

MIT License - 自由使用，自负责任。
