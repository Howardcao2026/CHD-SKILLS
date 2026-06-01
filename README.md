# CHD-SKILLS

Howard 的个人 AI 技能仓库，用于 WorkBuddy / Claude Code 等 Agent 环境的技能管理。

## 技能列表

| 技能名 | 版本 | 说明 | 触发词 |
|--------|------|------|--------|
| [shandong-meeting-minutes](./shandong-meeting-minutes/) | v1.2.0 | 山东公司生产月度会议纪要生成器 | 会议纪要、生产会纪要、生成纪要 |
| [ai-pastor](./ai-pastor/) | - | AI 基督教牧养助手（讲道稿、圣经讲解、灵修） | 讲道稿、灵修、圣经讲解、ai-pastor |

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

### ai-pastor

**功能**：AI 基督教牧养助手，用于讲道稿、圣经讲解、灵修、家庭敬拜等场景。

**特点**：
- 基于 2025 年讲道 60+ 场讲章经验
- 支持大纲生成、逐字稿扩展、经文深度解读
- 内置讲道风格指南（金句开头、生活化例子、呼召结尾）

## 自动同步脚本

在 PowerShell 中添加以下函数，实现一键推送更新：

```powershell
function skill-push {
    param([string]$Message = "Update skills")
    $repo = "E:\WORKBUDDY2026.5\2026-06-01-15-03-50\CHD-SKILLS-deploy"
    $skills = @(
        "$env:USERPROFILE\.workbuddy\skills\shandong-meeting-minutes",
        "$env:USERPROFILE\.workbuddy\skills\ai-pastor"
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
