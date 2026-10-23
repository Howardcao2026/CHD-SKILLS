# CHD-SKILLS

Howard 的个人 AI 技能仓库，用于 WorkBuddy / Claude Code 等 Agent 环境的技能管理。

## 技能列表

| 技能名 | 版本 | 说明 | 触发词 |
|--------|------|------|--------|
| [shandong-meeting-minutes](./shandong-meeting-minutes/) | v1.2.0 | 山东公司生产月度会议纪要生成器 | 会议纪要、生产会纪要、生成纪要 |
| [monthly-power-report](./monthly-power-report/) | v2.0 | 山东公司月度电量生产分析报告生成器 | 月报分析、电量报告、生成月报、限电率分析 |
| [ai-pastor](./ai-pastor/) | - | AI 基督教牧养助手（讲道稿、圣经讲解、灵修） | 讲道稿、灵修、圣经讲解、ai-pastor |
| [CCTV-NEWS](./CCTV-NEWS/) | v1.0 | 新闻稿生成器（4型模板+联网数据补充） | 写新闻稿、会见新闻稿、拜会新闻稿 |
| [BIZ-PPT-SKILL](./BIZ-PPT-SKILL/) | v1.0 | 企业级汇报 PPT 统一样式库（4套样式+母版安全区+矢量图表+Word逐字稿） | 做一版季度汇报、经营分析会 PPT、出财报风格、出第一性原理风格 |

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

### CCTV-NEWS

**功能**：根据会议/会见原始文字记录或要点，按企业公众号标准文风自动生成新闻稿。

**特点**：
- **8 种稿件类型全覆盖** ⭐：A 集团拜会省领导 / B 集团会见伙伴 / C 公司拜会市县 / C+ 公司拜会省厅 / D 公司会见伙伴 / E 内部研讨部署 / F 接待调研 / G 内部技术交流 / H 获奖喜讯
- 内置完整高层次词汇库、格式规范和 8 套填空模板
- ⭐ 自动搜索互联网最新数据补充稿件说服力（装机量、政策目标等）
- ⭐ 国资央企文风铁律（标题规范+7大类禁用词+速查表）
- 所有模板均已脱敏，可直接部署到公开仓库

**文件结构**：
```
CCTV-NEWS/
├── SKILL.md              # V3.0 完整工作流程 + 标题命名规范 + 数据补充规则 + 文风铁律
└── references/
    └── template.md        # V3.0 表达范式模板（8型+词库+速查表+禁用词清单）
```

**使用方式**：
```
用户：公司董事长今天去拜访了某市市长，谈了十五五新能源布局和储能建设，写一篇新闻稿
→ Agent 自动判断类型 → 搜索互联网补充数据 → 按模板生成稿件
```

### BIZ-PPT-SKILL

**功能**：给定数据和一份版式蓝本（总部/上级单位的分析会报告），产出可直接套公司母版的正式汇报 PPT；同一批数据可出四套样式。

**四套样式**：

| 样式 | 页数 | 底色 | 定位 | 逐字稿 |
|---|---|---|---|---|
| A 详细版 | 30 | 白 | 内部存档、季度分析会，明细全进表格 | 可选 |
| A′ 简略版 | 15 | 白 | 限时汇报，同类证据合并成页 | **必配** |
| B 财报故事版 | 15 | 白 | 对外讲述，一页一个指标，guidance 收尾 | **必配** |
| C 第一性原理版 | 14 | 近黑 | 内部研讨，巨数字 + 判断句，往下压结论 | **必配** |

**特点**：
- **母版安全区铁律**：顶部 0—0.95in 全空、不画横线、不画页码，底部止于 6.72in——套公司母版后自动补齐 logo 与页脚，不压字
- **矢量图表**：所有柱/条/折线由形状绘制，不用 PPT 原生图表对象。实测在线文档回存会把原生图表掏空（`c:cat`/`c:val` 丢失），形状图表不会被任何转换器破坏
- **一版一色卡**：色板全部从公司 VI 色卡取，缺正红时统一用负向橙；深色版有专门的高亮档位映射
- **字号阶梯**：正文 14pt 下限、结论句 20pt 带完整标点、图表类目 12—13pt，两套密度共用
- **强制配套 Word 逐字稿**：A′/B/C 页数少、版面删得干净，只给 PPT 讲不透，缺讲稿视为未交付
- **零动画**：成品不加转场与元素动画，翻页即硬切；节奏感用静态手段（竖向进度脊）解决
- 完整避坑表：负值柱空心、表格自带蓝白样式、批量导 PNG 出现旧图、`prs.save()` 被 PowerPoint 占用等 10 条

**文件结构**：
```
BIZ-PPT-SKILL/
├── SKILL.md                            # 选型、安全区、字号、配色、工作流、避坑、交付
├── README.md                           # 安装与使用
└── references/
    ├── deck_lib_template.py            # 版式库（色板/页骨架/KPI卡/表格/矢量图表）
    ├── standard_30_template.py         # 详细版骨架（含 30 页页序清单）
    ├── standard_15_template.py         # 简略版骨架（五条合并规则）
    ├── apple_earnings_template.py      # 财报故事版骨架
    ├── dark_firstprinciples_template.py# 第一性原理深色版骨架
    └── apple_script_template.html      # Word 逐字讲稿 HTML 模板（含 CSS）
```

> `references/` 里是**脱敏骨架**：页型、坐标、函数写法完整保留，文案为占位词、数值为演示值，照着扩页换数即可。

**使用方式**：
```
用户：@BIZ-PPT-SKILL 照这份总部报告的结构，用我们的生产数据做一版季度汇报 PPT
[附件：总部报告.pdf + 公司运营情况汇报.pptx]
→ Agent 拆版式 → 抽数据 → 选样式 → 建 build.py → 逐页导出 PNG 自查 → 交付 PPTX + PDF
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
        "$env:USERPROFILE\.workbuddy\skills\CCTV-NEWS"
        # BIZ-PPT-SKILL 不在此列：仓库里放的是脱敏骨架，
        # 本地工作版含真实公司数据，不要直接覆盖上来。
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
