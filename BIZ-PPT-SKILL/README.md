# BIZ-PPT-SKILL　企业级汇报 PPT 统一样式库

一个给 WorkBuddy / Claude Code 用的 PPT 生产技能：给定数据和一份版式蓝本，产出**可直接套公司母版**的正式汇报材料。
四套可选样式，全部遵守同一套安全区规范、字号阶梯与色卡约束；页数少的版本强制配套 Word 逐字讲稿。

## 一、装到哪里

把整个 `BIZ-PPT-SKILL` 文件夹复制到本机**用户级技能目录**：

```
Windows:  C:\Users\<你的用户名>\.workbuddy\skills\BIZ-PPT-SKILL
macOS:    ~/.workbuddy/skills/BIZ-PPT-SKILL
```

放好后目录结构应为：

```
BIZ-PPT-SKILL/
├── SKILL.md
├── README.md
└── references/
    ├── deck_lib_template.py            # 版式库（色板 / 页骨架 / KPI 卡 / 表格 / 矢量图表）
    ├── standard_30_template.py         # 详细版骨架（页序 + 每页写法）
    ├── standard_15_template.py         # 简略版骨架（五条合并规则）
    ├── apple_earnings_template.py      # 财报故事版骨架（一页一指标）
    ├── dark_firstprinciples_template.py# 第一性原理深色版骨架（巨数字 + 判断句）
    └── apple_script_template.html      # Word 逐字讲稿的 HTML 模板（含 CSS）
```

**重启 WorkBuddy**（或在技能管理里刷新）后生效。之后说「用 BIZ-PPT-SKILL 做一版季度汇报」即可触发。

> `references/` 里的四个 `.py` 是**脱敏骨架**，不是完整成品：页型、坐标、函数写法完整保留，文案为占位词、数值为演示值（`30.00`、`93.8` 这类好认的整数）。照骨架扩页、换数即可。

## 二、依赖

PPT 生成依赖 `python-pptx`（含 `lxml`）：

```bash
python -c "import pptx; print(pptx.__version__)"
pip install python-pptx        # 没装时
```

导 PDF 与逐页 PNG 自检用到 PowerShell 调 PowerPoint，因此**建议装 Microsoft PowerPoint**。没有也能生成 PPTX，只是不能自动导 PDF 和逐页自检。

## 三、四套样式怎么选

| 样式 | 页数 | 适用 | 讲稿 |
|---|---|---|---|
| A 详细版 | 30 | 内部存档、季度分析会 | 可选 |
| A′ 简略版 | 15 | 限时汇报、周会 | **必配** |
| B 财报故事版 | 15 | 对外讲述、对上汇报 | **必配** |
| C 第一性原理版 | 14 | 内部研讨、推动决策 | **必配** |

口径规则：**A 走单专业口径**（只讲本专业，其余整章剥离）；**A′、B、C 都是全口径**（讲完整因果链）。
A′ 的"少"是合并，不是删除——不要把 A 的删减逻辑套到 A′ 上。

## 四、字体（影响观感，不影响能否打开）

- 中文：微软雅黑 / 等线（Windows 自带）
- 数字：Arial（B）、Bahnschrift（C）

Windows 10 及以上均自带。macOS 上会被替换成近似字体，版式不散但观感有差异，建议在 Windows 机器上出稿。

## 五、使用要点

1. 生成前先确认三件事：用哪套样式、页数、数据口径（单专业 or 全口径）
2. 交付物默认是 **PPTX + PDF 双份**；A′、B、C 还要加一份 Word 逐字稿
3. 生成的页面**不含 logo、顶部横线、页码**——这三样由公司母版提供，套模板后自动补齐
4. 所有页面**不带任何动画与转场**（硬性规范，勿再加）
5. 改数改版：改 `build_*.py` 里对应的数值/文案，重跑一次即可，版式不会乱
6. 图表**不要用 PPT 原生图表对象**——在线文档/WPS 回存会把图表掏空，必须用 `schart` 形状绘制

## 六、常见问题

**Q：放进去了但技能没被调用？**
A：确认文件夹名是 `BIZ-PPT-SKILL`，且 `SKILL.md` 就在这一层（不要多套一层目录）。然后重启 WorkBuddy。

**Q：生成时报 `No module named pptx`？**
A：缺依赖，执行 `pip install python-pptx`。

**Q：套上公司母版后顶部压字？**
A：母版安全区和我这边量的不一样。打开母版视图量出 logo 与横线的实际位置，改 `deck_lib.py` 顶部 `TITLE_Y / HEAD_Y / BODY_T / BODY_B / FOOT_Y` 五个常量即可。

**Q：导 PDF 报错？**
A：本机没装 PowerPoint，或目标文件正被打开占用。关掉文件重试。

**Q：出来的字和我这边不一样？**
A：字体缺失导致替换。Windows 10 以上一般正常；若在 Mac 上用，建议改到 Windows 出稿。

**Q：能不能加动画？**
A：不能。这套体系成品一律不加转场与元素动画，翻页即硬切。要节奏感请用静态手段（深色版的竖向进度脊、满版色块起伏）。

## 许可证

MIT License - 自由使用，自负责任。
