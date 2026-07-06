# weekly-ai-report

AI 应用情报周报自动化生成 Skill，适用于 WorkBuddy / Claude Code 等 AI 编程助手。

## 功能

一键生成 AI 应用情报周报（.docx 格式），包含：

- **新闻类**：按日期分组的 AI 行业新闻（14-21 条），每条含标题、摘要、来源链接
- **项目类**：AI 娱乐/泛娱乐赛道创业项目表格（5-7 个），含公司名、描述、融资、团队、来源

## 安装

将 `weekly-ai-report/` 文件夹复制到你的 WorkBuddy skills 目录：

**项目级**（仅当前项目生效）：
```
your-project/.workbuddy/skills/weekly-ai-report/SKILL.md
```

**用户级**（所有项目生效）：
```
~/.workbuddy/skills/weekly-ai-report/SKILL.md
```

## 使用

在 WorkBuddy 中直接说：

- "生成本周 AI 周报"
- "生成一份本周的 AI 周报，格式和之前保持一致"

Skill 会自动加载，按标准流程完成新闻采集、项目筛选、格式排版和文件交付。

## 自定义

如需适配你自己的业务场景，主要修改 `SKILL.md` 中的：

- **第一节**：新闻源、项目筛选原则（赛道、排除项）
- **第六节**：用户偏好（语言、信息源、定位）
- **第二节**：如需调整 docx 格式（字体、字号、颜色等）

OOXML 格式规范以最新一期周报为模板，修改前建议先分析你的参考文档的 XML 结构。

## 许可

MIT
