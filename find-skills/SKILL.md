---
name: find-skills
description: 帮助用户发现、搜索和安装 Agent Skills。当用户问"有没有做 X 的 Skill"、"帮我找一个 Skill"、"怎么扩展你的能力"、"你能做 X 吗"（某项专门能力）时触发。也在用户想搜索工具/模板/工作流，或想了解 Skills 生态时触发。即使用户只是问"有没有现成的方案"也应考虑使用此 Skill。
---

# Find Skills

帮助用户从开放的 Agent Skills 生态中发现、搜索并安装合适的 Skill。

## 核心目标

当用户需要某项专门能力时，不要直接从零开始，而是先检索是否有现成 Skill 可用。

---

## 工作流程

### Step 1：理解需求

分析用户意图，识别：
- 用户想完成什么任务？
- 涉及哪个领域？（如：数据分析、文档生成、代码、设计…）
- 这个任务是否专门/复杂到足以存在独立 Skill？

### Step 2：搜索 Skills

**主要搜索渠道：**

1. **skills.sh** — 官方 Skill 目录网站（https://skills.sh/）
2. **vercel-labs/skills** — Vercel 官方 GitHub 仓库（https://github.com/vercel-labs/skills）
3. **Web 搜索** — 用 `site:github.com skills SKILL.md [关键词]` 搜索社区 Skill

**搜索策略：**
- 用英文关键词搜索效果更好
- 尝试多个同义词/近义词
- 查看仓库的 README 和 skills/ 目录结构

### Step 3：展示结果

找到候选 Skill 后，告诉用户：
- Skill 名称和来源仓库
- 核心用途（一句话描述）
- 安装命令
- 是否完全匹配需求，或只是部分匹配

### Step 4：安装

**标准安装命令格式：**
```bash
npx playbooks add skill <owner/repo> --skill <skill-name>
```

**例如安装 find-skills 本身：**
```bash
npx playbooks add skill vercel-labs/skills --skill find-skills
```

**在 Claude.ai 环境中安装：**
由于 Claude.ai 的技能安装通过界面完成，引导用户：
1. 在 Claude.ai 设置中找到 Skills/技能管理
2. 使用 `.skill` 文件上传安装

### Step 5：如果没有找到合适的 Skill

- 告诉用户目前没有现成 Skill 满足需求
- 提议用 **skill-creator** Skill 创建一个新的
- 或直接帮用户完成当前任务

---

## 常用 Skills 目录

| 类别 | 常见 Skill | 来源 |
|------|-----------|------|
| 文档处理 | docx, pdf, pptx, xlsx | 内置 |
| 数据分析 | csv-excel-analyst, operations-ledger-analyst | 内置/用户 |
| 前端设计 | frontend-design | 内置 |
| 信息图 | baoyu-infographic | 用户 |
| Skill 管理 | skill-creator, find-skills | 内置/vercel-labs |

---

## 关键命令速查

```bash
# 搜索 Skill（如果有 npx skills 环境）
npx skills find [query]

# 安装 Skill
npx skills add <owner/repo@skill>

# 检查更新
npx skills check

# 更新所有 Skill
npx skills update
```

浏览所有可用 Skill：**https://skills.sh/**

---

## 注意事项

- 在 Claude.ai 环境中，无法直接运行 `npx` 命令，需引导用户通过界面或文件方式安装
- 安装新 Skill 后，需要刷新对话才能生效
- 用户自定义 Skill 优先级高于公共 Skill
