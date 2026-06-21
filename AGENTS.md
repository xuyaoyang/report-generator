# AGENTS.md — 检测报告生成软件项目指引

## 项目概述
Windows 桌面应用，通过 Excel 参数表驱动 Word 模板，自动生成隔震支座出厂检验报告，支持 Word/PDF 输出和自动归档。

## 文档索引

| 文档 | 路径 | 说明 |
|------|------|------|
| 需求文档 | [docs/需求文档.md](docs/需求文档.md) | 功能需求、非功能需求 |
| 技术方案 | [docs/技术方案.md](docs/技术方案.md) | 技术选型、架构设计、核心流程 |
| 设计规范 | [docs/设计规范.md](docs/设计规范.md) | UI配色、布局、代码规范 |
| 开发步骤 | [docs/开发步骤.md](docs/开发步骤.md) | 分阶段开发任务清单 |
| 格式处理经验 | [docs/Word模板格式处理经验.md](docs/Word模板格式处理经验.md) | OOXML格式问题与解决方案汇总 |

## 开发日志
开发日志位于 [dev_logs/](dev_logs/) 目录，按日期命名（如 `2026-05-26.md`）。
每个工作日结束时自动记录完成事项和待办事项。

## 工作说明

### 开发原则
- **稳定优先**：每完成一个阶段验证后再进入下一阶段
- **小步推进**：一次只改一个模块，避免大面积修改
- **安全第一**：文件操作前备份，不删除用户原始模板
- **用户视角**：始终记住用户是不懂代码的业务人员，界面要简单直观

### 目录结构
```
报告生成/
├── main.py                 # 程序入口
├── config/                 # 配置文件
│   ├── settings.json       # 软件设置
│   └── products.json       # 产品注册
├── core/                   # 核心引擎
│   ├── excel_reader.py
│   ├── word_engine.py
│   ├── pdf_converter.py
│   └── archiver.py
├── products/               # 产品模块（可扩展）
│   ├── base_product.py
│   └── isolation_bearing/
├── ui/                     # GUI 界面
│   ├── main_window.py
│   ├── param_panel.py
│   ├── image_manager.py
│   └── resources/
├── image_lib/              # 材质单图片库
├── output/                 # 默认输出目录
├── docs/                   # 项目文档
├── dev_logs/               # 开发日志
└── 隔震支座出厂报告2025-2-7.doc  # 原始模板（只读，不修改）
```

### 技术栈
- Python 3.11 + PySide6 + python-docx + openpyxl + pywin32 + docx2pdf
- 打包工具：PyInstaller

### 关键约定
- 所有文件操作使用绝对路径
- 原始模板文件只读，不直接修改
- 配置与代码分离
- 新增产品类型只需在 products/ 下添加目录+配置


# AGENTS.md

Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

---

**These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.