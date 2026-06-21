# 项目状态说明

更新时间：2026-06-18

## 当前结论

本项目已经不是初始搭建阶段，当前处于“可运行试用 + 多产品模板接入 + 验证整理中”的阶段。主程序、核心生成流程、产品扩展结构、材质单管理、Word/PDF 输出能力都已经存在。

下一阶段重点是：补齐验收测试、整理文档、修复已知细节问题，并使用 PyInstaller 打包为 Windows 可执行程序。

## 已完成能力

- 桌面端主程序入口：`main.py`
- PySide6 主界面：产品选择、Excel 导入、参数预览/微调、Word 生成、PDF 生成、打开归档目录
- Excel 解析：`core/excel_reader.py`
- Word 报告生成：`core/word_engine.py`
- PDF 转换：`core/pdf_converter.py`
- 输出归档：`core/archiver.py`
- 多产品扩展框架：`products/base_product.py`
- 材质单图片管理：`core/material_manager.py`、`ui/image_manager.py`
- 分章节组合/模板处理：`core/doc_composer.py`、`core/template_preparer.py`、`core/template_converter.py`
- 动态型号页数/表格调整：`core/dynamic_adjuster.py`
- 黏滞阻尼器滞回曲线生成：`core/hysteresis_curve.py`

## 已接入产品

| 产品类型 | 目录 | 状态 |
| --- | --- | --- |
| 隔震支座 | `products/isolation_bearing/` | 已接入，早期完成端到端验证 |
| 黏滞阻尼器 | `products/viscous_damper/` | 已接入，含滞回曲线生成逻辑 |
| 摩擦摆支座 | `products/friction_pendulum/` | 已接入，有模板、Excel、配置和验证输出 |
| 阻尼器预埋件 | `products/embedded_damper_parts/` | 已接入，有模板、Excel、配置和验证输出 |

## 验证情况

- 2026-05-26 开发日志记录：隔震支座已完成 Excel 到 Word 报告再到归档的端到端测试。
- 2026-05-27 开发日志记录：处理了 Word 分页、空白页、VML 图片误删、分节符等模板格式问题。
- `output/` 中已有 2026-06-17 至 2026-06-18 的摩擦摆支座、阻尼器预埋件等验证输出目录，说明后续产品已经做过生成/渲染检查。

## 当前主要风险

- 早期文档没有持续更新，历史计划和实际代码进度不一致；已通过本文件和 `todo.md` 重新整理。
- 仍缺少一套正式的“每个产品一键验收清单”，需要确认四类产品都能稳定生成 Word/PDF。
- 打包分发尚未完成，未验证无 Python 环境的 Windows 机器运行情况。
- 输出目录、材质单库和临时文件包含大量业务数据，当前已在 `.gitignore` 中默认排除，后续如需纳入版本控制再单独放开。
- Git 仓库已初始化，但当前环境对 `.git` 目录存在显式写入拒绝权限。已提供 `scripts/first_commit.ps1`，可在用户本机 PowerShell 中执行首次提交。

## 文档整理结果

仍保留的有效文档：

- `project_status.md`：当前项目状态总览
- `todo.md`：下一步开发和验收任务
- `AGENTS.md`：项目协作与开发指引
- `CLAUDE.md`：历史协作指引，暂保留
- `docs/Word模板格式处理经验.md`：Word 模板格式问题经验库
- `docs/粘滞阻尼器滞回曲线生成说明.md`：黏滞阻尼器曲线生成说明
- `dev_logs/`：历史开发日志
- `.gitignore`：版本控制忽略规则，默认排除输出、缓存、调试渲染和本地业务素材

已归档的早期文档：

- `docs/archive/2026-06-18_early_docs/需求文档.md`
- `docs/archive/2026-06-18_early_docs/技术方案.md`
- `docs/archive/2026-06-18_early_docs/设计规范.md`
- `docs/archive/2026-06-18_early_docs/开发步骤.md`

这些文档仍可作为历史参考，但不再作为当前进度依据。

## 运行方式

在项目根目录运行：

```powershell
python main.py
```

PDF 转换依赖本机安装 Microsoft Word。
