# Word 模板格式处理 — 遇到的问题与解决办法

> 适用场景：python-docx 处理 .docx 模板，通过 XML 操作调整排版格式。
> 本文档记录隔震支座报告开发中遇到的所有格式相关问题及其解决方案，供设计其他产品报告时参考。

---

## 一、核心原则

1. **操作 XML 元素前先快照**：`list(doc.element.body)` 获取快照后再遍历，因为删除/插入会改变 live body。
2. **删除段落前多重检查**：空段落 ≠ 可删除。必须检查是否包含图片、分页符、分节符。
3. **处理后验证**：生成报告后通过解析 XML 验证关键元素数量（图片段落数、分页符数、分节数等）。
4. **VML vs DrawingML**：旧版 .doc 转 .docx 的图片可能使用 `w:pict`（VML），而非 `w:drawing`，检查函数需同时覆盖。

---

## 二、必备辅助函数

```python
from docx.oxml.ns import qn

def _elem_tag(elem):
    """返回元素标签的简写，如 'p', 'tbl', 'sectPr'"""
    return elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

def _elem_text(elem):
    """提取段落/元素中所有 w:t 的文本"""
    return ''.join(t.text or '' for t in elem.iter(qn('w:t'))).strip()

def _has_page_break(elem):
    """检查是否含硬分页符 w:br w:type='page'"""
    return any(br.get(qn('w:type')) == 'page' for br in elem.iter(qn('w:br')))

def _has_image(elem):
    """检查是否含图片：VML(w:pict) + DrawingML(w:drawing) + OLE(w:object)"""
    return any(elem.findall('.//' + qn(tag))
               for tag in ('w:drawing', 'w:pict', 'w:object'))

def _has_section_break(elem):
    """检查段落是否含有分节符 w:sectPr"""
    pPr = elem.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:sectPr')) is not None
```

---

## 三、问题清单

### 问题 1：图片段落被误删（VML 图片）

**现象**：营业执照页消失、8 页材质书只剩 1 页、封面印章丢失。

**根因**：`_remove_blank_pages()` 和 `_compact_certificates()` 通过 `not _elem_text(child)` 判断"空段落"然后删除。但 VML 图片（`w:pict`，常用于印章、扫描件）没有 `w:t` 文本，`_elem_text()` 返回空串，被误判为"空段落"删除。

**解决**：新增 `_has_image()` 辅助函数，检查 `w:drawing`、`w:pict`、`w:object` 三种图片/OLE 元素。在所有删除空段落的判断中加入 `not _has_image(child)` 条件。

**教训**：
- 旧版 .doc 转 .docx 的图片大概率是 VML 格式（`w:pict`），必须同时检查
- 印章、扫描件、营业执照等常以 VML 图片形式存在
- 所有"删除空段落"的逻辑都需要加入图片检查

---

### 问题 2：分节符 + 分页符 共存导致空白页

**现象**：产品合格证和力学性能检验报告后面各有一页空白页。

**根因**：模板中某些段落同时含有 `w:br type="page"`（硬分页符）和 `w:sectPr`（分节符，其类型为 NEW_PAGE）。两者各自独立触发换页 → 一次换页（分页符）+ 一次换页（分节符） = 中间多出一页空白。

关键 OOXML 知识：
- `w:br w:type="page"` — 硬分页，当前段落之前插入分页
- `w:sectPr` 在 `w:pPr` 中 — 标记节结束，如果下一节 `type="NEW_PAGE"` 则在下一节开始前换页
- 两者共存在同一段落 → **双倍换页** → 空白页

**解决**：新增 `_fix_section_breaks()` 函数，在生成报告的 pipeline 中早期运行（填充完数据后、紧凑排版前）。遍历所有段落，如果段落已有 `w:sectPr`，则移除其中的 `w:br type="page"`。分节符本身的 NEW_PAGE 类型已经提供了正确的一次换页。

**教训**：
- `_fix_section_breaks()` 需要在删除空段落之前运行，否则含分节符的段落可能被先删掉
- 移除分页符后该段落变成"空段落"（无文本、无图片），此时 `_has_section_break()` 保护它不被后续的清理函数删除
- WPS/Word 创建 .doc 再转 .docx 时，这种冗余分页符很常见

---

### 问题 3：含分节符的空段落被误删（V7 复现）

**现象**：修复问题 2 后，合格证和力学报告之间的分节边界消失，两者挤在同一页。

**根因**：`_fix_section_breaks()` 移除了含 `w:sectPr` 段落中的冗余 `w:br`，该段落变成"纯分节符段落"（无文本、无图片、无分页符）。但 `_compact_certificates` 的 step 3b 和 `_remove_blank_pages` 判断空段落时只检查了"无文本 + 无分页符 + 无图片"，分节符段落通过这三个检查后被删除。删除分节符 = 两个节合并。

**解决**：新增 `_has_section_break()` 辅助函数，在所有删除空段落的判断中加入 `not _has_section_break(child)` 条件。受影响的函数：
- `_compact_certificates` step 3（合格证组间空段落）
- `_compact_certificates` step 3b（最后一个合格证表后）
- `_remove_blank_pages`（连续空段落折叠）

**教训**：
- 分节符是实现"下一页"分节的 OOXML 机制，被删除会导致两个节合并、页码/页边距继承错乱
- 删除空段落的三重检查顺序：无文本 → 无图片 → 无分页符 → **无分节符**
- `_fix_section_breaks` 和删除空段落的函数之间有执行顺序依赖

---

### 问题 4：合格证表格跨页断裂

**现象**：一个合格证组（标题 + 表格）被分到两页。

**解决**：新建 `_fix_certificate_pagination()`：
1. 标题段落设 `keepWithNext`（段落属性）—— 标题与后续内容不分离
2. 标题和表格之间的空段落也设 `keepWithNext`
3. 表格每一行设 `cantSplit`（表格行属性）—— 单行不跨页
4. 除最后一行外，每行单元格内所有段落设 `keepWithNext`（行间不分离）

**教训**：
- `cantSplit` 只能防止单行跨页，不能防止行与行之间分离
- 需要配合 `keepWithNext` 才能将标题+整个表格锁定在一起
- 最后一行的段落不设 `keepWithNext`，否则可能与后续内容卡在一起

---

### 问题 5：合格证页占用过多空间

**现象**：6 个合格证占 6 页，每页大量空白。

**解决**：新建 `_compact_certificates()`：
- 行高：`w:trHeight` 从 603 twips → 360 twips，`hRule="atLeast"`
- 段落行距：`w:spacing` → `before=0, after=0, line=280, lineRule=auto`
- 删除合格证组之间的空段落（跳过含分页符/图片/分节符的段落）
- 效果：2 个合格证能放在一页

**教训**：
- `hRule="atLeast"` 表示"最小行高"，行高会自动扩展容纳内容
- 行距 `line=280` (twips) 约等于单倍行距
- 仅压缩行高不够，还需要压缩段落间距才能显著缩小

---

### 问题 6：多空段落序列形成空白页

**现象**：连续多个空段落（2+个）在页面顶部形成大量空白。

**解决**：新建 `_remove_blank_pages()`：
- 扫描 body 中连续 2+ 个空段落
- 保留第一个，删除其余
- 跳过含分页符、图片、分节符的段落
- 效果：移除 46 个冗余空段落

**教训**：
- 连续空段落是 .doc → .docx 转换的常见副产品
- 必须保留含分页符的空段落（它标记着"此处换页"的意图）
- 只折叠连续空段落，不删除孤立的单个空段落（可能是有意排版）

---

### 问题 7：目录与合格证挤在同一页

**现象**：目录后直接开始合格证，没有分页。

**解决**：新建 `_add_toc_page_break()`：
- 查找第一个"产品合格证"标题
- 在它前面的最近一个段落插入 `w:br type="page"`
- TOC 单独成页

**教训**：
- 这是最直接的分页方式 —— 在关键节点前插入硬分页符
- 适用于"X 内容之前必须换页"的场景

---

### 问题 8：外观检测表行距过大

**现象**：外观检测报告表格行高不均匀，部分内容行过高。

**解决**：新建 `_compact_visual_tables()`：
- 特殊行（3/4/5/8 行）保持 567 twips（内容较多的行）
- 其余行压缩为 280 twips，段落行距 240
- 遍历所有"隔震橡胶支座外观质量及尺寸检测报告"标题后的表格

**教训**：
- 并非所有行都需要同样高度 → 按行索引区分处理
- 行高值需根据实际内容通过几次迭代确定

---

## 四、生成流水线（Pipeline Order）

`generate_report()` 中后处理步骤的执行顺序很重要：

```
1. DynamicAdjuster          — 按型号数量增删页
2. _replace_all             — 替换占位符 {{FIELD_XXX}}
3. _direct_fill_all_tables  — 直接填充无占位符的表格单元格
4. _fix_section_breaks      — 移除分节符段落中的冗余分页符（★ 必须在删空段落前）
5. _fix_certificate_pagination — 合格证表格防跨页
6. _compact_certificates    — 压缩合格证行距+删除组间空段落
7. _remove_blank_pages      — 折叠连续空段落
8. _add_toc_page_break      — TOC 独立分页（最后做，避免被前面删除）
9. _compact_visual_tables   — 外观检测表行距压缩
```

**顺序要点**：
- `_fix_section_breaks` 必须在任何删空段落的步骤前执行
- `_add_toc_page_break` 在删除空段落后执行，避免插入的分页符被清理
- `_fix_certificate_pagination` 在紧凑化之前执行，确保 keepWithNext 在压缩后的行上生效

---

## 五、调试技巧

### 5.1 解析生成文档的关键指标

```python
from docx.oxml.ns import qn

body = doc.element.body
children = list(body)

# 统计各类元素
tags = {}
for c in children:
    t = c.tag.split('}')[-1] if '}' in c.tag else c.tag
    tags[t] = tags.get(t, 0) + 1
print(tags)

# 统计 VML 图片段落数
vml = sum(1 for c in children
          if (c.tag.split('}')[-1] if '}' in c.tag else c.tag) == 'p'
          and c.findall('.//' + qn('w:pict')))
print(f'VML paragraphs: {vml}')

# 统计分页符
breaks = sum(1 for c in children
             for br in c.iter(qn('w:br'))
             if br.get(qn('w:type')) == 'page')
print(f'Page breaks: {breaks}')

# 统计分节符
sections = sum(1 for c in children
               if c.find(qn('w:pPr')) is not None
               and c.find(qn('w:pPr')).find(qn('w:sectPr')) is not None)
print(f'Section breaks: {sections}')

# 列出 body 中每个子元素及其文本和特征
for i, c in enumerate(children):
    tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
    text = ''.join(t.text or '' for t in c.iter(qn('w:t')))[:60]
    flags = []
    if c.findall('.//' + qn('w:pict')): flags.append('VML')
    if c.findall('.//' + qn('w:drawing')): flags.append('DML')
    if any(br.get(qn('w:type'))=='page' for br in c.iter(qn('w:br'))): flags.append('PAGE_BREAK')
    if c.find(qn('w:pPr')) is not None and c.find(qn('w:pPr')).find(qn('w:sectPr')) is not None:
        flags.append('SECT_BREAK')
    if flags:
        print(f'  [{i}] {tag} {"|".join(flags)} text="{text}"')
```

### 5.2 OOXML 关键元素速查

| 元素 | 路径 | 说明 |
|------|------|------|
| 硬分页符 | `w:r/w:br w:type="page"` | 段落前换页 |
| 分节符 | `w:pPr/w:sectPr` | 节结束，新节可设 NEW_PAGE |
| 行高 | `w:trPr/w:trHeight` | val 单位 twips，hRule="atLeast" 为最小 |
| 段落间距 | `w:pPr/w:spacing` | before/after/line，单位 twips |
| VML 图片 | `w:r/w:pict` | 旧版 Office 图片格式 |
| DrawingML | `w:r/w:drawing` | 新版 Office 图片格式 |
| 嵌入对象 | `w:r/w:object` | OLE 对象 |
| keepWithNext | `w:pPr/w:keepNext` | 段落与后续内容保持同页 |
| cantSplit | `w:trPr/w:cantSplit` | 表格行不可跨页 |

---

## 六、新增产品类型的注意事项

1. **优先排查 VML 图片**：如果原模板是 .doc 格式转来的，印章、扫描件大概率是 VML，检查 `_has_image()` 是否覆盖。
2. **检查节结构**：用调试脚本列出 body 中所有 SECT_BREAK 和 PAGE_BREAK，确认分节逻辑正确。
3. **空段落默认不要删**：先了解每个空段落的作用（分页？留白？图片占位？），再决定处理方式。
4. **压缩表格前先看内容**：行高和行距的值需根据实际字体大小/行数调试确定，不同内容密度需要不同的值。
5. **流水线顺序不变**：上述 Pipeline 顺序已在隔震支座产品上验证，新增产品直接沿用。
6. **动态增删页面**：如果模板中有按型号重复的区块（合格证组、检测报告组等），需实现类似 `DynamicAdjuster` 的逻辑。
