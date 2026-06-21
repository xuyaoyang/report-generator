"""
Dynamic template adjustment: add or remove certificate pages and
visual inspection report pages based on the actual number of product models.

Strategy:
- N < 6: delete extra certificate/visual groups from the template
- N > 6: clone the last group and fill cloned placeholders with actual data
"""
import re
import copy
from docx.oxml.ns import qn


class DynamicAdjuster:
    """Adjust document structure based on model count."""

    def __init__(self, doc):
        self.doc = doc
        self.body = doc.element.body
        self._build_index()

    def _build_index(self):
        """Index body children."""
        self._para_by_text = {}
        for i, child in enumerate(list(self.body)):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'p':
                text = ''
                for t in child.iter(qn('w:t')):
                    if t.text:
                        text += t.text
                text = text.strip()
                if text:
                    self._para_by_text.setdefault(text, []).append((i, child))

    # =================================================================
    # Certificate adjustment
    # =================================================================

    def _find_cert_groups(self):
        """Find (start_body_idx, end_body_idx) for each certificate group."""
        cert_text = '产品合格证'
        cert_refs = self._para_by_text.get(cert_text, [])
        if not cert_refs:
            return []

        groups = []
        current_children = list(self.body)
        for body_idx, para_elem in cert_refs:
            tbl_body_idx = None
            for j in range(body_idx + 1, len(current_children)):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'tbl':
                    tbl_body_idx = j
                    break
                elif tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    if text.strip():
                        break
            if tbl_body_idx is None:
                continue

            # Scan backwards for preceding empty paragraphs
            start_idx = body_idx
            for j in range(body_idx - 1, -1, -1):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    if not text.strip():
                        start_idx = j
                    else:
                        break
                else:
                    break

            groups.append((start_idx, tbl_body_idx))

        return groups

    def _remove_cert_groups(self, count):
        """Remove the last `count` certificate groups."""
        groups = self._find_cert_groups()
        if count <= 0 or count > len(groups):
            return

        to_remove = groups[-count:]
        current_children = list(self.body)
        removed = set()
        for start, end in to_remove:
            for i in range(start, end + 1):
                removed.add(current_children[i])
        for elem in removed:
            self.body.remove(elem)
        self._build_index()  # Rebuild index after body change
        print(f'Removed {count} certificate group(s), '
              f'{len(groups) - count} remaining.')

    def _clone_cert_group(self, product_data):
        """Clone the last certificate group and fill with product_data inline.
        product_data: dict with 产品型号, 支座编号范围, 生产日期, 数量, 检验标准."""
        groups = self._find_cert_groups()
        if not groups:
            return

        last_start, last_end = groups[-1]
        current_children = list(self.body)
        template_elements = current_children[last_start:last_end + 1]

        # Field order matches the certificate table structure:
        # Row 1 = 产品型号, Row 2 = 支座编号范围, Row 3 = 生产日期,
        # Row 4 = 数量, Row 5 = 检验标准
        cert_values = [
            str(product_data.get('产品型号', '')),
            str(product_data.get('支座编号范围', '')),
            str(product_data.get('生产日期', '')),
            str(product_data.get('数量', '')),
            str(product_data.get('检验标准', '')),
        ]

        new_elements = []
        for elem in template_elements:
            cloned = copy.deepcopy(elem)
            new_elements.append(cloned)

            # Replace all {{FIELD_XXX}} with values from cert_values
            placeholder_idx = 0
            for t_elem in cloned.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, idx=[0]):
                    val_idx = idx[0]
                    idx[0] += 1
                    if val_idx < len(cert_values):
                        return cert_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

        # Insert after the last certificate group
        ref_elem = current_children[last_end]
        for elem in reversed(new_elements):
            ref_elem.addnext(elem)

        print(f'Cloned certificate group for model: '
              f'{product_data.get("产品型号", "?")}')

    # =================================================================
    # Visual inspection report adjustment
    # =================================================================

    def _find_visual_groups(self):
        """Find (start_body_idx, end_body_idx) for each visual report group."""
        visual_text = '隔震橡胶支座外观质量及尺寸检测报告'
        visual_refs = self._para_by_text.get(visual_text, [])
        if not visual_refs:
            return []

        groups = []
        current_children = list(self.body)

        for body_idx, para_elem in visual_refs:
            tbl_body_idx = None
            for j in range(body_idx + 1, min(len(current_children), body_idx + 5)):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'tbl':
                    tbl_body_idx = j
                    break
            if tbl_body_idx is None:
                continue

            # Find start: company name paragraph + signature/empty before it
            start_idx = body_idx
            for j in range(body_idx - 1, max(-1, body_idx - 5), -1):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    text = text.strip()
                    if '四川融海运通' in text:
                        start_idx = j
                        # Include preceding signature/empty
                        for k in range(j - 1, max(-1, j - 4), -1):
                            child2 = current_children[k]
                            tag2 = child2.tag.split('}')[-1] if '}' in child2.tag else child2.tag
                            if tag2 == 'p':
                                text2 = ''
                                for t in child2.iter(qn('w:t')):
                                    if t.text:
                                        text2 += t.text
                                if text2.strip() and ('检验员' in text2 or '审核' in text2):
                                    start_idx = k
                                elif not text2.strip():
                                    start_idx = k
                                break
                            else:
                                break
                        break

            groups.append((start_idx, tbl_body_idx))

        return groups

    def _remove_visual_groups(self, count):
        """Remove the last `count` visual groups."""
        groups = self._find_visual_groups()
        if count <= 0 or count > len(groups):
            return

        to_remove = groups[-count:]
        current_children = list(self.body)
        removed = set()
        for start, end in to_remove:
            for i in range(start, end + 1):
                removed.add(current_children[i])
        for elem in removed:
            self.body.remove(elem)
        self._build_index()
        print(f'Removed {count} visual group(s), '
              f'{len(groups) - count} remaining.')

    def _clone_visual_group(self, product_data, project_name=''):
        """Clone the last visual group and fill with data inline."""
        groups = self._find_visual_groups()
        if not groups:
            return

        last_start, last_end = groups[-1]
        current_children = list(self.body)
        template_elements = current_children[last_start:last_end + 1]

        # Visual fields: project name, quantity
        vis_values = [
            str(project_name),
            str(product_data.get('数量', '')),
        ]

        new_elements = []
        for elem in template_elements:
            cloned = copy.deepcopy(elem)
            new_elements.append(cloned)

            placeholder_idx = 0
            for t_elem in cloned.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, idx=[0]):
                    val_idx = idx[0]
                    idx[0] += 1
                    if val_idx < len(vis_values):
                        return vis_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

        ref_elem = current_children[last_end]
        for elem in reversed(new_elements):
            ref_elem.addnext(elem)

        print(f'Cloned visual group for model: '
              f'{product_data.get("产品型号", "?")}')

    # =================================================================
    # Mechanical detail report adjustment (Tables 7, 8)
    # =================================================================

    def _find_mech_detail_groups(self):
        """Find mechanical detail report groups by pattern:
        [company] → [title] → [table] → [notes...]
        Skips Table 6 summary where pattern is [company] → [table] → [title]."""
        groups = []
        current_children = list(self.body)

        for i, child in enumerate(current_children):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag != 'p':
                continue

            text = ''
            for t in child.iter(qn('w:t')):
                if t.text:
                    text += t.text
            text = text.strip()
            if '四川融海运通' not in text:
                continue

            # Detail group: next is title, then table
            if i + 2 >= len(current_children):
                continue
            next_child = current_children[i + 1]
            next_tag = next_child.tag.split('}')[-1] if '}' in next_child.tag else next_child.tag
            if next_tag != 'p':
                continue
            next_text = ''
            for t in next_child.iter(qn('w:t')):
                if t.text:
                    next_text += t.text
            if '隔震支座力学性能检验报告' not in next_text:
                continue

            tbl_child = current_children[i + 2]
            tbl_tag = tbl_child.tag.split('}')[-1] if '}' in tbl_child.tag else tbl_child.tag
            if tbl_tag != 'tbl':
                continue

            # Skip summary table (12+ columns in first row)
            tr_elems = tbl_child.findall(qn('w:tr'))
            if tr_elems and len(tr_elems[0].findall(qn('w:tc'))) >= 10:
                continue

            # Find start: include preceding signature/empty paragraphs
            start_idx = i
            for k in range(i - 1, max(-1, i - 5), -1):
                prev = current_children[k]
                prev_tag = prev.tag.split('}')[-1] if '}' in prev.tag else prev.tag
                if prev_tag == 'p':
                    prev_text = ''
                    for t in prev.iter(qn('w:t')):
                        if t.text:
                            prev_text += t.text
                    if prev_text.strip() and ('检验员' in prev_text or '审核' in prev_text):
                        start_idx = k
                    elif not prev_text.strip():
                        start_idx = k
                    else:
                        break
                else:
                    break

            # Find end: include notes paragraphs after the table
            end_idx = i + 2  # the table
            for k in range(i + 3, min(len(current_children), i + 6)):
                after = current_children[k]
                after_tag = after.tag.split('}')[-1] if '}' in after.tag else after.tag
                if after_tag == 'p':
                    after_text = ''
                    for t in after.iter(qn('w:t')):
                        if t.text:
                            after_text += t.text
                    if after_text.strip().startswith('注') or '标准要求' in after_text:
                        end_idx = k
                    else:
                        break
                else:
                    break

            groups.append((start_idx, end_idx))

        return groups

    def _remove_mech_detail_groups(self, count):
        """Remove the last `count` mechanical detail groups."""
        groups = self._find_mech_detail_groups()
        if count <= 0 or count > len(groups):
            return

        to_remove = groups[-count:]
        current_children = list(self.body)
        removed = set()
        for start, end in to_remove:
            for i in range(start, end + 1):
                removed.add(current_children[i])
        for elem in removed:
            self.body.remove(elem)
        self._build_index()
        print(f'Removed {count} mechanical detail group(s), '
              f'{len(groups) - count} remaining.')

    def _trim_mech_table_rows(self, tbl_body_idx, needed_models):
        """Trim a mechanical detail table to `needed_models` data rows.
        Each model group = 4 table rows (section header + sub-headers + data row)."""
        current_children = list(self.body)
        tbl_elem = current_children[tbl_body_idx]

        tr_elems = tbl_elem.findall(qn('w:tr'))
        total_rows = len(tr_elems)
        needed_rows = needed_models * 4

        if needed_rows >= total_rows:
            return

        rows_to_remove = total_rows - needed_rows
        for i in range(rows_to_remove):
            tbl_elem.remove(tr_elems[-(i + 1)])

        print(f'Trimmed {rows_to_remove} rows from mech detail table, '
              f'{needed_rows} rows remaining ({needed_models} models).')

    def _clone_mech_detail_group(self, models_batch, project_info, mech_by_model):
        """Clone the last mechanical detail group and fill with data for models_batch.
        models_batch: list of up to 4 product dicts.
        mech_by_model: dict of model_name → mechanical data dict."""
        groups = self._find_mech_detail_groups()
        if not groups:
            return

        last_start, last_end = groups[-1]
        current_children = list(self.body)
        template_elements = current_children[last_start:last_end + 1]

        judgment_keys = [
            '设计压缩刚度(KN/mm)', '水平等效刚度(KN/mm)',
            '屈服后刚度(KN/mm)', '屈服力(KN)', '等效阻尼比(%)'
        ]

        # Build value list matching placeholder order in the template
        # Table 8 model 0 row: cols 0,1,2,3 → 4 placeholders
        # Table 8 model 1-3 rows: cols 0-6 → 7 placeholders each
        mech_values = []
        for i, product in enumerate(models_batch):
            model_name = product.get('产品型号', '')
            serial = product.get('支座编号范围', '')
            mech = mech_by_model.get(model_name, {})

            mech_values.append(str(model_name))
            mech_values.append(str(serial))

            num_judgment = 2 if i == 0 else 5
            for key in judgment_keys[:num_judgment]:
                val = mech.get(key, '/') if mech else '/'
                if val is not None and str(val).strip() not in ('/', '0', '', 'None'):
                    try:
                        float(str(val))
                        mech_values.append('合格')
                    except ValueError:
                        mech_values.append(str(val))
                else:
                    mech_values.append('/')

        # Deep clone and replace placeholders inline
        new_elements = []
        idx_counter = [0]
        for elem in template_elements:
            cloned = copy.deepcopy(elem)
            new_elements.append(cloned)

            for t_elem in cloned.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, ctr=idx_counter):
                    val_idx = ctr[0]
                    ctr[0] += 1
                    if val_idx < len(mech_values):
                        return mech_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

        # Insert after last group
        ref_elem = current_children[last_end]
        for elem in reversed(new_elements):
            ref_elem.addnext(elem)

        # Trim the newly inserted table to match batch size
        if len(models_batch) < 4:
            self._build_index()
            new_groups = self._find_mech_detail_groups()
            if new_groups:
                start, end = new_groups[-1]
                new_children = list(self.body)
                for k in range(start, end + 1):
                    child = new_children[k]
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if tag == 'tbl':
                        self._trim_mech_table_rows(k, len(models_batch))
                        break

        model_names = ', '.join(p.get('产品型号', '?') for p in models_batch)
        print(f'Cloned mechanical detail group for models: {model_names}')

    # =================================================================
    # Viscous damper: certificate adjustment
    # =================================================================

    def _clone_vd_cert_group(self, product_data, project_info):
        """Clone last cert group for viscous damper and fill with data."""
        groups = self._find_cert_groups()
        if not groups:
            return

        last_start, last_end = groups[-1]
        current_children = list(self.body)
        template_elements = current_children[last_start:last_end + 1]

        # Cert table field order: 规格型号, 生产日期, 出厂编号,
        # 检验员, 检验依据, 检验部门, 检验结果
        cert_values = [
            str(product_data.get('产品型号', '')),
            str(product_data.get('生产日期', '')),
            str(product_data.get('阻尼器编号范围', '')),
            str(project_info.get('检验员') or ''),
            str(product_data.get('检验标准', '')),
            '质量部',
            '合格',
        ]

        new_elements = []
        idx_counter = [0]
        for elem in template_elements:
            cloned = copy.deepcopy(elem)
            new_elements.append(cloned)

            for t_elem in cloned.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, ctr=idx_counter):
                    val_idx = ctr[0]
                    ctr[0] += 1
                    if val_idx < len(cert_values):
                        return cert_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

        ref_elem = current_children[last_end]
        for elem in reversed(new_elements):
            ref_elem.addnext(elem)

        print(f'Cloned VD cert group for: '
              f'{product_data.get("产品型号", "?")}')

    # =================================================================
    # Viscous damper: mechanical detail adjustment
    # =================================================================

    def _find_vd_mech_groups(self):
        """Find viscous damper mechanical report groups.
        Pattern: '成品力学性能检测报告' paragraph → Table (mechanical test)."""
        mech_text = '成品力学性能检测报告'
        mech_refs = self._para_by_text.get(mech_text, [])
        if not mech_refs:
            return []

        groups = []
        current_children = list(self.body)

        for body_idx, para_elem in mech_refs:
            # Find the table following this paragraph
            tbl_body_idx = None
            for j in range(body_idx + 1, len(current_children)):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'tbl':
                    tbl_body_idx = j
                    break
                elif tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    if text.strip() and '力学' not in text:
                        break
            if tbl_body_idx is None:
                continue

            # Include empty paragraphs between title and table
            start_idx = body_idx
            for j in range(body_idx - 1, max(-1, body_idx - 5), -1):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    if not text.strip():
                        start_idx = j
                    else:
                        break
                else:
                    break

            groups.append((start_idx, tbl_body_idx))

        return groups

    def _fill_vd_cert_group_at(self, start, end, product_data, project_info):
        """Fill placeholders in an existing cert group (no clone)."""
        current_children = list(self.body)
        cert_values = [
            str(product_data.get('产品型号', '')),
            str(product_data.get('生产日期', '')),
            str(product_data.get('阻尼器编号范围', '')),
            str(project_info.get('检验员') or ''),
            str(product_data.get('检验标准', '')),
            '质量部',
            '合格',
        ]
        idx_counter = [0]
        for i in range(start, end + 1):
            elem = current_children[i]
            for t_elem in elem.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, ctr=idx_counter):
                    val_idx = ctr[0]
                    ctr[0] += 1
                    if val_idx < len(cert_values):
                        return cert_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

    def _fill_vd_mech_group_at(self, start, end, product_data, project_info, mech_data):
        """Fill placeholders in an existing mechanical group (no clone).

        Note: Table 9 has merged cells for 检测结果+判定 columns, so the
        prepare script produces 12 placeholders (not 16). The merged cell
        uses the derived result ('符合'/'不符合') from the verdict.
        """
        current_children = list(self.body)

        def _result(verdict):
            return '符合' if verdict == '合格' else '不符合'

        mech_data = mech_data or {}
        test_keys = [
            '设计阻尼力(合格/不合格)',
            '极限位移(合格/不合格)',
            '阻尼系数(合格/不合格)',
            '阻尼指数(合格/不合格)',
        ]
        mech_values = [
            str(project_info.get('项目名称', '')),
            str(product_data.get('产品型号', '')),
            str(mech_data.get('最大阻尼力(kN)', '')),
            str(mech_data.get('设计位移(mm)', '')),
            str(mech_data.get('极限位移(mm)', '')),
            str(mech_data.get('阻尼系数(kN(s/mm)α)', '')),
            str(mech_data.get('阻尼力指数(α)', '')),
            str(mech_data.get('结构基频(Hz)', '')),
        ]
        for key in test_keys:
            verdict = str(mech_data.get(key, '合格'))
            mech_values.append(_result(verdict))  # merged cell: show result
        idx_counter = [0]
        for i in range(start, end + 1):
            elem = current_children[i]
            for t_elem in elem.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, ctr=idx_counter):
                    val_idx = ctr[0]
                    ctr[0] += 1
                    if val_idx < len(mech_values):
                        return mech_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

    def _remove_vd_mech_groups(self, count):
        """Remove the last `count` viscous damper mechanical groups."""
        groups = self._find_vd_mech_groups()
        if count <= 0 or count > len(groups):
            return

        to_remove = groups[-count:]
        current_children = list(self.body)
        removed = set()
        for start, end in to_remove:
            for i in range(start, end + 1):
                removed.add(current_children[i])
        for elem in removed:
            self.body.remove(elem)
        self._build_index()
        print(f'Removed {count} VD mechanical group(s), '
              f'{len(groups) - count} remaining.')

    def _clone_vd_mech_group(self, product_data, project_info, mech_data=None):
        """Clone last mechanical group for viscous damper and fill with data."""
        groups = self._find_vd_mech_groups()
        if not groups:
            return

        last_start, last_end = groups[-1]
        current_children = list(self.body)
        template_elements = current_children[last_start:last_end + 1]

        model_name = str(product_data.get('产品型号', ''))
        project_name = str(project_info.get('项目名称', ''))

        def _result(verdict):
            return '符合' if verdict == '合格' else '不符合'

        mech = mech_data or {}
        test_keys = [
            '设计阻尼力(合格/不合格)',
            '极限位移(合格/不合格)',
            '阻尼系数(合格/不合格)',
            '阻尼指数(合格/不合格)',
        ]
        mech_values = [
            project_name,
            model_name,
            str(mech.get('最大阻尼力(kN)', '')),
            str(mech.get('设计位移(mm)', '')),
            str(mech.get('极限位移(mm)', '')),
            str(mech.get('阻尼系数(kN(s/mm)α)', '')),
            str(mech.get('阻尼力指数(α)', '')),
            str(mech.get('结构基频(Hz)', '')),
        ]
        for key in test_keys:
            verdict = str(mech.get(key, '合格'))
            mech_values.append(_result(verdict))  # merged cell: show result

        new_elements = []
        idx_counter = [0]
        for elem in template_elements:
            cloned = copy.deepcopy(elem)
            new_elements.append(cloned)

            for t_elem in cloned.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, ctr=idx_counter):
                    val_idx = ctr[0]
                    ctr[0] += 1
                    if val_idx < len(mech_values):
                        return mech_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

        ref_elem = current_children[last_end]
        for elem in reversed(new_elements):
            ref_elem.addnext(elem)

        print(f'Cloned VD mechanical group for model: {model_name}')

    # =================================================================
    # Viscous damper: visual inspection adjustment
    # =================================================================

    def _find_vd_visual_groups(self):
        """Find viscous damper visual report groups.
        Pattern: company name -> title -> Table -> signature."""
        visual_text = '建筑消能阻尼器外观质量及尺寸检测报告'
        visual_refs = self._para_by_text.get(visual_text, [])
        if not visual_refs:
            return []

        groups = []
        current_children = list(self.body)

        for body_idx, para_elem in visual_refs:
            tbl_body_idx = None
            for j in range(body_idx + 1, min(len(current_children), body_idx + 4)):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'tbl':
                    tbl_body_idx = j
                    break
                elif tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    if text.strip():
                        break
            if tbl_body_idx is None:
                continue

            sig_idx = None
            for j in range(tbl_body_idx + 1, min(len(current_children), tbl_body_idx + 5)):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    if '检验员' in text or '审核' in text:
                        sig_idx = j
                        break
                    elif text.strip():
                        break
            end_idx = sig_idx if sig_idx is not None else tbl_body_idx

            start_idx = body_idx
            for j in range(body_idx - 1, max(-1, body_idx - 5), -1):
                child = current_children[j]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'p':
                    text = ''
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            text += t.text
                    if '四川融海运通' in text or 'FIELD' in text:
                        start_idx = j
                        for k in range(j - 1, max(-1, j - 3), -1):
                            child2 = current_children[k]
                            tag2 = child2.tag.split('}')[-1] if '}' in child2.tag else child2.tag
                            if tag2 == 'p':
                                text2 = ''
                                for t in child2.iter(qn('w:t')):
                                    if t.text:
                                        text2 += t.text
                                if not text2.strip():
                                    start_idx = k
                                else:
                                    break
                            else:
                                break
                        break

            groups.append((start_idx, end_idx))

        return groups

    def _remove_vd_visual_groups(self, count):
        """Remove the last `count` viscous damper visual groups."""
        groups = self._find_vd_visual_groups()
        if count <= 0 or count > len(groups):
            return

        to_remove = groups[-count:]
        current_children = list(self.body)
        removed = set()
        for start, end in to_remove:
            for i in range(start, end + 1):
                removed.add(current_children[i])
        for elem in removed:
            self.body.remove(elem)
        self._build_index()
        print(f'Removed {count} VD visual group(s), '
              f'{len(groups) - count} remaining.')

    def _vd_visual_values(self, product_data, project_info, vis_data):
        """Build the list of values to fill VD visual group placeholders."""
        def _measured(item_name, verdict):
            VD_MEASURED_MAP = {
                '表面平滑度': {'合格': '光滑平整', '不合格': '不平整'},
                '机械损伤': {'合格': '无', '不合格': '有'},
                '锈蚀毛刺': {'合格': '无', '不合格': '有'},
                '无渗漏': {'合格': '无', '不合格': '有渗漏'},
                '产品标识': {'合格': '标识清晰易辨认', '不合格': '标识不清晰'},
                '长度偏差': {'合格': '符合标准要求', '不合格': '超出允许偏差'},
                '截面尺寸偏差': {'合格': '符合标准要求', '不合格': '超出允许偏差'},
            }
            desc = VD_MEASURED_MAP.get(item_name, {})
            return desc.get(verdict, verdict)

        vis_data = vis_data or {}
        items = ['表面平滑度', '机械损伤', '锈蚀毛刺', '无渗漏',
                 '产品标识', '长度偏差', '截面尺寸偏差']

        values = [
            str(project_info.get('产品供应商', '')),
            str(project_info.get('项目名称', '')),
            str(product_data.get('产品型号', '')),
        ]
        for item_name in items:
            key = f'{item_name}(合格/不合格)'
            verdict = str(vis_data.get(key, '合格'))
            values.append(_measured(item_name, verdict))
            values.append(str(product_data.get('数量', '')))
            values.append(verdict)

        values.append(str(product_data.get('产品型号', '')))
        values.append('该项目黏滞阻尼器')
        values.append('质量部')

        # Signature paragraph fields (project-level, same for all models)
        values.append(str(project_info.get('检验员') or ''))
        values.append(str(project_info.get('审核') or ''))
        values.append(str(project_info.get('制造日期') or ''))

        return values

    def _fill_vd_visual_group_at(self, start, end, product_data, project_info, vis_data):
        """Fill placeholders in an existing visual group (no clone)."""
        current_children = list(self.body)
        vis_values = self._vd_visual_values(product_data, project_info, vis_data)
        idx_counter = [0]
        for i in range(start, end + 1):
            elem = current_children[i]
            for t_elem in elem.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, ctr=idx_counter):
                    val_idx = ctr[0]
                    ctr[0] += 1
                    if val_idx < len(vis_values):
                        return vis_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

    def _clone_vd_visual_group(self, product_data, project_info, vis_data=None):
        """Clone last visual group for viscous damper and fill with data."""
        groups = self._find_vd_visual_groups()
        if not groups:
            return

        last_start, last_end = groups[-1]
        current_children = list(self.body)
        template_elements = current_children[last_start:last_end + 1]

        vis_values = self._vd_visual_values(product_data, project_info, vis_data)

        new_elements = []
        idx_counter = [0]
        for elem in template_elements:
            cloned = copy.deepcopy(elem)
            new_elements.append(cloned)

            for t_elem in cloned.iter(qn('w:t')):
                if not t_elem.text:
                    continue

                def replace_with_value(m, ctr=idx_counter):
                    val_idx = ctr[0]
                    ctr[0] += 1
                    if val_idx < len(vis_values):
                        return vis_values[val_idx]
                    return ''

                t_elem.text = re.sub(
                    r'\{\{FIELD_\d+\}\}',
                    replace_with_value,
                    t_elem.text
                )

        ref_elem = current_children[last_end]
        for elem in reversed(new_elements):
            ref_elem.addnext(elem)

        print(f'Cloned VD visual group for model: '
              f'{product_data.get("产品型号", "?")}')

    # =================================================================
    # Master adjustment
    # =================================================================

    def adjust(self, data, max_in_template=6, product_type='isolation_bearing'):
        """Run all adjustments. data = result from excel_reader.read_excel_data()"""
        products = data.get('product_list', [])
        project_info = data.get('project_info', {})
        project_name = project_info.get('项目名称', '')
        model_count = len(products)

        if product_type == 'viscous_damper':
            self._adjust_viscous_damper(data, products, project_info, model_count)
        else:
            self._adjust_isolation_bearing(data, products, project_name, model_count)

    def _adjust_viscous_damper(self, data, products, project_info, model_count):
        """Adjust template for viscous damper product."""
        # --- Certificates ---
        cert_groups = self._find_cert_groups()
        current_certs = len(cert_groups)

        if model_count < current_certs:
            self._remove_cert_groups(current_certs - model_count)
        elif model_count > current_certs:
            for i in range(current_certs, min(model_count, 20)):
                self._clone_vd_cert_group(products[i], project_info)
            self._build_index()

        # Fill ALL cert groups' placeholders (both remaining original and cloned)
        self._build_index()
        cert_groups = self._find_cert_groups()
        for i, (start, end) in enumerate(cert_groups):
            if i < model_count:
                self._fill_vd_cert_group_at(start, end, products[i], project_info)

        # --- Visual reports (1 per model) ---
        # IMPORTANT: Re-find groups after cert modifications (body indices shifted)
        self._build_index()
        vis_groups = self._find_vd_visual_groups()
        current_vis = len(vis_groups)
        needed_vis = max(1, model_count)

        vis_by_model = {v['产品型号']: v for v in data.get('visual_data', [])}

        if needed_vis < current_vis:
            self._remove_vd_visual_groups(current_vis - needed_vis)
        elif needed_vis > current_vis:
            for i in range(current_vis, needed_vis):
                vis = vis_by_model.get(products[i].get('产品型号', ''))
                self._clone_vd_visual_group(products[i], project_info, vis)

        # Fill ALL vis groups' placeholders (both remaining original and cloned)
        self._build_index()
        vis_groups = self._find_vd_visual_groups()
        for i, (start, end) in enumerate(vis_groups):
            if i < model_count:
                vis = vis_by_model.get(products[i].get('产品型号', ''))
                self._fill_vd_visual_group_at(start, end, products[i], project_info, vis)

        # --- Mechanical detail (1 per model) ---
        # IMPORTANT: Re-find groups after cert modifications (body indices shifted)
        self._build_index()
        mech_groups = self._find_vd_mech_groups()
        current_mech = len(mech_groups)
        needed_mech = max(1, model_count)

        mech_by_model = {m['产品型号']: m for m in data.get('mechanical_data', [])}

        if needed_mech < current_mech:
            self._remove_vd_mech_groups(current_mech - needed_mech)
        elif needed_mech > current_mech:
            for i in range(current_mech, needed_mech):
                mech = mech_by_model.get(products[i].get('产品型号', ''))
                self._clone_vd_mech_group(products[i], project_info, mech)

        # Fill ALL mech groups' placeholders (both remaining original and cloned)
        self._build_index()
        mech_groups = self._find_vd_mech_groups()
        for i, (start, end) in enumerate(mech_groups):
            if i < model_count:
                mech = mech_by_model.get(products[i].get('产品型号', ''))
                self._fill_vd_mech_group_at(start, end, products[i], project_info, mech)

        print(f'Adjusted VD to {model_count} models '
              f'(certs: {current_certs}->{min(model_count, 20)}, '
              f'visual: {current_vis}->{needed_vis}, '
              f'mech: {current_mech}->{needed_mech}).')

    def _adjust_isolation_bearing(self, data, products, project_name, model_count):
        """Adjust template for isolation bearing product (original logic)."""
        # --- Certificates ---
        cert_groups = self._find_cert_groups()
        current_certs = len(cert_groups)

        if model_count < current_certs:
            self._remove_cert_groups(current_certs - model_count)
        elif model_count > current_certs:
            for i in range(current_certs, min(model_count, 12)):
                self._clone_cert_group(products[i])
            self._build_index()

        # --- Visual reports ---
        vis_groups = self._find_visual_groups()
        current_vis = len(vis_groups)

        if model_count < current_vis:
            self._remove_visual_groups(current_vis - model_count)
        elif model_count > current_vis:
            for i in range(current_vis, min(model_count, 12)):
                self._clone_visual_group(products[i], project_name)

        # --- Mechanical detail tables ---
        MECH_PER_PAGE = 4
        mech_groups = self._find_mech_detail_groups()
        current_mech = len(mech_groups)
        needed_mech = max(1, (model_count + MECH_PER_PAGE - 1) // MECH_PER_PAGE)

        mech_by_model = {m['产品型号']: m for m in data.get('mechanical_data', [])}

        if needed_mech < current_mech:
            self._remove_mech_detail_groups(current_mech - needed_mech)
        elif needed_mech > current_mech:
            for i in range(current_mech, needed_mech):
                start_model = i * MECH_PER_PAGE
                end_model = min(start_model + MECH_PER_PAGE, model_count)
                batch = products[start_model:end_model]
                self._clone_mech_detail_group(batch, project_name, mech_by_model)

        # Trim the last group's table to match remaining model count
        self._build_index()
        final_mech = self._find_mech_detail_groups()
        if final_mech:
            start, end = final_mech[-1]
            current_children = list(self.body)
            for k in range(start, end + 1):
                child = current_children[k]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'tbl':
                    last_page_models = model_count - (needed_mech - 1) * MECH_PER_PAGE
                    if last_page_models < MECH_PER_PAGE:
                        self._trim_mech_table_rows(k, last_page_models)
                    break

        print(f'Adjusted to {model_count} models '
              f'(certs: {current_certs}→{min(model_count,12)}, '
              f'visual: {current_vis}→{min(model_count,12)}, '
              f'mech: {current_mech}→{needed_mech}).')
