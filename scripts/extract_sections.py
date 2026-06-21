"""
Extract individual section templates from template_prepared.docx.

Strategy: Load the full template, save a copy, then remove all body elements
EXCEPT those in the desired range. This preserves ALL formatting (styles,
numbering, headers/footers) since we're working within the same document.

Produces: templates/*.docx with {{FIELD_XXX}} placeholders intact.
"""
import os, sys, shutil
from docx import Document
from docx.oxml.ns import qn
from lxml import etree


def remove_element(el):
    el.getparent().remove(el)


def extract_section(source_path, output_path, keep_range):
    """Extract a body element range from source to a new file.
    keep_range: (start_idx, end_idx) inclusive indices in body children.
    """
    doc = Document(source_path)
    body = doc.element.body
    children = list(body)

    start, end = keep_range
    # Remove elements AFTER the range first (indices stay stable)
    for i in range(len(children) - 1, end, -1):
        remove_element(children[i])
    # Remove elements BEFORE the range
    children = list(body)  # refresh after removals
    for i in range(start - 1, -1, -1):
        remove_element(children[i])

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)

    # Count remaining elements
    remaining = len(list(body))
    print(f'  {os.path.basename(output_path)}: [{start}-{end}] → {remaining} elements')


def extract_ib():
    """Extract IB section templates.

    Body structure (160 elements, 15 tables):
      [0-17]:    Cover page
      [18-31]:   TOC (paragraphs only)
      [32-51]:   Empty/business license area
      [52-53]:   Cert 1 (paragraph + Table 0, 9r x 2c)
      [54-77]:   Cert 2-6 (5 more certs)
      [78-80]:   Empty
      [81-82]:   Mech summary (company para + Table 6, 17r x 13c)
      [83-85]:   Mech title + notes (part of mech summary page)
      [86]:      Empty
      [87-91]:   Mech detail 1 (company + title + Table 7 + notes)
      [92-96]:   Mech detail 2 (company + title + Table 8 + notes)
      [97-101]:  Visual 1 (company + title + Table 9 + signature)
      [102-132]: Visual 2-6
      [133-135]: Empty
      [136-159]: Quality commitment (paragraphs only)
    """
    base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        'products', 'isolation_bearing')
    source = os.path.join(base, 'template_prepared.docx')
    out = os.path.join(base, 'templates')

    print('=== IB sections ===')
    # 01_cover: elements 0-17
    extract_section(source, os.path.join(out, '01_cover.docx'), (0, 17))
    # 02_toc: elements 18-31 (paragraphs before certs)
    extract_section(source, os.path.join(out, '02_toc.docx'), (18, 31))
    # 03_business_license: elements 32-51 (contains image)
    extract_section(source, os.path.join(out, '03_business_license.docx'), (32, 51))
    # 04_cert: elements 52-53 (single cert = paragraph + table)
    extract_section(source, os.path.join(out, '04_cert.docx'), (52, 53))
    # 05_mech_summary: elements 81-85 (company + summary table + title + notes)
    extract_section(source, os.path.join(out, '05_mech_summary.docx'), (81, 85))
    # 06_mech_detail: elements 86-91 (empty + company + title + table + notes)
    extract_section(source, os.path.join(out, '06_mech_detail.docx'), (86, 91))
    # 07_visual: elements 97-101 (company + title + table + signature)
    extract_section(source, os.path.join(out, '07_visual.docx'), (97, 101))
    # 08_quality_commitment: elements 136-159
    extract_section(source, os.path.join(out, '08_quality_commitment.docx'), (136, 159))


def extract_vd():
    """Extract VD section templates.

    Body structure (162 elements, 10 tables):
      [0-33]:    Cover
      [34-49]:   TOC
      [50-55]:   Company intro
      [56-59]:   Business license (image)
      [60-120]:  Certs ×8 (paragraphs + Table 0-7)
      [121-142]: Quality commitment
      [143-149]: Visual report (company + title + Table 8 + signature)
      [150-152]: Mechanical report (title + Table 9)
      [153-161]: Material cert placeholder pages (static)
    """
    base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        'products', 'viscous_damper')
    source = os.path.join(base, 'template_prepared.docx')
    out = os.path.join(base, 'templates')

    print('\n=== VD sections ===')
    extract_section(source, os.path.join(out, '01_cover.docx'), (0, 33))
    extract_section(source, os.path.join(out, '02_toc.docx'), (34, 49))
    extract_section(source, os.path.join(out, '03_company_intro.docx'), (50, 55))
    extract_section(source, os.path.join(out, '04_business_license.docx'), (56, 59))
    # Single cert: paragraph "产品合格证" + table (elements 61 + 64-65)
    # Actually cert 1 = [60] section-title + [61] para + [62-63] empty + [64] table
    extract_section(source, os.path.join(out, '05_cert.docx'), (60, 64))
    extract_section(source, os.path.join(out, '06_quality_commitment.docx'), (121, 142))
    extract_section(source, os.path.join(out, '07_visual.docx'), (143, 149))
    extract_section(source, os.path.join(out, '08_mechanical.docx'), (150, 152))


if __name__ == '__main__':
    extract_ib()
    extract_vd()
    print('\nDone. Review templates/ directories.')
