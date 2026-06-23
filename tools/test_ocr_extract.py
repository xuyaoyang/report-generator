#!/usr/bin/env python
"""Quick test of improved OCR extractors."""

import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Copy the new extractor functions here temporarily
def extract_cert_date(text):
    lines = text.split('\n')
    for pat_label, pat_date in [
        (r'发货日期|DATE\s*TO\s*PULL', r'[：:]\s*(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})'),
        (r'出厂日期', r'[：:]\s*(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})'),
        (r'(?:制造日期|签发日期|检测日期)', r'[：:]\s*(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})'),
    ]:
        for i, line in enumerate(lines):
            if re.search(pat_label, line, re.IGNORECASE):
                m = re.search(pat_date, line, re.IGNORECASE)
                if m:
                    d_clean = re.sub(r'[-/\.]', '', m.group(1))
                    if len(d_clean) >= 8:
                        return d_clean[:6]
                for j in range(i+1, min(i+3, len(lines))):
                    nl = lines[j].strip()
                    m2 = re.search(r'(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})', nl)
                    if m2:
                        try:
                            return f'{int(m2.group(1)):04d}{int(m2.group(2)):02d}'
                        except ValueError:
                            pass
                    m3 = re.search(r'(\d{8,14})', nl)
                    if m3 and m3.group(1).startswith('20'):
                        return m3.group(1)[:6]
    for i, line in enumerate(lines):
        if re.search(r'(?:日期|DATE)\s*[：:]?', line, re.IGNORECASE):
            for j in range(i, min(i+3, len(lines))):
                t = lines[j].strip()
                m = re.search(r'(\d{4})[-/\.年](\d{1,2})[-/\.月](\d{1,2})', t)
                if m:
                    try:
                        y, mo = int(m.group(1)), int(m.group(2))
                        if 2020 <= y <= 2026 and 1 <= mo <= 12:
                            return f'{y:04d}{mo:02d}'
                    except ValueError:
                        pass
                m2 = re.search(r'(\d{8,14})', t)
                if m2 and m2.group(1).startswith('20'):
                    return m2.group(1)[:6]
    return None


def extract_heat_batch_no(text, category):
    """Extract batch/heat number from OCR text.

    重钢 format: PLATE NO | BATCH No | HEAT No | DIMENSION
      Data: 6P23729353200  13180659RLL3500  23106654  20.00*2000*10500
      Batch = column after PLATE NO (pattern: digits+letter+digits, 10+ chars)

    攀钢 format: 序号 | 熔炼号 | 钢卷号 | 检验批号
      Data: X21106932  X11015105000  X11015105000  Q235B
      Batch = the 检验批号 column (same value as 钢卷号 in this case)
      Pattern: X + 11 digits
    """
    tokens = text.split()
    n = len(tokens)

    if category == '钢板':
        # Method A: scan consecutive pairs matching "6P2372... 13180659RLL3500"
        for ti in range(n - 1):
            a = tokens[ti]
            b = tokens[ti+1]
            if re.match(r'^\d+[A-Z]\d{6,}', a) and len(a) >= 10:
                if re.match(r'^\d{6,}[A-Z]{2,4}\d{0,}', b) and len(b) >= 10:
                    return b

        # Method B: 攀钢 — header: 序号 熔炼号 钢卷号 检验批号 ...
        #   熔炼号=Heat No, 钢卷号=Coil No, 检验批号=Check/Batch No
        #   检验批号 is column 3 (0-indexed) after 序号. Data has X-prefix codes.
        for ti, tok in enumerate(tokens):
            if re.search(r'检验批号|Check\s*No', tok, re.IGNORECASE):
                # Found the header token. Look for first X-code AFTER this position
                # as the batch number (it's in the 检验批号 column).
                for dj in range(ti+1, min(ti+30, n)):
                    if re.match(r'^X\d{10,}$', tokens[dj]):
                        # Make sure it's not the FIRST X-code (that's 熔炼号)
                        # Actually: the first X-code is熔炼号, 2nd is钢卷号, 3rd is检验批号
                        # Better: collect all X-codes starting from header, take the 3rd one
                        x_codes = []
                        for dk in range(ti+1, min(ti+30, n)):
                            if re.match(r'^X\d{10,}$', tokens[dk]):
                                x_codes.append(tokens[dk])
                            if len(x_codes) >= 3:
                                return x_codes[2]  # 3rd X-code = 检验批号
                        if x_codes:
                            return x_codes[-1]  # fallback: last one
                        break
                break

        # Method C: general batch pattern in close proximity to header
        for ti, tok in enumerate(tokens):
            if re.search(r'BATCH|批号|检验批号', tok, re.IGNORECASE):
                for dj in range(ti+1, min(ti+10, n)):
                    cand = tokens[dj]
                    if re.match(r'^[A-Z0-9]{6,25}$', cand) and re.search(r'[A-Z]', cand):
                        if not re.search(r'^(?:Size|Weight|Chemica|PRODUCT)', cand, re.IGNORECASE):
                            return cand

        # Method D: fallback — batch no pattern in full text
        m = re.search(r'\b(\d{8,14}[A-Z]{2,4}\d{0,6})\b', text)
        if m:
            return m.group(1)

        # Method E: X-prefix codes (攀钢 batch numbers)
        m = re.search(r'\b(X\d{10,})\b', text)
        if m:
            # Return the first X-code that appears AFTER '检验批号' header position
            header_match = re.search(r'检验批号', text)
            if header_match:
                for m2 in re.finditer(r'\b(X\d{10,})\b', text[header_match.end():]):
                    return m2.group(1)
            return m.group(1)

    # Non-钢板
    for pat in [
        r'(?:炉号|熔炼号|HEAT\s*No)[：:]?\s*(\S{4,30})',
        r'(?:批号|组批号|产品批号|BATCH\s*No)[：:]?\s*(\S{4,30})',
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if 3 <= len(val) <= 25:
                return val
    return None


if __name__ == '__main__':
    from rapidocr_onnxruntime import RapidOCR
    engine = RapidOCR()

    imgs = [
        'image_lib/钢板/00aa78e9_20厚钢板235B质保书-生产日期20211014.png',
        'image_lib/钢板/01984947_Q235B 20 2000.png',
        'image_lib/钢板/6247d89f_2211MD09875_6820b65e-39b5-4f03-abaf-50aadb3751d6.png',
        'image_lib/钢板/50b80986_8 1.51.png',
        'image_lib/钢板/7636279a_质保书-2024.11.05(2)(1)(1).png',
        'image_lib/钢筋/00aa78e9_...' if False else 'image_lib/钢板/0d82b982_重钢普 16-2米.png',
    ]
    for img_path in imgs:
        if not os.path.exists(img_path):
            print(f'MISSING: {img_path}')
            continue
        results, _ = engine(img_path)
        text = '\n'.join(t for _, t, _ in results)
        heat = extract_heat_batch_no(text, '钢板')
        date = extract_cert_date(text)
        print(f'[{os.path.basename(img_path)[:55]}]')
        print(f'  heat: {heat}')
        print(f'  date: {date}')
        print()
