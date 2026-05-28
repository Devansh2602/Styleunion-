import re
import pdfplumber


VALID_SIZES = {'XS', 'S', 'M', 'L', 'XL', 'XXL', '2XL', '3XL'}


def _format_date_dots(date_str):
    """Convert DD-MM-YYYY to DD.MM.YYYY"""
    return date_str.replace('-', '.') if date_str else ''


def _extract_color_size_desc(prefix, line2):
    """
    Extract color, size, description from:
      prefix  - text between style code and HSN on line 1
      line2   - full text of the second line of the item

    Handles 4 PDF formats:
      A)  prefix="POWDER BLUE",        line2="XS SU-YWET00080-126 SS26"
      A2) prefix="BOMBAY",             line2="BROWN XL SU-EEST00422-126 SS26"
      B)  prefix="BROWN L SU-",        line2="ES00434-126 SS26"          (desc split, hyphen)
      C)  prefix="SHEEPSKIN XXL",      line2="SU-ESPT00003-126 SS26"
    """
    tokens1 = prefix.strip().split()
    tokens2 = line2.strip().split()
    all_tokens = tokens1 + tokens2

    # Find first valid size token across both lines
    size_idx = None
    for i, t in enumerate(all_tokens):
        if t.upper() in VALID_SIZES:
            size_idx = i
            break

    if size_idx is None:
        return prefix.strip(), line2.strip(), ''

    color = ' '.join(all_tokens[:size_idx])
    size = all_tokens[size_idx].upper()
    n1 = len(tokens1)

    if size_idx < n1:
        # Size was on line 1; description starts after it
        after = tokens1[size_idx + 1:]
        if after and after[-1].endswith('-'):
            # Last word on line 1 ends with hyphen — join it directly to first token on line 2
            pre = ' '.join(after)
            if tokens2:
                desc = pre.rstrip() + tokens2[0]
                if len(tokens2) > 1:
                    desc += ' ' + ' '.join(tokens2[1:])
            else:
                desc = pre
        elif after:
            desc = ' '.join(after)
            if tokens2:
                desc += ' ' + ' '.join(tokens2)
        else:
            desc = ' '.join(tokens2)
    else:
        # Size was on line 2; description is everything after it
        desc = ' '.join(all_tokens[size_idx + 1:])

    return color.strip(), size, desc.strip()


def parse_po(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ''
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + '\n'

    result = {}

    # PO number
    m = re.search(r'Order No\.\s*:\s*([A-Z0-9()/\-]+)', full_text)
    result['po_number'] = m.group(1).strip() if m else ''

    # PO date (first DD-MM-YYYY occurrence near "Date :")
    m = re.search(r'Date\s*:\s*(\d{2}-\d{2}-\d{4})', full_text)
    result['po_date'] = _format_date_dots(m.group(1)) if m else ''

    # Consignee GST (format: GSTIN :XXXXX[...] or GSTIN :- XXXXX)
    m = re.search(r'GSTIN\s*[:-]+\s*([A-Z0-9]{15})', full_text)
    result['consignee_gst'] = m.group(1) if m else '36AAQCS3643G1Z2'

    # Consignee name (first occurrence of "Nexon Omniverse")
    m = re.search(r'(Nexon Omniverse Limited)', full_text)
    result['consignee_name'] = m.group(1).strip() if m else 'Nexon Omniverse Limited'

    # Consignee address lines (AWL India block)
    m = re.search(
        r'(AWL INDIA PVT LTD.*?)\n(.*?)\n(.*?(?:Malkajgiri|Medchal))\n(.*?500078)',
        full_text, re.DOTALL
    )
    if m:
        result['consignee_address_lines'] = [
            m.group(1).strip(),
            (m.group(2).strip() + ' ' + m.group(3).strip()).strip(),
            m.group(4).strip(),
            'Telangana,INDIA',
        ]
    else:
        result['consignee_address_lines'] = [
            'AWL INDIA PVT LTD, SV NO 705/2, 706/4 DEVARYAMJAL, Opp.',
            'Sai Geetha Ashramam, Tumkunta Municipality, Medchal Malkajgiri',
            'Hyderabad, 500078',
            'Telangana,INDIA',
        ]

    # IGST rate
    m = re.search(r'IGST\s*\[?\s*@\s*(\d+)%', full_text)
    result['igst_rate'] = float(m.group(1)) / 100 if m else 0.05

    # Line items
    # Each item spans two PDF lines:
    # Line 1: {barcode} STYLE UNION {style} {color/size/desc_prefix} {HSN} {rate}.000{mrp}.00 {qty}.000 PCS {amount}
    # Line 2: {remainder of color/size/description}
    #
    # The text between {style} and {HSN} (captured as group 3) may include just the color,
    # or color+size, or color+size+start-of-description (if PDF wrapped the description).
    item_re = re.compile(
        r'(\d{13})\s+STYLE\s+UNION\s+(\w+)\s+(.+?)\s+(\d{7,8})\s+'
        r'(\d+\.\d{3})(\d+\.\d{2})\s+'     # rate (3dp) glued to mrp (2dp)
        r'([\d,]+\.\d{3})\s+PCS\s+'         # qty
        r'([\d,]+\.\d{2})\n'                # amount
        r'(.+)',                             # entire line 2
        re.MULTILINE,
    )

    items = []
    for m in item_re.finditer(full_text):
        barcode, style, prefix, hsn, rate, mrp, qty_str, amount_str, line2 = m.groups()
        color, size, desc = _extract_color_size_desc(prefix, line2)
        items.append({
            'barcode': barcode.strip(),
            'style':   style.strip(),
            'color':   color,
            'hsn':     hsn.strip(),
            'rate':    float(rate),
            'mrp':     float(mrp),
            'qty':     int(float(qty_str.replace(',', ''))),
            'amount':  float(amount_str.replace(',', '')),
            'size':    size,
            'description': desc,
        })

    result['items'] = items
    return result
