import os
import re
import openpyxl

# Configuration Paths
REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
PORTS_DIR = os.path.join(REPO_ROOT, "ports")
CSS_FILE = os.path.join(REPO_ROOT, "assets", "css", "port-pages.css")
EXCEL_PATH = r"D:\FalconXoft Ship Chandler\Website Data\Ship_Chandler_SEO_49_Pages_FULL.xlsx"

# Section heading mapping: (id, callout_wrapper_class)
H2_MAP = [
    ('professional-services', None),
    ('provisions', None),
    ('bonded', None),
    ('deck', None),
    ('engine', None),
    ('cabin', None),
    ('safety', None),
    ('electrical', None),
    ('chemicals', None),
    ('welding', None),
    ('spare-parts', None),
    ('publications', None),
    ('tools', None),
    ('technical-support', None),
    ('urgent', 'callout-box callout-urgent'),
    ('vessels', None),
    ('rfq', 'callout-box callout-rfq'),
    ('why-choose', None),
    (None, None),
    (None, None),
    (None, None),
    (None, None),
]


def clean_port_name(name):
    """
    Strips generic words to get a clean, human-friendly port/city name.
    """
    cleaned = re.sub(r'(?i)\bship\s*chandlers?\b|\bchandlers?\b|\bship\s*supply\b', '', str(name)).strip(' -')
    return cleaned if cleaned else str(name)


def render_article_html(desc_text, h1, clean_name):
    """
    Renders semantic HTML for the port article matching the new Spain theme UI:
    - Overview card with gold top line and 3 lead paragraphs
    - Port operations showcase image (warehouse-about.webp)
    - Headings with proper TOC IDs
    - Modern items-grid for bullet lists with check icons
    - Structured callout boxes for Urgent Supply and How to Send RFQ
    - High-conversion Pre-Footer CTA block
    """
    blocks = [b.strip() for b in desc_text.split('\n\n') if b.strip()]
    lines = []

    # 1. Elevated Lead Overview Card (paragraphs 1, 2, 3)
    p1 = blocks[1] if len(blocks) > 1 else ""
    p2 = blocks[2] if len(blocks) > 2 else ""
    p3 = blocks[3] if len(blocks) > 3 else ""

    lines.append('          <!-- Elevated Lead Overview Box -->')
    lines.append('          <div class="intro-overview-card">')
    lines.append('            <span class="overview-pill"><i class="fa-solid fa-circle-info"></i> Port Chandling Overview</span>')
    lines.append(f'            <p>{p1}</p>')
    lines.append(f'            <p>{p2}</p>')
    lines.append(f'            <p>{p3}</p>')
    lines.append('          </div>')
    lines.append('')

    # 2. Featured Supply Operations Image Box
    lines.append('          <!-- Featured Supply Operations Image Box -->')
    lines.append('          <div class="article-image-box">')
    lines.append(f'            <img src="../assets/imgs/warehouse-about.webp" alt="Ship Supply Operations at {clean_name}">')
    lines.append('            <div class="image-tag-overlay">')
    lines.append('              <i class="fa-solid fa-boxes-packing"></i> Port Logistics & Vessel Chandling Operations')
    lines.append('            </div>')
    lines.append('          </div>')
    lines.append('')

    h2_idx = 0
    in_callout = False

    # 3. Main Article Body
    for b in blocks[4:]:
        if b.startswith('# '):
            continue
        elif b.startswith('## '):
            if in_callout:
                lines.append('          </div>')
                lines.append('')
                in_callout = False

            heading_text = b[3:].strip()
            cfg = H2_MAP[h2_idx] if h2_idx < len(H2_MAP) else (None, None)
            h2_idx += 1
            hid, callout_cls = cfg

            if callout_cls:
                in_callout = True
                lines.append(f'          <div class="{callout_cls}">')
                lines.append(f'            <h2 id="{hid}" style="border: none; margin: 0 0 14px; padding: 0;">{heading_text}</h2>')
            elif hid:
                lines.append(f'          <h2 id="{hid}">{heading_text}</h2>')
            else:
                lines.append(f'          <h2>{heading_text}</h2>')
        elif b.startswith('- '):
            style_attr = ' style="margin-bottom: 18px;"' if in_callout else ''
            lines.append(f'          <ul class="items-grid"{style_attr}>')
            for item in b.split('\n'):
                item = item.strip()
                if not item:
                    continue
                val = re.sub(r'^\s*-\s*', '', item)
                lines.append(f'            <li><i class="fa-solid fa-circle-check item-bullet-icon"></i> {val}</li>')
            lines.append('          </ul>')
        else:
            style_attr = ' style="margin-bottom: 0;"' if in_callout and b == blocks[-1] else ''
            if b.startswith('**') and b.endswith('**'):
                bold_content = b[2:-2].strip()
                lines.append(f'          <p><strong>{bold_content}</strong></p>')
            else:
                lines.append(f'          <p{style_attr}>{b}</p>')

    if in_callout:
        lines.append('          </div>')
        lines.append('')
        in_callout = False

    # 4. Elevated Pre-Footer Conversion Block
    lines.append('')
    lines.append('          <!-- Elevated Conversion Block -->')
    lines.append('          <div class="article-final-cta">')
    lines.append('            <div class="cta-inner-content">')
    lines.append(f'              <h3>Dispatch Supplies for Your Port Call in {clean_name}</h3>')
    lines.append('              <p>Connect with our technical marine procurement coordinators for competitive quotations, IMPA/ISSA code matching, and coordinated delivery straight to your vessel\'s berth or anchorage.</p>')
    lines.append('              <div class="cta-buttons-cluster">')
    lines.append('                <a href="../request-quote.html" class="btn btn-primary">Submit RFQ / Request Quote <i class="fa-solid fa-arrow-right"></i></a>')
    lines.append('                <a href="../contact.html" class="btn btn-outline">Contact Port Coordinators</a>')
    lines.append('              </div>')
    lines.append('            </div>')
    lines.append('          </div>')

    return '\n'.join(lines)


def get_template_parts():
    """
    Extracts the master HTML shell from bissau-ship-chandler-guinea.html
    which already contains the centralized CSS link, Spain navbar & footer.
    """
    bissau_path = os.path.join(PORTS_DIR, "bissau-ship-chandler-guinea.html")
    with open(bissau_path, 'r', encoding='utf-8') as f:
        master_html = f.read()

    art_start = '<article class="port-article">'
    art_end = '</article>'
    p1 = master_html.find(art_start) + len(art_start)
    p2 = master_html.find(art_end, p1)

    if p1 == -1 or p2 == -1:
        raise ValueError("Could not find article markers in bissau-ship-chandler-guinea.html")

    head_part = master_html[:p1]
    tail_part = master_html[p2:]
    return head_part, tail_part


def generate_page_html(head_part, tail_part, page_name, h1, meta_title, meta_desc, full_desc, slug):
    """
    Generates complete HTML for a port page using the master template.
    """
    clean_name = clean_port_name(page_name)
    new_keywords = f"{clean_name} ship chandler, {clean_name} vessel supply, ship supply {clean_name}, marine stores, provisions, bonded stores, deck stores, engine stores, marine procurement"

    h = head_part

    # Replace Title
    h = re.sub(r'<title>.*?</title>', f'<title>{meta_title}</title>', h)

    # Replace Meta Description
    h = re.sub(r'<meta name="description"\s+content=".*?">', f'<meta name="description"\n    content="{meta_desc}">', h, flags=re.DOTALL)

    # Replace Meta Keywords
    h = re.sub(r'<meta name="keywords"\s+content=".*?">', f'<meta name="keywords"\n    content="{new_keywords}">', h, flags=re.DOTALL)

    # Replace Open Graph Tags
    h = re.sub(r'<meta property="og:title"\s+content=".*?">', f'<meta property="og:title" content="{meta_title}">', h)
    h = re.sub(r'<meta property="og:description"\s+content=".*?">', f'<meta property="og:description"\n    content="{meta_desc}">', h, flags=re.DOTALL)

    # Replace Hero H1 and Description
    h = re.sub(r'<h1>.*?</h1>', f'<h1>{h1}</h1>', h)
    h = re.sub(r'<p class="port-hero-desc">.*?</p>', f'<p class="port-hero-desc">{meta_desc}</p>', h, flags=re.DOTALL)

    # Replace Sidebar CTA
    h = re.sub(r'<h4>Calling at .*?\?</h4>', f'<h4>Calling at {clean_name}?</h4>', h)

    # Update Mega Menu Active Link
    # Remove any existing active styling in mega menu
    h = re.sub(r'<a href="([^"]+)" style="color: var\(--sun\); font-weight: 700;">(.*?)</a>', r'<a href="\1">\2</a>', h)

    # Add active styling for current page's slug in mega menu
    target_link = f'<a href="{slug}">{page_name}</a>'
    active_target_link = f'<a href="{slug}" style="color: var(--sun); font-weight: 700;">{page_name}</a>'
    h = h.replace(target_link, active_target_link)

    # Render article content from Excel description
    art = '\n\n' + render_article_html(full_desc, h1, clean_name) + '\n\n        '

    return h + art + tail_part


def main():
    print(f"Reading Excel data from: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet = wb['SEO Pages']

    head_part, tail_part = get_template_parts()

    rows = list(sheet.iter_rows(values_only=True))[1:]  # Skip header row
    print(f"Found {len(rows)} port rows in Excel.")

    generated_count = 0
    generated_files = []

    for row in rows:
        num, page_name, h1, meta_title, meta_desc, full_desc = row[:6]
        if not page_name:
            continue

        slug = re.sub(r'[^a-z0-9]+', '-', str(page_name).lower()).strip('-') + '.html'
        out_path = os.path.join(PORTS_DIR, slug)

        html_content = generate_page_html(
            head_part=head_part,
            tail_part=tail_part,
            page_name=page_name,
            h1=h1,
            meta_title=meta_title,
            meta_desc=meta_desc,
            full_desc=full_desc,
            slug=slug
        )

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        generated_count += 1
        generated_files.append((num, page_name, slug))
        print(f"[{num:02d}/49] Generated: {slug} ({page_name})")

    print("\n" + "=" * 60)
    print(f"SUCCESS: Generated {generated_count} port pages across {len(set(f[2] for f in generated_files))} unique files.")
    print("Centralized CSS: assets/css/port-pages.css")
    print("Content Image: assets/imgs/warehouse-about.webp")
    print("Navbar & Footer Logo: assets/imgs/logo.png")
    print("=" * 60)


if __name__ == "__main__":
    main()
