#!/usr/bin/env python3
"""
Build script to convert DOCX manuscript to HTML chapter pages.
Uses pandoc for reliable DOCX-to-HTML conversion, then splits the output
into individual chapter pages wrapped in a site template.
"""

import argparse
import re
import subprocess
from html import escape
from pathlib import Path
from typing import Dict, List, Optional, Set


def find_hr_following_texts(manuscript_path: Path) -> Set[str]:
    """Find text of paragraphs that follow horizontal rules in the DOCX.

    Word stores horizontal lines as VML <v:rect o:hr="t"> shapes inside
    <w:pict> elements.  Pandoc silently drops these, so we detect them with
    python-docx and record the *next* paragraph's text so we can inject
    <hr> tags into the pandoc HTML later.

    Returns a set of paragraph-text strings (first 80 chars, stripped) that
    should be preceded by an <hr>.
    """
    from docx import Document
    from docx.oxml.ns import qn

    VML_NS = "urn:schemas-microsoft-com:vml"
    OFFICE_NS = "urn:schemas-microsoft-com:office:office"

    doc = Document(str(manuscript_path))
    body = doc.element.body
    all_paras = body.findall(qn("w:p"))

    following_texts: Set[str] = set()

    for i, p in enumerate(all_paras):
        picts = p.findall(".//" + qn("w:pict"))
        for pict in picts:
            rects = pict.findall(f"{{{VML_NS}}}rect")
            for rect in rects:
                if rect.get(f"{{{OFFICE_NS}}}hr") == "t":
                    # Record first 80 chars of the NEXT paragraph's text
                    if i + 1 < len(all_paras):
                        next_text = "".join(
                            r.text or ""
                            for r in all_paras[i + 1].findall(".//" + qn("w:t"))
                        ).strip()
                        if next_text:
                            following_texts.add(next_text[:80])

    return following_texts


def inject_horizontal_rules(html: str, hr_following_texts: Set[str]) -> str:
    """Insert <hr /> before paragraphs whose text matches an HR-following text.

    We match by checking if a <p> tag's inner text (tags stripped) starts with
    one of the recorded following-texts.  Because the same text may appear in
    both the TOC and the content, *all* matching positions get an <hr> — but
    the TOC entries are short link text that won't match the 80-char snippets.
    """
    if not hr_following_texts:
        return html

    # Build a regex that finds <p> or <p class="..."> tags
    # We'll iterate through all <p...>...</p> segments
    tag_re = re.compile(r"<p(?:\s[^>]*)?>")
    strip_tags_re = re.compile(r"<[^>]+>")

    result_parts = []
    last_end = 0

    for m in tag_re.finditer(html):
        p_start = m.start()
        # Find the closing </p>
        close_idx = html.find("</p>", m.end())
        if close_idx == -1:
            continue

        # Extract inner text (strip HTML tags)
        inner_html = html[m.end():close_idx]
        inner_text = strip_tags_re.sub("", inner_html).strip()

        # Check if this paragraph's text matches any HR-following text
        matched = False
        for ft in hr_following_texts:
            if inner_text.startswith(ft[:40]):
                matched = True
                break

        if matched:
            # Insert <hr /> before this <p> tag
            result_parts.append(html[last_end:p_start])
            result_parts.append("<hr />\n")
            result_parts.append(html[p_start:close_idx + 4])
            last_end = close_idx + 4

    # Append remainder
    result_parts.append(html[last_end:])
    return "".join(result_parts)


# Chapter definitions: marker is used to find chapter boundaries in the pandoc HTML
CHAPTERS = [
    {
        "num": None,
        "marker": "Introduction",
        "title": "Introduction to Being an Executor",
        "slug": "introduction-to-being-an-executor",
        "skip": False,
    },
    {
        "num": 1,
        "marker": "Chapter 1",
        "title": "First Steps After Death",
        "slug": "first-steps-after-death",
        "skip": False,
    },
    {
        "num": 2,
        "marker": "Chapter 2",
        "title": "Understanding the Will and Probate Process",
        "slug": "understanding-the-will-and-probate-process",
        "skip": False,
    },
    {
        "num": 3,
        "marker": "Chapter 3",
        "title": "Probate",
        "slug": "probate",
        "skip": False,
    },
    {
        "num": 4,
        "marker": "Chapter 4",
        "title": "Identifying and Protecting Estate Assets",
        "slug": "identifying-and-protecting-estate-assets",
        "skip": False,
    },
    {
        "num": 5,
        "marker": "Chapter 5",
        "title": "Handling Estate Liabilities and Debts",
        "slug": "handling-estate-liabilities-and-debts",
        "skip": False,
    },
    {
        "num": 6,
        "marker": "Chapter 6",
        "title": "Taxes and the Estate",
        "slug": "taxes-and-the-estate",
        "skip": False,
    },
    {
        "num": 7,
        "marker": "Chapter 7",
        "title": "Distributing the Estate to Beneficiaries",
        "slug": "distributing-the-estate-to-beneficiaries",
        "skip": False,
    },
    {
        "num": 8,
        "marker": "Chapter 8",
        "title": "Executor Compensation and Duties",
        "slug": "executor-compensation-and-duties",
        "skip": False,
    },
    {
        "num": 9,
        "marker": "Chapter 9",
        "title": "Working with Professionals",
        "slug": "working-with-professionals",
        "skip": False,
    },
    {
        "num": 10,
        "marker": "Chapter 10",
        "title": "Special Considerations for Complex Estates",
        "slug": "special-considerations-for-complex-estates",
        "skip": False,
    },
    {
        "num": 11,
        "marker": "Chapter 11",
        "title": "Handling Family Disputes and Legal Challenges",
        "slug": "handling-family-disputes-and-legal-challenges",
        "skip": False,
    },
    {
        "num": 12,
        "marker": "Chapter 12",
        "title": "Common Pitfalls and How to Avoid Them",
        "slug": "common-pitfalls-and-how-to-avoid-them",
        "skip": False,
    },
    {
        "num": 13,
        "marker": "Chapter 13",
        "title": "Legal Challenges and Litigation Risks",
        "slug": "legal-challenges-and-litigation-risks",
        "skip": False,
    },
    {
        "num": 14,
        "marker": "Chapter 14",
        "title": "Creditor Lawsuits",
        "slug": "creditor-lawsuits",
        "skip": False,
    },
    {
        "num": 15,
        "marker": "Chapter 15",
        "title": "Finalizing the Estate",
        "slug": "finalizing-the-estate",
        "skip": False,
    },
]


def run_pandoc(manuscript_path: Path) -> str:
    """Convert DOCX to HTML using pandoc."""
    result = subprocess.run(
        ["pandoc", str(manuscript_path), "-t", "html", "--wrap=none"],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"ERROR: pandoc failed: {result.stderr}")
        raise RuntimeError("pandoc conversion failed")
    return result.stdout


def find_content_chapter_positions(html: str) -> List[Dict]:
    """Find chapter start positions in the pandoc HTML.

    Pandoc outputs the DOCX as flat HTML. Chapter markers appear as bold paragraphs
    like <p><strong>Introduction</strong></p> or <p><strong>Chapter 1: Title</strong></p>.
    They appear twice: once in the table of contents, once in the actual content.
    We want the SECOND occurrence of each marker (the content, not the TOC).
    """
    positions = []

    for chapter in CHAPTERS:
        if chapter["skip"]:
            continue

        marker = chapter["marker"]

        # Build patterns to match this chapter's marker in bold paragraphs.
        # Pandoc may produce:
        #   <p><strong>Chapter 1</strong></p>
        #   <p><strong>Chapter 1: First Steps After Death</strong></p>
        #   <p><strong>Introduction</strong></p>
        # Some markers may also appear inside blockquotes.
        patterns = [
            # Exact marker in a bold paragraph
            re.compile(
                rf'<p><strong>{re.escape(marker)}(?::?\s+[^<]*)?' + r'</strong></p>'
            ),
            # Marker inside a blockquote bold paragraph
            re.compile(
                rf'<blockquote>\s*<p><strong>{re.escape(marker)}(?::?\s+[^<]*)?' + r'</strong></p>'
            ),
        ]

        # Find all occurrences and take the last one (content, not TOC)
        all_matches = []
        for pattern in patterns:
            for m in pattern.finditer(html):
                all_matches.append(m)

        if not all_matches:
            print(f"WARNING: Could not find marker for '{marker}'")
            continue

        # Sort by position, take the last match (which is the content version)
        all_matches.sort(key=lambda m: m.start())
        content_match = all_matches[-1]

        positions.append({
            "chapter": chapter,
            "start": content_match.start(),
            "match_end": content_match.end(),
        })

    # Sort by position in the document
    positions.sort(key=lambda p: p["start"])
    return positions


def strip_chapter_header(html: str, chapter: Dict) -> str:
    """Remove the chapter marker and title paragraphs from the start of chapter HTML.

    The chapter HTML starts with the bold marker paragraph (e.g. <p><strong>Chapter 1</strong></p>)
    followed optionally by a title paragraph. We strip these since we add our own <h1>.
    """
    # Strip leading whitespace
    html = html.strip()

    # Remove the marker paragraph (bold paragraph at the start)
    marker = chapter["marker"]

    # Patterns for the marker paragraph and optional title paragraph that follows
    header_patterns = [
        # <blockquote><p><strong>Chapter N</strong></p><p><strong>Title</strong></p></blockquote>
        re.compile(
            rf'^<blockquote>\s*<p><strong>{re.escape(marker)}(?::?\s+[^<]*)?</strong></p>'
            rf'(?:\s*<p><strong>[^<]*</strong></p>)*'
            rf'\s*</blockquote>'
        ),
        # <p><strong>Chapter N: Title</strong></p> (combined marker+title)
        re.compile(
            rf'^<p><strong>{re.escape(marker)}:\s+[^<]*</strong></p>'
        ),
        # <p><strong>Chapter N</strong></p> followed by <p><strong>Title</strong></p>
        re.compile(
            rf'^<p><strong>{re.escape(marker)}</strong></p>'
            rf'(?:\s*<p><strong>[^<]*</strong></p>)*'
        ),
        # <p><strong>Introduction</strong></p> followed by <p><strong>Being an Executor</strong></p>
        re.compile(
            rf'^<p><strong>{re.escape(marker)}</strong></p>'
            rf'(?:\s*<p><strong>[^<]*</strong></p>)*'
        ),
    ]

    for pattern in header_patterns:
        html = pattern.sub('', html, count=1).strip()
        if not html.startswith(f'<p><strong>{marker}'):
            break

    return html



def clean_pandoc_html(html: str) -> str:
    """Clean up pandoc's HTML output for our site.

    Removes unwanted attributes, normalizes table markup, etc.
    """
    # Remove type attribute from ordered lists (e.g. <ol type="1">)
    html = re.sub(r'<ol[^>]*>', '<ol>', html)

    # Remove class attributes from table elements
    html = re.sub(r'<tr class="[^"]*">', '<tr>', html)
    html = re.sub(r'<th class="[^"]*">', '<th>', html)
    html = re.sub(r'<td class="[^"]*">', '<td>', html)

    # Remove empty paragraphs inside list items: <li><p>text</p></li> → <li>text</li>
    html = re.sub(r'<li><p>(.*?)</p></li>', r'<li>\1</li>', html)

    # Remove <br /> tags (soft returns that pandoc preserves)
    html = re.sub(r'<br\s*/?>', '', html)

    # Unwrap blockquotes — the DOCX uses shifted margins rather than intentional
    # blockquotes, so pandoc's blockquote interpretation is incorrect
    html = html.replace('<blockquote>', '').replace('</blockquote>', '')

    # Remove span.underline (pandoc wraps underlined text)
    html = re.sub(r'<span class="underline">(.*?)</span>', r'\1', html)

    # Convert single-cell tables to callout divs
    html = re.sub(
        r'<table>\s*<tbody>\s*<tr>\s*<td>(.*?)</td>\s*</tr>\s*</tbody>\s*</table>',
        r'<div class="callout">\1</div>',
        html,
        flags=re.DOTALL,
    )

    # Add class to remaining tables
    html = html.replace('<table>', '<table class="content-table">')

    # Add section-title class to bold-only paragraphs for visual separation
    html = re.sub(
        r'<p><strong>((?:(?!</strong>).)*)</strong></p>',
        r'<p class="section-title"><strong>\1</strong></p>',
        html,
    )

    # Clean up excessive whitespace
    html = re.sub(r'\n{3,}', '\n\n', html)

    return html.strip()


def generate_nav_html(chapters: List[Dict], current_slug: str, from_chapter: bool = False) -> str:
    """Generate navigation HTML with current page marked as active."""
    nav_html = ""
    chapter_prefix = "" if from_chapter else "chapters/"
    root_prefix = "../" if from_chapter else ""

    for chapter in chapters:
        if chapter["skip"]:
            continue
        is_active = chapter["slug"] == current_slug
        class_attr = "nav-link active" if is_active else "nav-link"
        nav_html += f'<a href="{chapter_prefix}{chapter["slug"]}.html" class="{class_attr}">{escape(chapter["title"])}</a>\n'

    # Add Executor Resources link at the end
    is_active = "executor-resources" == current_slug
    class_attr = "nav-link active" if is_active else "nav-link"
    nav_html += f'<a href="{root_prefix}executor-resources.html" class="{class_attr}">Executor Resources</a>\n'

    return nav_html


def build_site(manuscript_path: Optional[Path] = None):
    """Main build function."""
    script_dir = Path(__file__).parent
    site_dir = script_dir
    template_path = site_dir / "template.html"

    # Determine manuscript path
    if manuscript_path is None:
        manuscript_path = script_dir / "manuscript" / "Canadian_Executor_Guide_Book_Manuscript.docx"
        if not manuscript_path.exists():
            manuscript_path = (
                script_dir.parent
                / "mnt/Canadian Executor Guide — Claude/Claude Reference/Canadian_Executor_Guide_Book_Manuscript.docx"
            )

    # Read template
    if not template_path.exists():
        print(f"ERROR: template.html not found at {template_path}")
        return False

    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    # Ensure chapters output directory exists
    chapters_dir = site_dir / "chapters"
    chapters_dir.mkdir(exist_ok=True)

    # Check manuscript exists
    if not manuscript_path.exists():
        print(f"ERROR: Manuscript not found at {manuscript_path}")
        return False

    # Detect horizontal rules in the DOCX (pandoc drops VML shapes)
    print("Scanning DOCX for horizontal rules...")
    hr_following_texts = find_hr_following_texts(manuscript_path)
    print(f"Found {len(hr_following_texts)} horizontal rules")

    # Convert DOCX to HTML via pandoc
    print(f"Converting manuscript with pandoc: {manuscript_path}")
    full_html = run_pandoc(manuscript_path)
    print(f"Pandoc output: {len(full_html)} characters")

    # Inject <hr> tags where pandoc dropped them
    full_html = inject_horizontal_rules(full_html, hr_following_texts)

    # Find chapter positions in the HTML
    chapter_positions = find_content_chapter_positions(full_html)
    print(f"Found {len(chapter_positions)} chapter boundaries")

    # Extract and process each chapter
    generated_chapters = []
    for idx, pos in enumerate(chapter_positions):
        chapter = pos["chapter"]
        if chapter["skip"]:
            continue

        # Extract HTML between this chapter's start and the next chapter's start
        start = pos["start"]
        end = chapter_positions[idx + 1]["start"] if idx + 1 < len(chapter_positions) else len(full_html)
        chapter_html = full_html[start:end]

        # For the last chapter, truncate at the Disclaimer (back matter)
        if idx + 1 == len(chapter_positions):
            disclaimer_match = re.search(
                r'<p[^>]*><strong>Disclaimer</strong></p>', chapter_html
            )
            if disclaimer_match:
                chapter_html = chapter_html[:disclaimer_match.start()]

        # Strip the chapter header (marker + title paragraphs)
        chapter_html = strip_chapter_header(chapter_html, chapter)

        # Clean up pandoc output
        chapter_html = clean_pandoc_html(chapter_html)

        # Add our own h1
        chapter_html = f"<h1>{escape(chapter['title'])}</h1>\n{chapter_html}"

        # Add next-chapter link
        if idx + 1 < len(chapter_positions):
            next_chapter = chapter_positions[idx + 1]["chapter"]
            chapter_html += (
                f'\n<nav class="chapter-nav">'
                f'<a href="{next_chapter["slug"]}.html" class="next-chapter">'
                f'{escape(next_chapter["title"])} <span class="next-chapter-arrow">&rarr;</span></a>'
                f'</nav>'
            )

        # Generate navigation
        nav_html = generate_nav_html(CHAPTERS, chapter["slug"], from_chapter=True)

        # Fill template
        page_content = (
            template.replace("{{TITLE}}", escape(chapter["title"]))
            .replace("{{SLUG}}", chapter["slug"])
            .replace("{{CONTENT}}", chapter_html)
            .replace("{{NAV_ITEMS}}", nav_html)
        )

        # Write output file
        output_file = chapters_dir / f"{chapter['slug']}.html"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(page_content)

        # Clean up old flat files at root level
        old_flat_file = site_dir / f"{chapter['slug']}.html"
        if old_flat_file.exists():
            try:
                old_flat_file.unlink()
            except OSError:
                pass

        print(f"✓ Generated chapters/{chapter['slug']}.html ({len(chapter_html)} chars)")
        generated_chapters.append(chapter["slug"])

    print(f"\n{'='*60}")
    print(f"Build complete! Generated {len(generated_chapters)} pages:")
    for slug in generated_chapters:
        print(f"  - {slug}.html")

    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Build HTML pages from manuscript DOCX file"
    )
    parser.add_argument(
        "--manuscript", "-m",
        type=Path,
        default=None,
        help="Path to the manuscript DOCX file"
    )
    args = parser.parse_args()

    success = build_site(manuscript_path=args.manuscript)
    exit(0 if success else 1)
