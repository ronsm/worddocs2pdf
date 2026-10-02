#!/usr/bin/env python3
"""Regenerate the worddocs2pdf test fixtures.

Writes three minimal .docx files of 1, 2 and 3 pages, named to exercise natural
sort (01_, 02_, 10_), plus two files the tool should skip: an orphan Word lock
file and a plain text file.

Run from the repository root:  python3 test/make-fixtures.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def paragraph(text, half_points=28):
    return (
        f'<w:p><w:pPr><w:rPr><w:sz w:val="{half_points}"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="{half_points}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def make_docx(path, title, pages):
    body = []
    for n in range(1, pages + 1):
        if n > 1:
            body.append(PAGE_BREAK)
        body.append(paragraph(f"{title} — page {n} of {pages}", 40))
        body.append(paragraph("Generated as a worddocs2pdf test fixture."))

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}"><w:body>{"".join(body)}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
        "</w:sectPr></w:body></w:document>"
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document)

    print(f"{os.path.relpath(path)}  ({pages} page{'s' if pages != 1 else ''})")


def main():
    for name, title, pages in [
        ("01_intro.docx", "Intro", 1),
        ("02_middle.docx", "Middle", 2),
        ("10_appendix.docx", "Appendix", 3),
    ]:
        make_docx(os.path.join(HERE, name), title, pages)

    # Files the tool must skip. The lock file is deliberately an orphan: a ~$ file
    # matching a real document would make Word itself complain when it tries to
    # write its own lock file for that document.
    skipped = {
        "~$zz_orphan.docx": "orphan lock file\n",
        "notes.txt": "not a Word document\n",
    }
    for name, text in skipped.items():
        with open(os.path.join(HERE, name), "w") as f:
            f.write(text)
        print(f"{os.path.join('test', name)}  (should be skipped)")


if __name__ == "__main__":
    main()
