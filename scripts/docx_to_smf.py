from pathlib import Path
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl

SRC = Path("docs/Historical_Battle_of_Tukayyid_Campaign_Rules.docx")
OUT = Path("docs/Historical_Battle_of_Tukayyid_Campaign_Rules_SMF.txt")

def blocks(doc):
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def clean(text):
    return text.replace("\u00a0", " ").strip()

def run_text(paragraph):
    parts = []
    for run in paragraph.runs:
        text = run.text.replace("\u00a0", " ")
        if not text.strip():
            parts.append(text)
            continue
        if run.bold:
            text = f"[b]{text}[/b]"
        if run.italic:
            text = f"[i]{text}[/i]"
        parts.append(text)
    return "".join(parts) if parts else clean(paragraph.text)

def cell_text(cell):
    vals = []
    for p in cell.paragraphs:
        value = run_text(p)
        if value:
            vals.append(value)
    return "[br]".join(vals)

out = []
seen_title = False
list_open = False
for block in blocks(Document(SRC)):
    if isinstance(block, Table):
        if list_open:
            out.append("[/list]"); list_open = False
        out.append("[table]")
        for ri, row in enumerate(block.rows):
            out.append("[tr]")
            for cell in row.cells:
                value = cell_text(cell)
                if ri == 0 and not value.startswith("[b]"):
                    value = f"[b]{value}[/b]"
                out.append(f"[td]{value}[/td]")
            out.append("[/tr]")
        out.append("[/table]\n")
        continue

    text = clean(block.text)
    if not text:
        continue
    style = block.style.name if block.style else ""

    if style.startswith("List Bullet"):
        if not list_open:
            out.append("[list]"); list_open = True
        out.append(f"[li]{run_text(block)}[/li]")
        continue
    if list_open:
        out.append("[/list]"); list_open = False

    if text == "HISTORICAL:":
        out.append("[center][size=18pt][b][color=#607FAA]HISTORICAL:[/color][/b][/size]")
        seen_title = True
    elif text == "BATTLE OF TUKAYYID":
        out.append("[size=24pt][b]BATTLE OF TUKAYYID[/b][/size]")
    elif text == "Changes-Only Campaign Playtest Draft":
        out.append("[i]Changes-Only Playtest Draft[/i][/center]\n")
    elif seen_title and text.startswith("HISTORICAL: BATTLE OF TUKAYYID"):
        continue
    elif text.startswith("In May 3052, the armies of ComStar"):
        paragraphs = [part.strip() for part in text.splitlines() if part.strip()]
        body = "\n\n".join(paragraphs[:-1])
        attribution = paragraphs[-1]
        out.append(f"[quote]{body}\n\n[i]{attribution}[/i][/quote]\n")
    elif style == "Heading 2":
        out.append(f"\n[size=14pt][b][color=#607FAA]{text}[/color][/b][/size]")
    elif text[:1].isdigit() and text.upper() == text:
        out.append(f"\n[hr]\n[size=18pt][b][color=#213854]{text}[/color][/b][/size]\n")
    else:
        out.append(run_text(block) + "\n")

if list_open:
    out.append("[/list]")
text = "\n".join(out)
OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(text.strip() + "\n", encoding="utf-8")
print(OUT.resolve())
