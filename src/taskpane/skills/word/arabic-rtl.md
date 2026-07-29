# Arabic and RTL Word Documents

Use this skill for Arabic or Hebrew text, mixed right-to-left/left-to-right
content, document sections, and tables.

## Safety rules

- Keep Unicode text in its logical order. Never reverse Arabic strings.
- Do not insert bidi control characters such as RLO/LRO/PDF into document text
  unless the user explicitly requests them.
- Do not apply RTL to the whole document when the user asked to change only a
  paragraph, selection, table, or section.
- Preserve embedded Latin words, numbers, URLs, email addresses, and product
  names in their natural order.
- `WordApiDesktop` APIs are desktop-only. Check the requirement set before
  using them and return a clear message when the current client cannot apply
  the requested property.

## Basic Arabic text (cross-platform)

Insertion uses normal Unicode strings. Alignment is available in the regular
Word API:

```javascript
const paragraph = context.document.body.insertParagraph(
  "مرحبًا بك في AutoOffice 2026",
  Word.InsertLocation.end,
);
paragraph.alignment = Word.Alignment.right;
await context.sync();
```

Right alignment is not the same as bidirectional reading order. Use the
desktop APIs below only when the user asks for real RTL section/table
formatting.

## Arabic language and bidirectional font properties

Requires `WordApiDesktop 1.3`.

```javascript
if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")) {
  throw new Error("Arabic language tagging requires WordApiDesktop 1.3.");
}

const range = context.document.getSelection();
range.languageId = Word.LanguageId.arabic;
range.font.nameBidirectional = "Arial";
range.font.sizeBidirectional = 14;
await context.sync();
```

## Right-to-left section

Requires `WordApiDesktop 1.3`. Apply only to the relevant section(s).

```javascript
if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")) {
  throw new Error("RTL section direction requires WordApiDesktop 1.3.");
}

const section = context.document.sections.getFirst();
section.pageSetup.sectionDirection = Word.SectionDirection.rightToLeft;
await context.sync();
```

## Right-to-left table

Requires `WordApiDesktop 1.4`.

```javascript
if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.4")) {
  throw new Error("RTL table direction requires WordApiDesktop 1.4.");
}

const table = context.document.body.tables.getFirstOrNullObject();
table.load("isNullObject");
await context.sync();

if (table.isNullObject) {
  throw new Error("No table was found.");
}

table.tableDirection = Word.TableDirection.rightToLeft;
await context.sync();
```

## Mixed Arabic and Latin verification

After inserting or replacing text, load the affected range's `text` property
and return it. This verifies that Arabic, Latin text, and numbers were not
reordered or corrupted:

```javascript
const range = context.document.getSelection();
range.insertText("التقرير Q3 — 2026", Word.InsertLocation.replace);
range.load("text");
await context.sync();
return { text: range.text };
```
