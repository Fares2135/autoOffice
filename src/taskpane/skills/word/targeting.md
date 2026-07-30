# Targeting the Right Part of the Document

Read this before any edit that applies to part of a document rather than all of
it. Most wrong edits are not bad API calls — they are correct calls aimed at
the wrong range.

## The rule

Locate, verify, then edit the narrowest range that satisfies the request.
Never widen the range to make the code simpler.

Scope ladder, narrowest first:

1. `context.document.getSelection()` — the user said "this", "here", "the
   selected text"
2. One paragraph, table or list you located by index
3. One section
4. `context.document.body` — only when the user asked for the whole document

## Editing the selection

```javascript
const range = context.document.getSelection();
range.font.bold = true;
await context.sync();
```

If the selection is empty, do not silently fall back to the whole document.
Say the selection is empty and ask what to target.

```javascript
const range = context.document.getSelection();
range.load("text");
await context.sync();
if (!range.text) {
  throw new Error("Nothing is selected. Ask the user what to change.");
}
```

## Editing a paragraph by index

`find_text` and `inspect_document` return 0-based paragraph indexes. They line
up with `body.paragraphs.items[i]` **in the document state at the moment they
were called**.

```javascript
const paragraphs = context.document.body.paragraphs;
paragraphs.load("items");
await context.sync();

const target = paragraphs.items[7]; // index from find_text
target.insertText("النص الجديد", Word.InsertLocation.replace);
await context.sync();
```

Indexes go stale. Inserting or deleting paragraphs shifts every index after the
edit point, so after such an edit the old indexes are wrong. Either:

- collect every target index first, then edit from the highest index down, or
- call `find_text` again after the edit.

## Editing search matches only

When the request is "replace X with Y", operate on the match ranges, not on the
body text.

```javascript
const results = context.document.body.search("التقرير", { matchCase: false });
results.load("items");
await context.sync();

// One sync for the whole batch — never sync inside the loop.
for (const item of results.items) {
  item.insertText("التقرير النهائي", Word.InsertLocation.replace);
}
await context.sync();
```

Never rebuild the document text with string replacement and write it back with
`body.insertText(..., Word.InsertLocation.replace)`: that destroys every style,
comment, field and tracked change in the document.

## Editing one table, row or cell

```javascript
const tables = context.document.body.tables;
tables.load("items");
await context.sync();

const table = tables.items[0];
const cell = table.getCell(2, 1); // row 2, column 1, both 0-based
cell.value = "42";
await context.sync();
```

To change only a column, loop the rows and touch that one cell per row — do not
restyle the whole table.

## Editing one section

```javascript
const sections = context.document.sections;
sections.load("items");
await context.sync();

const body = sections.items[0].body;
body.font.name = "Arial";
await context.sync();
```

## Verify before reporting success

Load the text of what you changed and return it, so the change report shows the
edit landed where it was meant to:

```javascript
const paragraphs = context.document.body.paragraphs;
paragraphs.load("items");
await context.sync();

const target = paragraphs.items[7];
target.font.bold = true;
target.load("text");
await context.sync();
return { edited: 7, text: target.text };
```

## Operations that are never partial

These touch everything. Only write them when the user asked for the whole
document, in so many words:

- `body.clear()`
- `body.insertText(..., Word.InsertLocation.replace)`
- setting a font or style on `body` directly
- looping every paragraph to restyle it
- `sectionDirection` on every section
