# Search — Find and Replace

## Key Types
- `Word.SearchOptions` — matchCase, matchWholeWord, matchWildcards
- `Word.RangeCollection` — result of `.search()`

## Basic Search

```javascript
await Word.run(async (context) => {
  const results = context.document.body.search("hello");
  results.load("items");
  await context.sync();
  
  console.log("Found:", results.items.length, "matches");
  
  for (const range of results.items) {
    range.load("text");
  }
  await context.sync();
  
  for (const range of results.items) {
    console.log(range.text);
  }
});
```

## Search with Options

```javascript
await Word.run(async (context) => {
  const options = {
    matchCase: true,
    matchWholeWord: true,
  };
  
  const results = context.document.body.search("Word", options);
  results.load("items");
  await context.sync();
  
  // Highlight all matches
  for (const range of results.items) {
    range.font.highlightColor = Word.HighlightColor.yellow;
  }
  await context.sync();
});
```

## Find and Replace

```javascript
await Word.run(async (context) => {
  const results = context.document.body.search("old text");
  results.load("items");
  await context.sync();
  
  for (const range of results.items) {
    range.insertText("new text", Word.InsertLocation.replace);
  }
  await context.sync();
});
```

## Wildcard Search

```javascript
await Word.run(async (context) => {
  // Find dates in format XX/XX/XXXX
  const results = context.document.body.search("[0-9]{2}/[0-9]{2}/[0-9]{4}", {
    matchWildcards: true,
  });
  results.load("items");
  await context.sync();
  
  console.log("Found", results.items.length, "dates");
});
```

## Common Pitfalls

- `.search()` returns a `RangeCollection` — always load "items" and sync before iterating
- Wildcard syntax follows Word's wildcard rules, not standard regex
- Search is performed on the body; you can also search within a specific range or paragraph

## Arabic and mixed text: invisible marks break patterns

Mixed Arabic/Latin text usually carries invisible directional marks. An index
entry that looks like:

```
Module 26 | المقطع الصوتي
```

is frequently stored as `Module 26 <RLM>| المقطع الصوتي`, with U+200F between
the number and the pipe. So this fails, silently finding nothing:

```javascript
text.match(/Module\s+\d+\s*\|/); // no match — there is an RLM before the |
```

Strip the marks before matching, or use the `replace_text` tool, which ignores
them by default:

```javascript
const clean = text.replace(/[‎‏‪-‮⁦-⁩؜]/g, '');
```

`read_paragraphs` and `find_text` render these marks as `<RLM>`, `<LRM>` and
tabs as `\t`, so you can see what is really there without dumping character
codes.

## Never replace a paragraph to change part of it

An index paragraph often holds many entries separated by tabs:

```
040\tModule 88 | المقطع الصوتي\t\t\t041\tModule 89 | المقطع الصوتي\t\t\t
```

A regex ending in `.*$` matches to the end of the whole paragraph, so this
deletes every other entry while reporting success:

```javascript
// WRONG — destroys the rest of the paragraph
const m = p.text.match(/(Module\s+\d+)\s*\|.*$/);
p.insertText(m[1], Word.InsertLocation.replace);
```

Replace the matched range instead, or use `replace_text`:

```javascript
// Right: operate on the match, not the paragraph
const hits = p.search(oldFragment, { matchCase: true });
hits.load("items");
await context.sync();
hits.items[0].insertText(newFragment, Word.InsertLocation.replace);
await context.sync();
```
