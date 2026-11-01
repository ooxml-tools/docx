<h1>
    <picture>
        <source media="(prefers-color-scheme: dark)" srcset="https://ooxml-tools.github.io/design/images/docx-dark.png">
        <source media="(prefers-color-scheme: light)" srcset="https://ooxml-tools.github.io/design/images/docx-light.png">
        <img alt="@ooxml-tools/docx" height="56" src="https://ooxml-tools.github.io/design/images/docx-light.png">
    </picture>
</h1>

Helpers for the WordProcessingML (`.docx`) part of the OOXML spec

> [!NOTE]
> These are all based around [xml-js](https://www.npmjs.com/package/xml-js) and output that modules internal format


## API

### `renderDocument`
Rather than changing the struture of the document it's often easier to build up a new version and override it

```ts
import { safeXml } from "@ooxml-tools/xml"
import { renderDocument } from "@ooxml-tools/docx"

// This is a full document.xml you can write to `word/document.xml` 
const doc = renderDocument(safeXml`
  <w:p>
    <w:r>
        <w:t>
            Hello world!
        </w:t>
    </w:r>
  </w:p>
`)
```

### `insertLinks`
Creating links is a pain in OOXML, as you need to add the files to references to `word/_rels/document.xml.rels` then reference `rId`s in the main `document.xml`. The `insertLinks` helper makes that easier

```ts
import { insertLinks } from "@ooxml-tools/docx"

const links = await insertLinks(doc, {
    foo: "https://example.com/test1",
    bar: "https://example.com/test1",
    baz: "https://example.com/test2",
});

assert.equal(links, {
    foo: "rId8",
    // NOTE: This is the same as the above, it'll reuse existing links
    bar: "rId8",
    baz: "rId9"
});
```

## License
MIT