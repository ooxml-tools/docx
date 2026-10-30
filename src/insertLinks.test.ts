import { describe, expect, test } from "vitest";
import { insertLinks } from "./insertLinks";
import { init, docxBlankFiles } from "@ooxml-tools/file";
import { DOCUMENT_RELS } from "./constants";

describe("insertLinks", () => {
  test("generates single link", async () => {
    const doc = init("docx", docxBlankFiles);
    const links = await insertLinks(doc, {
      test: "https://example.com",
    });
    expect(links).toEqual({
      test: "rId7"
    });

    const fileContents = await doc.readFile(DOCUMENT_RELS, "text");
    expect(fileContents).toContain("https://example.com");
    expect(fileContents).toMatchSnapshot();
  });
  
  test("generates multiple links with different rId's", async () => {
    const doc = init("docx", docxBlankFiles);
    const links = await insertLinks(doc, {
      test1: "https://example.com/test1",
      test2: "https://example.com/test2",
    });
    expect(links).toEqual({
      test1: "rId7",
      test2: "rId8"
    });
    
    const fileContents = await doc.readFile(DOCUMENT_RELS, "text");
    expect(fileContents).toContain("https://example.com/test1");
    expect(fileContents).toContain("https://example.com/test2");
    expect(fileContents).toMatchSnapshot();
  });
  
  test("generates same id's when same link present", async () => {
    const doc = init("docx", docxBlankFiles);
    const links = await insertLinks(doc, {
      test1: "https://example.com",
      test2: "https://example.com",
    });
    expect(links).toEqual({
      test1: "rId7",
      test2: "rId7"
    });

    const fileContents = await doc.readFile(DOCUMENT_RELS, "text");
    expect(fileContents).toContain("https://example.com");
    expect(fileContents).toMatchSnapshot();
  });
});
