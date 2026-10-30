import { describe, expect, test } from "vitest";
import { renderDocument } from "./renderDocument";
import { safeXml } from "@ooxml-tools/xml";

describe("renderDocument", () => {
  test("renders with contents", () => {
    const bodyContent = safeXml`<w:p><w:r><w:t>Hello World!<w:t><w:r><w:p>`
    const output = renderDocument(bodyContent);
    expect(output).toContain(bodyContent);
    expect(output).toMatchSnapshot();
  });
});
