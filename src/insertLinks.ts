import { asXmlElement, safeXml } from "@ooxml-tools/xml"
import { Element, js2xml } from "xml-js";
import { OfficeOpenXml } from "@ooxml-tools/file";
import { DOCUMENT_RELS } from "./constants";


function getExistingUrls(element: Element) {
    if (element.elements) {
        return element.elements
            .map((element: Element) => {
                return {
                    id: String(element.attributes?.Id),
                    url: String(element.attributes?.Target)
                }
            });
    }
    return [];
}

function getLargestId(idMap: { id: string, url: string }[]) {
    const allRIds = idMap
        .map(obj => parseInt(obj.id.replace(/^rId/, ""), 10))
        .filter(Number.isFinite)

    return Math.max(
        0,
        ...allRIds
    );
}

function getRelationshipsRoot(json: Element) {
    const expectedRelationshipsNode = json.elements?.[0]
    if (
        expectedRelationshipsNode &&
        expectedRelationshipsNode.type === "element" &&
        expectedRelationshipsNode.name === "Relationships" &&
        expectedRelationshipsNode.attributes?.xmlns === "http://schemas.openxmlformats.org/package/2006/relationships"
    ) {
        return expectedRelationshipsNode
    } else {
        const out = asXmlElement(safeXml`
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                </Relationships>       
        `)
        return json.elements![0]!
    }
}

export async function insertLinks<T extends Record<string, string>>(
    doc: OfficeOpenXml,
    links: T,
) {
    const content = await doc.readFile(DOCUMENT_RELS, "uint8array");
    const json = asXmlElement(new TextDecoder().decode(content));

    const subRoot = getRelationshipsRoot(json)

    const existingUrls = getExistingUrls(subRoot);

    const output: Record<keyof typeof links, string> = {} as Record<
        keyof typeof links,
        string
    >;

    Object.entries(links).forEach(([key, url]) => {
        // Check if the URL is already present in the existing IDs
        const existingUrl = existingUrls.find(item => item.url === url);
        if (existingUrl) {
            output[key as keyof T] = existingUrl.id;
            return;
        }
        const currentId = getLargestId(existingUrls)+1;
        const id = "rId" + currentId;
        existingUrls.push({ id, url });
        output[key as keyof T] = id;

        subRoot.elements!.push(
            asXmlElement(safeXml`
            <Relationship
                Id="${id}"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="${url}"
                TargetMode="External"
            />
        `).elements![0]!
        );
    });

    const fileContent = js2xml(json);
    doc.writeFile(DOCUMENT_RELS, fileContent);
    return output;
}