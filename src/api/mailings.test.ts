import { describe, expect, it } from "vitest";
import {
  createDocx,
  type Docx,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";
import {
  addEnvelope,
  addressBlockInstruction,
  attachRecipientList,
  autoFieldMap,
  createLabelDocument,
  greetingLineInstruction,
  insertAddressBlock,
  insertGreetingLine,
  insertMergeField,
  insertMergeRule,
  LABEL_PRODUCTS,
  mailMergeFieldMap,
  mailMergeSettings,
  mergeFieldErrors,
  mergeFieldNames,
  mergeToNewDocument,
  parseRecipientCsv,
  previewMailMerge,
  recipientInclusion,
  recipientListToCsv,
  removeEnvelope,
  setMailMergeDocumentType,
  setRecipientInclusion,
  updateLabels,
} from "./mailings.js";
import { validatePackage } from "./validator.js";
import { paragraphFlow, visibleParagraphText } from "../internal/wordprocessingml/field-flow.js";

function partXml(doc: Docx, name: string): string {
  return new TextDecoder().decode(doc.opc.parts.get(name)?.data ?? new Uint8Array());
}

function roundTrip(doc: Docx): Docx {
  const reopened = openDocx(toUint8Array(doc));
  expect(validatePackage(reopened.opc)).toEqual([]);
  return reopened;
}

/** Visible text per paragraph (body and table cells), field codes hidden, line breaks as \n. */
function texts(doc: Docx): string[] {
  const reopened = roundTrip(doc);
  return paragraphFlow(reopened.document.body).map((ref) =>
    visibleParagraphText({
      ...ref.paragraph,
      children: ref.paragraph.children.map((c) =>
        c.kind === "run"
          ? { ...c, pieces: c.pieces.map((piece) => (piece.kind === "break" ? { kind: "text" as const, value: "\n", preserveSpace: true } : piece)) }
          : c,
      ),
    }),
  );
}

const CSV = 'First Name,Last Name,Title,Company,Address,City,State,ZIP,Country\r\nAda,Lovelace,Ms.,"Analytical, Ltd",1 Engine St,London,,N1,United Kingdom\r\nAlan,Turing,Mr.,,2 Bletchley Rd,Bletchley,Bucks,MK3,United States\r\n"Grace ""Amazing""",Hopper,,Navy,3 Cobol Ave,Arlington,VA,22201,United States\r\n';

describe("recipient lists", () => {
  it("parses RFC 4180 CSV and writes it back", () => {
    const list = parseRecipientCsv(CSV);
    expect(list.columns).toEqual(["First Name", "Last Name", "Title", "Company", "Address", "City", "State", "ZIP", "Country"]);
    expect(list.records).toHaveLength(3);
    expect(list.records[0]?.Company).toBe("Analytical, Ltd");
    expect(list.records[2]?.["First Name"]).toBe('Grace "Amazing"');
    expect(parseRecipientCsv(recipientListToCsv(list))).toEqual(list);
    expect(() => parseRecipientCsv("a,a\n1,2")).toThrow(/Duplicate/);
    expect(() => parseRecipientCsv("\n\n")).toThrow(/header/);
  });

  it("auto-matches address fields", () => {
    expect(autoFieldMap(parseRecipientCsv(CSV).columns)).toMatchObject({
      "First Name": "First Name",
      "Last Name": "Last Name",
      "Courtesy Title": "Title",
      "Address 1": "Address",
      "Postal Code": "ZIP",
      "Country or Region": "Country",
    });
  });

  it("links the list the way Word does and records exclusions", () => {
    const doc = createDocx({ paragraphs: ["Hello"] });
    const list = parseRecipientCsv(CSV);
    setMailMergeDocumentType(doc, "formLetters");
    attachRecipientList(doc, "/Users/me/Recipients.csv", list);
    setRecipientInclusion(doc, list, [true, false, true]);
    const reopened = roundTrip(doc);
    expect(mailMergeSettings(reopened)).toEqual({
      type: "formLetters",
      dataSource: "/Users/me/Recipients.csv",
      preview: false,
    });
    const settings = partXml(reopened, "/word/settings.xml");
    expect(settings).toMatch(/<w:mailMerge><w:mainDocumentType w:val="formLetters"\/><w:linkToQuery\/><w:dataType w:val="textFile"\/><w:connectString w:val=""\/><w:query w:val="SELECT \* FROM \/Users\/me\/Recipients\.csv"\/><w:dataSource r:id="rId\d+"\/><w:odso><w:udl w:val=""\/><w:table w:val="Recipients\.csv"\/><w:src r:id="rId\d+"\/><w:colDelim w:val="44"\/><w:fHdr\/><w:fieldMapData>/);
    expect(settings).toContain('<w:fieldMapData><w:type w:val="dbColumn"/><w:name w:val="ZIP"/><w:mappedName w:val="Postal Code"/><w:column w:val="7"/><w:lid w:val="en-US"/></w:fieldMapData>');
    expect(settings).toMatch(/<w:recipientData r:id="rId\d+"\/><\/w:odso><\/w:mailMerge>/);
    const rels = partXml(reopened, "/word/_rels/settings.xml.rels");
    expect(rels).toContain('Target="file:///Users/me/Recipients.csv" TargetMode="External"');
    expect(partXml(reopened, "/word/recipientData.xml")).toContain('<w:active w:val="0"/>');
    expect(recipientInclusion(reopened, list)).toEqual([true, false, true]);
    expect(mailMergeFieldMap(reopened)["Postal Code"]).toBe("ZIP");
    setMailMergeDocumentType(reopened, "normal");
    const plain = roundTrip(reopened);
    expect(mailMergeSettings(plain)).toBeUndefined();
    expect(partXml(plain, "/word/_rels/settings.xml.rels")).not.toContain("mailMergeSource");
  });
});

describe("merge fields", () => {
  it("builds Word's ADDRESSBLOCK and GREETINGLINE instructions", () => {
    expect(addressBlockInstruction()).toBe(
      'ADDRESSBLOCK \\f "<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>\n<<_COMPANY_\n>><<_STREET1_\n>><<_STREET2_\n>><<_CITY_>><<, _STATE_>><< _POSTAL_>><<\n_COUNTRY_>>" \\l 1033 \\c 2 \\e "United States" \\d',
    );
    expect(greetingLineInstruction()).toBe(
      'GREETINGLINE \\f "<<_BEFORE_ Dear >><<_TITLE0_ >><<_LAST0_>><<_AFTER_ ,>>" \\l 1033 \\e "Dear Sir or Madam,"',
    );
  });

  it("previews a recipient and goes back to placeholders", () => {
    const doc = createDocx({ paragraphs: ["", "", "Hi ", ""] });
    const list = parseRecipientCsv(CSV);
    attachRecipientList(doc, "Recipients.csv", list);
    const [p0, p1, p2, p3] = paragraphs(doc);
    if (!p0 || !p1 || !p2 || !p3) throw new Error("fixture");
    insertAddressBlock(doc, p0, 0);
    insertGreetingLine(doc, p1, 0);
    insertMergeField(doc, p2, 3, "First Name");
    insertMergeRule(doc, p3, 0, { kind: "if", field: "Country", comparison: "=", value: "United States", trueText: "Domestic", falseText: "International" });
    insertMergeRule(doc, p3, 0, { kind: "mergeRecord" });
    expect(texts(doc)).toEqual(["«AddressBlock»", "«GreetingLine»", "Hi «First Name»", "«Merge Record #»International"]);
    previewMailMerge(doc, list, 0);
    expect(texts(doc)).toEqual([
      "Ms. Ada Lovelace\nAnalytical, Ltd\n1 Engine St\nLondon N1\nUnited Kingdom",
      "Dear Ms. Lovelace,",
      "Hi Ada",
      "1International",
    ]);
    expect(mailMergeSettings(doc)).toMatchObject({ preview: true, activeRecord: 1 });
    previewMailMerge(doc, list, 2);
    expect(texts(doc)).toEqual([
      "Grace \"Amazing\" Hopper\nNavy\n3 Cobol Ave\nArlington, VA 22201",
      "Dear Hopper,",
      'Hi Grace "Amazing"',
      "3Domestic",
    ]);
    previewMailMerge(doc, list, undefined);
    expect(texts(doc)[2]).toBe("Hi «First Name»");
    expect(mailMergeSettings(doc)?.preview).toBe(false);
  });

  it("reports unknown merge fields", () => {
    const doc = createDocx({ paragraphs: [""] });
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    insertMergeField(doc, p, 0, "First Name");
    insertMergeField(doc, p, 0, "Nickname");
    insertMergeField(doc, p, 0, "Shoe Size");
    expect(mergeFieldNames(doc)).toEqual(["Shoe Size", "Nickname", "First Name"]);
    expect(mergeFieldErrors(doc, parseRecipientCsv(CSV))).toEqual(["Shoe Size"]);
  });

  it("writes every rule as its ECMA-376 field", () => {
    const doc = createDocx({ paragraphs: [""] });
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    insertMergeRule(doc, p, 0, { kind: "ask", bookmark: "Rate", prompt: "Rate?", defaultText: "5", askOnce: true });
    insertMergeRule(doc, p, 0, { kind: "fillIn", prompt: "Note?" });
    insertMergeRule(doc, p, 0, { kind: "mergeSequence" });
    insertMergeRule(doc, p, 0, { kind: "nextRecord" });
    insertMergeRule(doc, p, 0, { kind: "nextRecordIf", field: "City", comparison: "isBlank" });
    insertMergeRule(doc, p, 0, { kind: "setBookmark", bookmark: "Fee", value: "10" });
    insertMergeRule(doc, p, 0, { kind: "skipRecordIf", field: "State", comparison: "<>", value: "VA" });
    const xml = partXml(roundTrip(doc), "/word/document.xml");
    expect(xml).toContain('ASK Rate "Rate?" \\d "5" \\o');
    expect(xml).toContain('FILLIN "Note?"');
    expect(xml).toContain(" MERGESEQ ");
    expect(xml).toContain(" NEXT ");
    expect(xml).toMatch(/NEXTIF <\/w:instrText>.*?MERGEFIELD City.*? = "" /);
    expect(xml).toContain('SET Fee "10"');
    expect(xml).toMatch(/SKIPIF <\/w:instrText>.*?MERGEFIELD State.*? &lt;> "VA" /);
    expect(() => insertMergeRule(doc, p, 0, { kind: "setBookmark", bookmark: "1 bad", value: "" })).toThrow();
  });
});

describe("finish and merge", () => {
  it("makes one section per recipient, honouring range, exclusions and SKIPIF", () => {
    const doc = createDocx({ paragraphs: ["Dear "] });
    const list = parseRecipientCsv(CSV);
    attachRecipientList(doc, "r.csv", list);
    const [p] = paragraphs(doc);
    if (!p) throw new Error("fixture");
    insertMergeField(doc, p, 5, "Last Name");
    insertMergeRule(doc, p, 0, { kind: "skipRecordIf", field: "City", comparison: "=", value: "Bletchley" });
    const merged = mergeToNewDocument(doc, list);
    expect(texts(merged).filter((t) => t !== "")).toEqual(["Dear Lovelace", "Dear Hopper"]);
    const xml = partXml(roundTrip(merged), "/word/document.xml");
    expect(xml).not.toContain("MERGEFIELD");
    expect(xml).toContain('<w:type w:val="nextPage"/>');
    expect(mailMergeSettings(roundTrip(merged))).toBeUndefined();
    const ranged = mergeToNewDocument(doc, list, { from: 3, to: 3 });
    expect(texts(ranged).filter((t) => t !== "")).toEqual(["Dear Hopper"]);
    const excluded = mergeToNewDocument(doc, list, { included: [false, true, true] });
    expect(texts(excluded).filter((t) => t !== "")).toEqual(["Dear Hopper"]);
  });

  it("merges a label sheet with NEXT fields", () => {
    const labels = createLabelDocument({ text: "", product: LABEL_PRODUCTS[4]!, mailMerge: true });
    const table = labels.document.body.blocks[0];
    if (table?.kind !== "table") throw new Error("fixture");
    const first = table.rows[0]?.cells[0]?.paragraphs[0];
    if (!first) throw new Error("fixture");
    insertMergeField(labels, first, 0, "First Name");
    expect(updateLabels(labels)).toBe(5);
    const merged = mergeToNewDocument(labels, parseRecipientCsv(CSV));
    expect(texts(merged).filter((t) => t !== "")).toEqual(["Ada", "Alan", "Grace \"Amazing\""]);
  });
});

describe("envelopes and labels", () => {
  it("adds and replaces an envelope section", () => {
    const doc = createDocx({ paragraphs: ["Letter body"] });
    addEnvelope(doc, { deliveryAddress: "Ada Lovelace\n1 Engine St", returnAddress: "Me\nHere", size: "DL" });
    addEnvelope(doc, { deliveryAddress: "Alan Turing", returnAddress: "Me", omitReturnAddress: true });
    const reopened = roundTrip(doc);
    const xml = partXml(reopened, "/word/document.xml");
    expect(xml).toMatch(/^.*<w:body><w:p><w:pPr><w:pStyle w:val="EnvelopeAddress"\/><\/w:pPr><w:r><w:t>Alan Turing<\/w:t><\/w:r><\/w:p><w:p><w:pPr><w:sectPr><w:pgSz w:w="13680" w:h="5940" w:orient="landscape"\/>/s);
    expect(xml).not.toContain("Ada Lovelace");
    expect(partXml(reopened, "/word/styles.xml")).toContain('<w:framePr w:w="7920" w:h="1980"');
    expect(removeEnvelope(reopened)).toBe(true);
    expect(texts(reopened)).toEqual(["Letter body"]);
  });

  it("lays out a sheet of labels", () => {
    const doc = createLabelDocument({ text: "Ada\nLondon", product: LABEL_PRODUCTS[0]! });
    const xml = partXml(roundTrip(doc), "/word/document.xml");
    expect(xml).toContain('<w:gridCol w:w="3780"/><w:gridCol w:w="180"/><w:gridCol w:w="3780"/>');
    expect((xml.match(/<w:trHeight w:val="1440" w:hRule="exact"\/>/g) ?? []).length).toBe(10);
    expect((xml.match(/<w:t>Ada<\/w:t>/g) ?? []).length).toBe(30);
    expect(xml).toContain('<w:pgMar w:top="720" w:right="270" w:bottom="0" w:left="270"');
    const single = createLabelDocument({ text: "One", product: LABEL_PRODUCTS[0]!, single: { row: 2, column: 3 } });
    const singleXml = partXml(roundTrip(single), "/word/document.xml");
    expect((singleXml.match(/<w:t>One<\/w:t>/g) ?? []).length).toBe(1);
    expect(() => createLabelDocument({ text: "x", product: LABEL_PRODUCTS[0]!, single: { row: 11, column: 1 } })).toThrow();
  });
});

