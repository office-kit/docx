import { describe, expect, it } from "vitest";
import { getPart } from "../internal/opc/index.js";
import {
  legacyPasswordBytes,
  legacyVerifier,
  protectionHash,
  toBase64,
} from "../internal/wordprocessingml/protection-hash.js";
import {
  createDocx,
  type Docx,
  openDocx,
  paragraphs,
  setDocumentSettingOnOff,
  toUint8Array,
} from "./docx.js";
import {
  addEditableRange,
  documentProtection,
  editableRanges,
  protectDocument,
  removeEditableRange,
  setWriteProtection,
  unprotectDocument,
  verifyProtectionPassword,
  writeProtection,
} from "./protection.js";
import { validatePackage } from "./validator.js";

const SALT = Uint8Array.from({ length: 16 }, (_, i) => i);
// Few iterations keep the tests fast; the algorithm is the same.
const SPIN = 10;

function settingsXml(doc: Docx): string {
  const part = getPart(openDocx(toUint8Array(doc)).opc, "/word/settings.xml");
  return new TextDecoder().decode(part?.data);
}

describe("legacy password key", () => {
  it("matches the published verifier for 'password' (0x83AF)", () => {
    // The low word is the XOR-obfuscation verifier SpreadsheetML uses for
    // sheet protection, whose value for "password" is widely documented.
    expect(legacyVerifier(legacyPasswordBytes("password"))).toBe(0x83af);
  });

  it("hashes deterministically from salt and spin count", () => {
    const a = toBase64(protectionHash("secret", SALT, SPIN));
    expect(a).toBe(toBase64(protectionHash("secret", SALT, SPIN)));
    expect(a).not.toBe(toBase64(protectionHash("secreT", SALT, SPIN)));
    expect(a).not.toBe(toBase64(protectionHash("secret", SALT, SPIN + 1)));
  });
});

describe("document protection", () => {
  it("writes w:documentProtection in schema order and checks the password", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    setDocumentSettingOnOff(doc, "evenAndOddHeaders", true);
    protectDocument(doc, { edit: "readOnly", password: "pw", salt: SALT, spinCount: SPIN });
    const xml = settingsXml(doc);
    expect(xml).toMatch(
      /<w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="10" w:hash="[^"]+" w:salt="AAECAwQFBgcICQoLDA0ODw=="\/>/,
    );
    expect(xml.indexOf("documentProtection")).toBeLessThan(xml.indexOf("evenAndOddHeaders"));
    expect(documentProtection(doc)).toEqual({
      edit: "readOnly",
      formatting: false,
      enforced: true,
      hasPassword: true,
    });
    expect(verifyProtectionPassword(doc, "nope")).toBe(false);
    expect(unprotectDocument(doc, "nope")).toBe(false);
    expect(unprotectDocument(doc, "pw")).toBe(true);
    expect(documentProtection(doc)).toEqual({
      edit: "readOnly",
      formatting: false,
      enforced: false,
      hasPassword: false,
    });
    expect(validatePackage(openDocx(toUint8Array(doc)).opc)).toEqual([]);
  });

  it("protects for tracked changes with formatting limits and no password", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    protectDocument(doc, { edit: "trackedChanges", formatting: true });
    expect(settingsXml(doc)).toContain(
      '<w:documentProtection w:edit="trackedChanges" w:formatting="1" w:enforcement="1"/>',
    );
    expect(unprotectDocument(doc)).toBe(true);
  });

  it("writes Always Open Read-Only before every other setting", () => {
    const doc = createDocx({ paragraphs: ["a"] });
    protectDocument(doc, { edit: "comments" });
    setWriteProtection(doc, { recommended: true });
    expect(settingsXml(doc)).toMatch(/<w:settings[^>]*><w:writeProtection w:recommended="1"\/>/);
    expect(writeProtection(doc)).toEqual({ recommended: true, hasPassword: false });
    setWriteProtection(doc, undefined);
    expect(writeProtection(doc)).toBeUndefined();
  });
});

describe("editable ranges", () => {
  it("adds permStart / permEnd around the characters and removes them", () => {
    const doc = createDocx({ paragraphs: ["locked open locked"] });
    const para = paragraphs(doc)[0];
    if (!para) throw new Error("paragraph");
    const id = addEditableRange(doc, para, 7, 11);
    const xml = new TextDecoder().decode(
      getPart(openDocx(toUint8Array(doc)).opc, "/word/document.xml")?.data,
    );
    expect(xml).toContain(
      `<w:permStart w:id="${id}" w:edGrp="everyone"/><w:r><w:t>open</w:t></w:r><w:permEnd w:id="${id}"/>`,
    );
    expect(editableRanges(doc)).toEqual([{ id, group: "everyone", block: 0 }]);
    expect(removeEditableRange(doc, id)).toBe(true);
    expect(editableRanges(doc)).toEqual([]);
  });
});
