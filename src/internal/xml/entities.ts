const STANDARD_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

/**
 * Decode the five predefined XML entities and numeric character references.
 *
 * Unknown named entities are left untouched (with a leading `&`), which is
 * the safest behavior for OOXML (Word does not emit non-standard entities,
 * so seeing one suggests we should round-trip it verbatim rather than
 * silently corrupt it).
 */
export function decodeEntities(input: string): string {
  return input.replace(
    /&([a-zA-Z][a-zA-Z0-9]*|#x[0-9A-Fa-f]+|#[0-9]+);/g,
    (whole, body: string) => {
      if (body[0] === "#") {
        if (body[1] === "x" || body[1] === "X") {
          const code = Number.parseInt(body.slice(2), 16);
          return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        const code = Number.parseInt(body.slice(1), 10);
        return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
      }
      const lit = STANDARD_ENTITIES[body];
      return lit !== undefined ? lit : whole;
    },
  );
}

// Anything outside the XML 1.0 `Char` production
// (https://www.w3.org/TR/xml/#charsets): C0 controls other than tab/LF/CR,
// lone surrogates, U+FFFE and U+FFFF. These cannot appear in a well-formed
// document even as character references, so there is no escape to fall back
// on; the `u` flag makes a lone surrogate match as its own code point.
const NON_XML_CHAR = /[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u;

function rejectNonXmlChars(input: string): void {
  const match = NON_XML_CHAR.exec(input);
  if (match === null) return;
  // Every non-Char code point is in the BMP, so one UTF-16 unit is enough.
  const hex = match[0].charCodeAt(0).toString(16).toUpperCase().padStart(4, "0");
  throw new RangeError(`U+${hex} is not allowed in XML 1.0`);
}

/**
 * Escape a string for use inside an XML text node. `>` is escaped too because
 * the sequence `]]>` is not allowed in character data. Throws a `RangeError`
 * for characters XML 1.0 cannot represent at all.
 */
export function encodeText(input: string): string {
  rejectNonXmlChars(input);
  return input.replace(/[&<>]/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : "&gt;"));
}

/**
 * Escape a string for use inside an XML attribute value. The `quote` argument
 * determines which quote character is being used so we only escape the
 * matching variant. Throws a `RangeError` for characters XML 1.0 cannot
 * represent at all.
 */
export function encodeAttrValue(input: string, quote: '"' | "'" = '"'): string {
  rejectNonXmlChars(input);
  let out = input.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  if (quote === '"') {
    out = out.replace(/"/g, "&quot;");
  } else {
    out = out.replace(/'/g, "&apos;");
  }
  // Carriage returns and tabs must be normalized in attribute values per
  // XML 1.0 §3.3.3; we keep them readable by escaping rather than dropping.
  out = out.replace(/\r/g, "&#xD;").replace(/\n/g, "&#xA;").replace(/\t/g, "&#x9;");
  return out;
}
