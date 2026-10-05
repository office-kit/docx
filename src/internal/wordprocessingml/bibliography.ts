/**
 * Bibliography sources (ECMA-376 Part 1 §22.6): the `<b:Sources>` custom XML
 * part, and the citation / bibliography text a CITATION or BIBLIOGRAPHY field
 * shows in each citation style.
 *
 * Word formats citations with XSLT style sheets that are not part of the
 * standard. APA (6th edition) and MLA (7th edition) are implemented in full
 * for the source types Word's dialog offers; the other styles follow their
 * published rules for the common source types (author, year, title,
 * publisher, journal, web site) and fall back to a generic entry otherwise.
 */

import { parseXml, type XmlElement } from "../xml/index.js";
import { escapeXml } from "./field-flow.js";

export const BIBLIOGRAPHY_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

/** ST_SourceType (§22.6.2.66), the source types of Word's Create Source dialog. */
export const SOURCE_TYPES = [
  "Book",
  "BookSection",
  "JournalArticle",
  "ArticleInAPeriodical",
  "ConferenceProceedings",
  "Report",
  "InternetSite",
  "DocumentFromInternetSite",
  "ElectronicSource",
  "Art",
  "SoundRecording",
  "Performance",
  "Film",
  "Interview",
  "Patent",
  "Case",
  "Misc",
] as const;
export type SourceType = (typeof SOURCE_TYPES)[number];

/** The simple-text children of `<b:Source>` (§22.6.2) this library reads and writes. */
export const SOURCE_FIELDS = [
  "Title",
  "Year",
  "Month",
  "Day",
  "City",
  "StateProvince",
  "CountryRegion",
  "Publisher",
  "JournalName",
  "PeriodicalTitle",
  "BookTitle",
  "ConferenceName",
  "Volume",
  "Issue",
  "Pages",
  "Edition",
  "Institution",
  "InternetSiteTitle",
  "URL",
  "YearAccessed",
  "MonthAccessed",
  "DayAccessed",
  "Medium",
  "StandardNumber",
  "ShortTitle",
  "Comments",
] as const;
export type SourceField = (typeof SOURCE_FIELDS)[number];

export interface PersonName {
  readonly last: string;
  readonly first?: string;
  readonly middle?: string;
}

/** One bibliography source (`<b:Source>`). */
export interface BibliographySource {
  /** The citation tag CITATION fields refer to (`<b:Tag>`). */
  readonly tag: string;
  readonly type: SourceType;
  readonly authors?: readonly PersonName[];
  /** A corporate author, used instead of personal authors (`<b:Corporate>`). */
  readonly corporateAuthor?: string;
  readonly editors?: readonly PersonName[];
  readonly fields?: Readonly<Partial<Record<SourceField, string>>>;
}

/** Word's citation styles: the `StyleName` and `SelectedStyle` it writes on `<b:Sources>`. */
export const CITATION_STYLES = {
  APA: { selectedStyle: "\\APASixthEditionOfficeOnline.xsl", kind: "authorDate" },
  Chicago: { selectedStyle: "\\CHICAGO.XSL", kind: "authorDate" },
  GB7714: { selectedStyle: "\\GB.XSL", kind: "numeric" },
  "GOST - Name Sort": { selectedStyle: "\\GostName.XSL", kind: "numeric" },
  "GOST - Title Sort": { selectedStyle: "\\GostTitle.XSL", kind: "numeric" },
  "Harvard - Anglia": {
    selectedStyle: "\\HarvardAnglia2008OfficeOnline.xsl",
    kind: "authorDate",
  },
  IEEE: { selectedStyle: "\\IEEE2006OfficeOnline.xsl", kind: "numeric" },
  "ISO 690 - First Element and Date": { selectedStyle: "\\ISO690.XSL", kind: "authorDate" },
  "ISO 690 - Numerical Reference": { selectedStyle: "\\ISO690Nmerical.XSL", kind: "numeric" },
  MLA: { selectedStyle: "\\MLASeventhEditionOfficeOnline.xsl", kind: "authorPage" },
  SIST02: { selectedStyle: "\\SIST02.XSL", kind: "numeric" },
  Turabian: { selectedStyle: "\\TURABIAN.XSL", kind: "authorDate" },
} as const;
export type CitationStyle = keyof typeof CITATION_STYLES;

export function isCitationStyle(name: string): name is CitationStyle {
  return Object.hasOwn(CITATION_STYLES, name);
}

/** A run of formatted text in a citation or bibliography entry. */
export interface StyledText {
  readonly text: string;
  readonly italic?: boolean;
}

/** The parsed `<b:Sources>` part. */
export interface SourcesPart {
  readonly style: CitationStyle;
  readonly sources: BibliographySource[];
}

const DEFAULT_STYLE: CitationStyle = "APA";

function childText(el: XmlElement, local: string): string | undefined {
  const child = el.children.find(
    (c): c is XmlElement => c.kind === "element" && c.name.local === local,
  );
  if (!child) return undefined;
  return child.children.map((c) => (c.kind === "text" ? c.value : "")).join("");
}

function child(el: XmlElement | undefined, local: string): XmlElement | undefined {
  return el?.children.find((c): c is XmlElement => c.kind === "element" && c.name.local === local);
}

function readPeople(contributor: XmlElement | undefined): PersonName[] {
  const list = child(contributor, "NameList");
  if (!list) return [];
  return list.children
    .filter((c): c is XmlElement => c.kind === "element" && c.name.local === "Person")
    .map((p) => {
      const person: { last: string; first?: string; middle?: string } = {
        last: childText(p, "Last") ?? "",
      };
      const first = childText(p, "First");
      const middle = childText(p, "Middle");
      if (first) person.first = first;
      if (middle) person.middle = middle;
      return person;
    });
}

function isSourceType(value: string): value is SourceType {
  return (SOURCE_TYPES as readonly string[]).includes(value);
}

/** Parse a `<b:Sources>` document. Unknown source types read as `Misc`. */
export function parseSources(xml: string): SourcesPart {
  const root = parseXml(xml).root;
  const styleName = root.attrs.find((a) => a.name.local === "StyleName")?.value ?? "";
  const sources: BibliographySource[] = [];
  for (const el of root.children) {
    if (el.kind !== "element" || el.name.local !== "Source") continue;
    const type = childText(el, "SourceType") ?? "Misc";
    const roles = child(el, "Author");
    const authorRole = child(roles, "Author");
    const corporate = childText(authorRole ?? el, "Corporate");
    const authors = readPeople(authorRole);
    const editors = readPeople(child(roles, "Editor"));
    const fields: Partial<Record<SourceField, string>> = {};
    for (const name of SOURCE_FIELDS) {
      const value = childText(el, name);
      if (value !== undefined && value !== "") fields[name] = value;
    }
    sources.push({
      tag: childText(el, "Tag") ?? "",
      type: isSourceType(type) ? type : "Misc",
      ...(authors.length ? { authors } : {}),
      ...(corporate ? { corporateAuthor: corporate } : {}),
      ...(editors.length ? { editors } : {}),
      fields,
    });
  }
  return { style: isCitationStyle(styleName) ? styleName : DEFAULT_STYLE, sources };
}

function peopleXml(people: readonly PersonName[]): string {
  return `<b:NameList>${people
    .map(
      (p) =>
        `<b:Person><b:Last>${escapeXml(p.last)}</b:Last>${
          p.first ? `<b:First>${escapeXml(p.first)}</b:First>` : ""
        }${p.middle ? `<b:Middle>${escapeXml(p.middle)}</b:Middle>` : ""}</b:Person>`,
    )
    .join("")}</b:NameList>`;
}

/** Serialize sources as a `<b:Sources>` document. */
export function writeSources(part: SourcesPart): string {
  const style = CITATION_STYLES[part.style];
  const sources = part.sources.map((s) => {
    const roles: string[] = [];
    if (s.corporateAuthor) {
      roles.push(`<b:Author><b:Corporate>${escapeXml(s.corporateAuthor)}</b:Corporate></b:Author>`);
    } else if (s.authors?.length) {
      roles.push(`<b:Author>${peopleXml(s.authors)}</b:Author>`);
    }
    if (s.editors?.length) roles.push(`<b:Editor>${peopleXml(s.editors)}</b:Editor>`);
    const fields = SOURCE_FIELDS.flatMap((name) => {
      const value = s.fields?.[name];
      return value ? [`<b:${name}>${escapeXml(value)}</b:${name}>`] : [];
    });
    return [
      "<b:Source>",
      `<b:Tag>${escapeXml(s.tag)}</b:Tag>`,
      `<b:SourceType>${s.type}</b:SourceType>`,
      roles.length ? `<b:Author>${roles.join("")}</b:Author>` : "",
      ...fields,
      "</b:Source>",
    ].join("");
  });
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<b:Sources xmlns:b="${BIBLIOGRAPHY_NS}" xmlns="${BIBLIOGRAPHY_NS}" SelectedStyle="${escapeXml(style.selectedStyle)}" StyleName="${escapeXml(part.style)}">`,
    ...sources,
    "</b:Sources>",
  ].join("");
}

// --- Formatting -------------------------------------------------------------

function initials(p: PersonName): string {
  return [p.first, p.middle]
    .filter((n): n is string => !!n)
    .flatMap((n) => n.split(/[\s-]+/))
    .filter(Boolean)
    .map((n) => `${n[0]?.toUpperCase()}.`)
    .join(" ");
}

function fullFirst(p: PersonName): string {
  return [p.first, p.middle].filter(Boolean).join(" ");
}

/** "A", "A and B", "A, B, and C" with the given conjunction. */
function joinNames(names: readonly string[], and: string, serialComma = true): string {
  if (names.length <= 1) return names[0] ?? "";
  if (names.length === 2) return `${names[0]} ${and} ${names[1]}`;
  return `${names.slice(0, -1).join(", ")}${serialComma ? "," : ""} ${and} ${names.at(-1)}`;
}

function lastNames(source: BibliographySource): string[] {
  if (source.corporateAuthor) return [source.corporateAuthor];
  return (source.authors ?? []).map((a) => a.last);
}

function field(source: BibliographySource, name: SourceField): string | undefined {
  const value = source.fields?.[name]?.trim();
  return value ? value : undefined;
}

/** The text a citation shows in author position (or the title when there is no author). */
function citationAuthor(source: BibliographySource, and: string, etAlFrom: number): string {
  const names = lastNames(source);
  if (names.length === 0) return field(source, "ShortTitle") ?? field(source, "Title") ?? "";
  if (names.length >= etAlFrom) return `${names[0]} et al.`;
  return joinNames(names, and, false);
}

const NO_DATE = "n.d.";

/** Options a CITATION field's switches carry (§17.16.5.9). */
export interface CitationOptions {
  readonly pages?: string;
  readonly volume?: string;
  readonly prefix?: string;
  readonly suffix?: string;
  readonly suppressAuthor?: boolean;
  readonly suppressYear?: boolean;
  readonly suppressTitle?: boolean;
}

/** One cited source within a CITATION field, with its 1-based reference number. */
export interface CitedSource {
  readonly source: BibliographySource;
  readonly number: number;
  readonly options: CitationOptions;
}

function authorDateCitation(cited: CitedSource, style: CitationStyle): string {
  const { source, options } = cited;
  const year = field(source, "Year") ?? NO_DATE;
  const and = style === "APA" ? "&" : "and";
  const etAl = style === "APA" ? 3 : 4;
  const author = options.suppressAuthor ? "" : citationAuthor(source, and, etAl);
  const shown = [author, options.suppressYear ? "" : year].filter(Boolean);
  // Chicago and Turabian write "Author Year, pages"; APA, Harvard and ISO 690
  // separate author and year with a comma.
  const spaceSeparated = style === "Chicago" || style === "Turabian";
  let text = shown.join(spaceSeparated ? " " : ", ");
  if (style === "ISO 690 - First Element and Date" && author && !options.suppressAuthor) {
    text = [author.toUpperCase(), options.suppressYear ? "" : year].filter(Boolean).join(", ");
  }
  if (options.volume) text += `, ${options.volume}`;
  if (options.pages) {
    const pageLabel = style === "APA" || style === "Harvard - Anglia" ? "p. " : "";
    text += `${text ? ", " : ""}${pageLabel}${options.pages}`;
  }
  return text;
}

function mlaCitation(cited: CitedSource): string {
  const { source, options } = cited;
  const author = options.suppressAuthor ? "" : citationAuthor(source, "and", 4);
  return [author, options.pages ?? ""].filter(Boolean).join(" ");
}

/**
 * The text of a CITATION field: one or more sources (`\m`) in the style's
 * brackets, with the field's prefix and suffix.
 */
export function formatCitation(cited: readonly CitedSource[], style: CitationStyle): string {
  const kind = CITATION_STYLES[style].kind;
  const parts = cited.map((c) => {
    let body: string;
    if (kind === "numeric")
      body = String(c.number) + (c.options.pages ? `, ${c.options.pages}` : "");
    else if (kind === "authorPage") body = mlaCitation(c);
    else body = authorDateCitation(c, style);
    return `${c.options.prefix ?? ""}${body}${c.options.suffix ?? ""}`;
  });
  if (kind === "numeric") {
    // GB7714, IEEE and GOST bracket numbers; ISO 690 numerical and SIST02 use parentheses.
    const paren = style === "ISO 690 - Numerical Reference" || style === "SIST02";
    return paren ? `(${parts.join(", ")})` : `[${parts.join("], [")}]`;
  }
  return `(${parts.join("; ")})`;
}

function apaAuthors(source: BibliographySource): string {
  if (source.corporateAuthor) return source.corporateAuthor;
  const names = (source.authors ?? []).map((a) => {
    const i = initials(a);
    return i ? `${a.last}, ${i}` : a.last;
  });
  if (names.length <= 1) return names[0] ?? "";
  return `${names.slice(0, -1).join(", ")}, & ${names.at(-1)}`;
}

function mlaAuthors(source: BibliographySource): string {
  if (source.corporateAuthor) return source.corporateAuthor;
  const people = source.authors ?? [];
  const first = people[0];
  if (!first) return "";
  const lead = fullFirst(first) ? `${first.last}, ${fullFirst(first)}` : first.last;
  if (people.length === 1) return lead;
  if (people.length >= 4) return `${lead}, et al`;
  const rest = people.slice(1).map((p) => [fullFirst(p), p.last].filter(Boolean).join(" "));
  return joinNames([lead, ...rest], "and");
}

function chicagoAuthors(source: BibliographySource): string {
  // Chicago and Turabian list the first author inverted and the rest in
  // natural order, which is MLA's author form without the "et al" cutoff.
  if (source.corporateAuthor) return source.corporateAuthor;
  const people = source.authors ?? [];
  const first = people[0];
  if (!first) return "";
  const lead = fullFirst(first) ? `${first.last}, ${fullFirst(first)}` : first.last;
  const rest = people.slice(1).map((p) => [fullFirst(p), p.last].filter(Boolean).join(" "));
  return joinNames([lead, ...rest], "and");
}

function ieeeAuthors(source: BibliographySource): string {
  if (source.corporateAuthor) return source.corporateAuthor;
  const names = (source.authors ?? []).map((a) => [initials(a), a.last].filter(Boolean).join(" "));
  return joinNames(names, "and");
}

function withPeriod(text: string): string {
  return /[.?!]$/.test(text) ? text : `${text}.`;
}

function place(source: BibliographySource): string {
  const city = field(source, "City");
  const publisher = field(source, "Publisher");
  return [city, publisher].filter(Boolean).join(": ");
}

/** The container a part-of-a-whole source sits in (journal, periodical, book, site). */
function container(source: BibliographySource): string | undefined {
  return (
    field(source, "JournalName") ??
    field(source, "PeriodicalTitle") ??
    field(source, "BookTitle") ??
    field(source, "ConferenceName") ??
    field(source, "InternetSiteTitle")
  );
}

const PART_TYPES: ReadonlySet<SourceType> = new Set([
  "JournalArticle",
  "ArticleInAPeriodical",
  "BookSection",
  "ConferenceProceedings",
  "DocumentFromInternetSite",
]);
const WEB_TYPES: ReadonlySet<SourceType> = new Set(["InternetSite", "DocumentFromInternetSite"]);

function apaEntry(source: BibliographySource): StyledText[] {
  const out: StyledText[] = [];
  const authors = apaAuthors(source);
  const title = field(source, "Title");
  const year = field(source, "Year") ?? NO_DATE;
  const date = WEB_TYPES.has(source.type)
    ? [year, [field(source, "Month"), field(source, "Day")].filter(Boolean).join(" ")]
        .filter(Boolean)
        .join(", ")
    : year;
  const lead = authors || (title ?? "");
  out.push({ text: `${withPeriod(lead)} (${date}). ` });
  const isPart = PART_TYPES.has(source.type);
  if (title && authors) {
    if (isPart) out.push({ text: `${withPeriod(title)} ` });
    else out.push({ text: withPeriod(title), italic: true }, { text: " " });
  }
  const host = container(source);
  if (isPart && host) {
    const volume = field(source, "Volume");
    const issue = field(source, "Issue");
    const pages = field(source, "Pages");
    if (source.type === "BookSection") {
      out.push({ text: "In " }, { text: host, italic: true });
      out.push({ text: `${pages ? ` (pp. ${pages})` : ""}. ` });
    } else {
      out.push({ text: volume ? `${host}, ${volume}` : host, italic: true });
      out.push({ text: `${issue ? `(${issue})` : ""}${pages ? `, ${pages}` : ""}. ` });
    }
  }
  const where = place(source);
  if (where && !WEB_TYPES.has(source.type)) out.push({ text: `${where}.` });
  const url = field(source, "URL");
  if (url) out.push({ text: `${out.length ? " " : ""}Retrieved from ${url}` });
  return trimEnd(out);
}

function mlaEntry(source: BibliographySource): StyledText[] {
  const out: StyledText[] = [];
  const authors = mlaAuthors(source);
  const title = field(source, "Title");
  const year = field(source, "Year");
  if (authors) out.push({ text: `${withPeriod(authors)} ` });
  const isPart = PART_TYPES.has(source.type);
  if (title) {
    if (isPart) out.push({ text: `"${withPeriod(title)}" ` });
    else out.push({ text: withPeriod(title), italic: true }, { text: " " });
  }
  const host = container(source);
  if (isPart && host) {
    out.push({ text: host, italic: true });
    const volume = field(source, "Volume");
    const issue = field(source, "Issue");
    const pages = field(source, "Pages");
    const vi = [volume, issue].filter(Boolean).join(".");
    out.push({
      text: `${vi ? ` ${vi}` : ""}${year ? ` (${year})` : ""}${pages ? `: ${pages}` : ""}. `,
    });
  } else {
    const where = place(source);
    const pub = [where, year].filter(Boolean).join(", ");
    if (pub) out.push({ text: `${pub}. ` });
  }
  out.push({
    text: WEB_TYPES.has(source.type) ? "Web." : (field(source, "Medium") ?? "Print") + ".",
  });
  const url = field(source, "URL");
  if (url) out.push({ text: ` <${url}>.` });
  return trimEnd(out);
}

function chicagoEntry(source: BibliographySource): StyledText[] {
  const out: StyledText[] = [];
  const authors = chicagoAuthors(source);
  const title = field(source, "Title");
  const year = field(source, "Year") ?? NO_DATE;
  if (authors) out.push({ text: `${withPeriod(authors)} ${year}. ` });
  const isPart = PART_TYPES.has(source.type);
  if (title) {
    if (isPart) out.push({ text: `"${withPeriod(title)}" ` });
    else out.push({ text: withPeriod(title), italic: true }, { text: " " });
  }
  const host = container(source);
  if (isPart && host) {
    const volume = field(source, "Volume");
    const issue = field(source, "Issue");
    const pages = field(source, "Pages");
    out.push({ text: host, italic: true });
    out.push({
      text: `${volume ? ` ${volume}` : ""}${issue ? ` (${issue})` : ""}${pages ? `: ${pages}` : ""}. `,
    });
  }
  const where = place(source);
  if (where) out.push({ text: `${where}.` });
  if (!authors) out.push({ text: ` ${year}.` });
  const url = field(source, "URL");
  if (url) out.push({ text: ` ${url}.` });
  return trimEnd(out);
}

function numericEntry(
  source: BibliographySource,
  number: number,
  style: CitationStyle,
): StyledText[] {
  const out: StyledText[] = [];
  const label =
    style === "ISO 690 - Numerical Reference" || style === "SIST02"
      ? `${number}. `
      : `[${number}] `;
  out.push({ text: label });
  const authors = style === "IEEE" ? ieeeAuthors(source) : chicagoAuthors(source);
  if (authors) out.push({ text: `${authors}, ` });
  const title = field(source, "Title");
  const isPart = PART_TYPES.has(source.type);
  if (title) {
    if (isPart) out.push({ text: `"${title}," ` });
    else out.push({ text: title, italic: true }, { text: ", " });
  }
  const host = container(source);
  if (isPart && host) out.push({ text: host, italic: true }, { text: ", " });
  const details = [
    field(source, "Volume") ? `vol. ${field(source, "Volume")}` : undefined,
    field(source, "Issue") ? `no. ${field(source, "Issue")}` : undefined,
    field(source, "Pages") ? `pp. ${field(source, "Pages")}` : undefined,
    place(source) || undefined,
    field(source, "Year"),
  ].filter(Boolean);
  if (details.length) out.push({ text: details.join(", ") });
  const url = field(source, "URL");
  if (url) out.push({ text: `. Available: ${url}` });
  out.push({ text: "." });
  return trimEnd(out);
}

function trimEnd(parts: StyledText[]): StyledText[] {
  const last = parts.at(-1);
  if (last) parts[parts.length - 1] = { ...last, text: last.text.trimEnd() };
  return parts.filter((p) => p.text !== "");
}

/** Sort key of a source in an author-sorted bibliography. */
function authorSortKey(source: BibliographySource): string {
  const lead = lastNames(source)[0] ?? field(source, "Title") ?? "";
  return `${lead.toLowerCase()}\u0000${field(source, "Year") ?? ""}\u0000${(field(source, "Title") ?? "").toLowerCase()}`;
}

/**
 * The entries of a BIBLIOGRAPHY field, in the style's order: by citation
 * number for numeric styles (`numbers` maps a tag to its number), by title
 * for GOST title sort, and by author otherwise. Sources never cited still
 * appear (after the cited ones in numeric styles), as in Word.
 */
export function formatBibliography(
  sources: readonly BibliographySource[],
  style: CitationStyle,
  numbers: ReadonlyMap<string, number>,
): StyledText[][] {
  const kind = CITATION_STYLES[style].kind;
  if (kind === "numeric") {
    let next = numbers.size;
    const numbered = sources.map((s) => ({ s, n: numbers.get(s.tag) ?? ++next }));
    numbered.sort((a, b) => a.n - b.n);
    if (style === "GOST - Title Sort" || style === "GOST - Name Sort") {
      const key =
        style === "GOST - Title Sort"
          ? (s: BibliographySource) => (field(s, "Title") ?? "").toLowerCase()
          : authorSortKey;
      numbered.sort((a, b) => key(a.s).localeCompare(key(b.s)));
      return numbered.map(({ s }, i) => numericEntry(s, i + 1, style));
    }
    return numbered.map(({ s, n }) => numericEntry(s, n, style));
  }
  const sorted = sources.toSorted((a, b) => authorSortKey(a).localeCompare(authorSortKey(b)));
  return sorted.map((s) => {
    if (style === "APA" || style === "Harvard - Anglia") return apaEntry(s);
    if (style === "MLA") return mlaEntry(s);
    return chicagoEntry(s);
  });
}
