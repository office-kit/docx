/**
 * The documents the editor opens on its own, per UI language: the blank one
 * "New" opens and the sample shown on first load. Word's Normal template
 * differs by language; the Japanese one is set for Japanese text, and a
 * Japanese user expects it.
 */

import {
  addTable,
  appendHeading,
  appendParagraph,
  createDocx,
  type Docx,
  ensureHeadingStyles,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
  setDefaultParagraphSpacing,
  setPageMargins,
  setPageSize,
  setParagraphAlignment,
  setRunFormat,
  setSectionProperties,
  updateStyleFormatting,
} from "@office-kit/docx";
import type { LocaleId } from "./i18n/locales";

// Japanese Word for Mac 16's Normal.dotm: A4 with 35 / 30 / 30 / 30 mm
// margins, a line grid of 18 pt, and 游明朝 10.5 pt justified with no
// space between paragraphs.
const JA_MARGINS = {
  top: 1985,
  bottom: 1701,
  left: 1701,
  right: 1701,
  header: 851,
  footer: 992,
  gutter: 0,
};
const JA_LINE_PITCH = 360;
const JA_FONT = "游明朝";
const JA_SIZE_HALF_POINTS = 21;
const SINGLE_LINE = 240;

export function newDocument(locale: LocaleId): Docx {
  return withTemplate(createDocx({ paragraphs: [""] }), locale);
}

function withTemplate(doc: Docx, locale: LocaleId): Docx {
  if (locale !== "ja") return doc;
  setPageSize(doc, PAGE_SIZE_A4);
  setPageMargins(doc, JA_MARGINS);
  setSectionProperties(doc, { documentGrid: { type: "lines", linePitch: JA_LINE_PITCH } });
  setDefaultParagraphSpacing(doc, { before: 0, after: 0, line: SINGLE_LINE });
  updateStyleFormatting(doc, "Normal", {
    run: (run) =>
      setRunFormat(run, {
        font: JA_FONT,
        fontEastAsia: JA_FONT,
        fontSizeHalfPoints: JA_SIZE_HALF_POINTS,
      }),
    paragraph: (p) => setParagraphAlignment(p, "both"),
  });
  return doc;
}

type SampleText = {
  title: string;
  intro: string;
  tryIt: string;
  tryBody: string;
  table: readonly (readonly string[])[];
};

const SAMPLE: Readonly<Record<LocaleId, SampleText>> = {
  en: {
    title: "Welcome to the word-kit editor",
    intro:
      "This is a Word-like editor. Every edit routes through @office-kit/docx, so what you save is a real .docx. Type here, use the ribbon above, then Save.",
    tryIt: "Try it",
    tryBody: "Select text and click Bold, or change alignment and color.",
    table: [
      ["Feature", "Status"],
      ["Text formatting", "Editable"],
      ["Tables", "Editable"],
    ],
  },
  ja: {
    title: "word-kit エディターへようこそ",
    intro:
      "Word のように使えるエディターです。編集はすべて @office-kit/docx を通るので、保存されるのは本物の .docx です。ここに入力し、上のリボンを使ってから保存してください。",
    tryIt: "試してみる",
    tryBody: "文字を選択して太字をクリックしたり、配置や色を変えたりしてみてください。",
    table: [
      ["機能", "状態"],
      ["文字の書式", "編集可能"],
      ["表", "編集可能"],
    ],
  },
  es: {
    title: "Bienvenido al editor word-kit",
    intro:
      "Este es un editor similar a Word. Cada edición pasa por @office-kit/docx, así que lo que guarda es un .docx real. Escriba aquí, use la cinta de opciones y guarde.",
    tryIt: "Pruébelo",
    tryBody: "Seleccione texto y haga clic en Negrita, o cambie la alineación y el color.",
    table: [
      ["Función", "Estado"],
      ["Formato de texto", "Editable"],
      ["Tablas", "Editable"],
    ],
  },
  fr: {
    title: "Bienvenue dans l’éditeur word-kit",
    intro:
      "Cet éditeur fonctionne comme Word. Chaque modification passe par @office-kit/docx : ce que vous enregistrez est un vrai fichier .docx. Saisissez du texte ici, utilisez le ruban, puis enregistrez.",
    tryIt: "Essayez",
    tryBody: "Sélectionnez du texte et cliquez sur Gras, ou changez l’alignement et la couleur.",
    table: [
      ["Fonctionnalité", "État"],
      ["Mise en forme du texte", "Modifiable"],
      ["Tableaux", "Modifiable"],
    ],
  },
  de: {
    title: "Willkommen im word-kit-Editor",
    intro:
      "Dies ist ein Editor wie Word. Jede Änderung läuft über @office-kit/docx, also speichern Sie eine echte .docx-Datei. Tippen Sie hier, nutzen Sie das Menüband und speichern Sie dann.",
    tryIt: "Ausprobieren",
    tryBody: "Markieren Sie Text und klicken Sie auf Fett, oder ändern Sie Ausrichtung und Farbe.",
    table: [
      ["Funktion", "Status"],
      ["Textformatierung", "Bearbeitbar"],
      ["Tabellen", "Bearbeitbar"],
    ],
  },
  zh: {
    title: "欢迎使用 word-kit 编辑器",
    intro:
      "这是一个类似 Word 的编辑器。每次编辑都通过 @office-kit/docx 完成，因此保存的是真正的 .docx 文件。在此输入，使用上方的功能区，然后保存。",
    tryIt: "试一试",
    tryBody: "选择文本并单击“加粗”，或更改对齐方式和颜色。",
    table: [
      ["功能", "状态"],
      ["文本格式", "可编辑"],
      ["表格", "可编辑"],
    ],
  },
};

/** The document shown when the editor first opens. */
export function sampleDocument(locale: LocaleId): Docx {
  const doc = createDocx({ paragraphs: [] });
  setPageSize(doc, PAGE_SIZE_A4);
  setPageMargins(doc, MARGINS_NORMAL);
  withTemplate(doc, locale);
  const text = SAMPLE[locale];
  ensureHeadingStyles(doc);
  appendHeading(doc, text.title, 1);
  appendParagraph(doc, text.intro);
  appendHeading(doc, text.tryIt, 2);
  appendParagraph(doc, text.tryBody);
  addTable(doc, text.table);
  return doc;
}
