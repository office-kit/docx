/**
 * SmartArt definition parts (ECMA-376 Part 1, §21.4): the layout definitions
 * (§21.4.2), style definitions (§21.4.4) and colour definitions (§21.4.4.x)
 * a diagram references. A diagram embeds its own copy of each, and Word lays
 * the diagram out again from the data model and the layout definition when
 * it opens the file; the `dsp:` drawing part Word also writes is an extension
 * and is not produced.
 *
 * The definitions are written from the schema, reusing the unique ids of
 * Word's built-in layouts (`urn:microsoft.com/office/officeart/2005/8/...`) so
 * Word matches them to its gallery entries.
 */

const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE_ART = "urn:microsoft.com/office/officeart/2005/8";

export const SMARTART_LAYOUTS = {
  basicBlockList: { id: `${OFFICE_ART}/layout/default`, category: "list" },
  verticalBulletList: { id: `${OFFICE_ART}/layout/vList2`, category: "list" },
  basicProcess: { id: `${OFFICE_ART}/layout/process1`, category: "process" },
  basicCycle: { id: `${OFFICE_ART}/layout/cycle2`, category: "cycle" },
  hierarchy: { id: `${OFFICE_ART}/layout/hierarchy1`, category: "hierarchy" },
} as const;
export type SmartArtLayout = keyof typeof SMARTART_LAYOUTS;

export const SMARTART_COLORS = {
  accent1_2: { category: "accent1", fill: ["accent1"] },
  accent2_2: { category: "accent2", fill: ["accent2"] },
  accent3_2: { category: "accent3", fill: ["accent3"] },
  accent4_2: { category: "accent4", fill: ["accent4"] },
  accent5_2: { category: "accent5", fill: ["accent5"] },
  accent6_2: { category: "accent6", fill: ["accent6"] },
  colorful1: {
    category: "colorful",
    fill: ["accent2", "accent3", "accent4", "accent5", "accent6"],
  },
} as const satisfies Readonly<Record<string, { category: string; fill: readonly string[] }>>;
export type SmartArtColors = keyof typeof SMARTART_COLORS;

/** SmartArt Styles: Simple Fill … Intense Effect, by the theme matrix entries they use. */
export const SMARTART_STYLES = {
  simple1: { line: 2, fill: 1, effect: 0 },
  simple2: { line: 3, fill: 1, effect: 0 },
  simple3: { line: 2, fill: 1, effect: 1 },
  simple4: { line: 2, fill: 1, effect: 2 },
  simple5: { line: 2, fill: 1, effect: 3 },
} as const;
export type SmartArtStyle = keyof typeof SMARTART_STYLES;

export const layoutUniqueId = (layout: SmartArtLayout): string => SMARTART_LAYOUTS[layout].id;
export const colorsUniqueId = (colors: SmartArtColors): string => `${OFFICE_ART}/colors/${colors}`;
export const styleUniqueId = (style: SmartArtStyle): string => `${OFFICE_ART}/quickstyle/${style}`;

const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

/** A text node: rectangle (or other shape) showing its point's text, sized by font rules. */
function textNode(name: string, shape: string, presOf: string, extraConstr = ""): string {
  return (
    `<dgm:layoutNode name="${name}"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst>` +
    `<dgm:alg type="tx"/><dgm:shape type="${shape}" r:blip=""><dgm:adjLst/></dgm:shape>` +
    `<dgm:presOf ${presOf}/>` +
    `<dgm:constrLst>` +
    `<dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/>` +
    `<dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/>` +
    `<dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/>` +
    `<dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/>${extraConstr}` +
    `</dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst>` +
    `</dgm:layoutNode>`
  );
}

const SPACE_NODE = (name: string): string =>
  `<dgm:layoutNode name="${name}"><dgm:alg type="sp"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape>` +
  `<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>`;

function layoutBody(layout: SmartArtLayout): string {
  switch (layout) {
    case "basicBlockList":
      return (
        `<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>` +
        `<dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/>` +
        `<dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg>` +
        `<dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/>` +
        `<dgm:constrLst>` +
        `<dgm:constr type="w" for="ch" forName="node" refType="w"/>` +
        `<dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/>` +
        `<dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.1"/>` +
        `<dgm:constr type="sp" refType="w" refFor="ch" refForName="sibTrans"/>` +
        `<dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>` +
        `</dgm:constrLst><dgm:ruleLst/>` +
        `<dgm:forEach name="nodesForEach" axis="ch" ptType="node">` +
        textNode("node", "rect", `axis="desOrSelf" ptType="node"`) +
        `<dgm:forEach name="sibTransForEach" axis="followSib" ptType="sibTrans" cnt="1">${SPACE_NODE("sibTrans")}</dgm:forEach>` +
        `</dgm:forEach></dgm:layoutNode>`
      );
    case "verticalBulletList":
      return (
        `<dgm:layoutNode name="linear"><dgm:varLst><dgm:animLvl val="lvl"/><dgm:resizeHandles val="exact"/></dgm:varLst>` +
        `<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/><dgm:param type="vertAlign" val="mid"/></dgm:alg>` +
        `<dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/>` +
        `<dgm:constrLst>` +
        `<dgm:constr type="w" for="ch" forName="parentText" refType="w"/>` +
        `<dgm:constr type="h" for="ch" forName="parentText" refType="primFontSz" refFor="ch" refForName="parentText" fact="0.52"/>` +
        `<dgm:constr type="w" for="ch" forName="childText" refType="w"/>` +
        `<dgm:constr type="h" for="ch" forName="childText" refType="primFontSz" refFor="ch" refForName="parentText" fact="0.46"/>` +
        `<dgm:constr type="h" for="ch" forName="parentText" op="equ"/>` +
        `<dgm:constr type="primFontSz" for="ch" forName="parentText" op="equ" val="65"/>` +
        `<dgm:constr type="primFontSz" for="ch" forName="childText" refType="primFontSz" refFor="ch" refForName="parentText" op="equ"/>` +
        `<dgm:constr type="h" for="ch" forName="spacer" refType="primFontSz" refFor="ch" refForName="parentText" fact="0.08"/>` +
        `</dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" for="ch" forName="parentText" val="5" fact="NaN" max="NaN"/></dgm:ruleLst>` +
        `<dgm:forEach name="Name0" axis="ch" ptType="node">` +
        textNode("parentText", "roundRect", `axis="self"`) +
        `<dgm:choose name="Name1"><dgm:if name="Name2" axis="ch" ptType="node" func="cnt" op="gte" val="1">` +
        `<dgm:layoutNode name="childText" styleLbl="revTx"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst>` +
        `<dgm:alg type="tx"><dgm:param type="stBulletLvl" val="1"/><dgm:param type="lnSpAfChP" val="20"/></dgm:alg>` +
        `<dgm:shape type="rect" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="des" ptType="node"/>` +
        `<dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.1"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.1"/>` +
        `<dgm:constr type="lMarg" refType="w" fact="0.09"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode>` +
        `</dgm:if><dgm:else name="Name3"/></dgm:choose>` +
        `<dgm:forEach name="Name4" axis="followSib" ptType="sibTrans" cnt="1">${SPACE_NODE("spacer")}</dgm:forEach>` +
        `</dgm:forEach></dgm:layoutNode>`
      );
    case "basicProcess":
      return (
        `<dgm:layoutNode name="Name0"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>` +
        `<dgm:alg type="lin"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/>` +
        `<dgm:constrLst>` +
        `<dgm:constr type="w" for="ch" forName="node" refType="w"/>` +
        `<dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/>` +
        `<dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.2"/>` +
        `<dgm:constr type="h" for="ch" forName="sibTrans" refType="h" refFor="ch" refForName="node" fact="0.35"/>` +
        `<dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>` +
        `<dgm:constr type="primFontSz" for="des" forName="connectorText" op="equ" val="55"/>` +
        `</dgm:constrLst><dgm:ruleLst/>` +
        `<dgm:forEach name="Name1" axis="ch" ptType="node">` +
        textNode("node", "roundRect", `axis="desOrSelf" ptType="node"`) +
        `<dgm:forEach name="Name2" axis="followSib" ptType="sibTrans" cnt="1">` +
        `<dgm:layoutNode name="sibTrans" styleLbl="sibTrans2D1"><dgm:alg type="conn"/>` +
        `<dgm:shape type="rightArrow" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/>` +
        `<dgm:constrLst><dgm:constr type="h" refType="w" fact="0.62"/><dgm:constr type="connDist"/>` +
        `<dgm:constr type="begPad" refType="connDist" fact="0.25"/><dgm:constr type="endPad" refType="connDist" fact="0.22"/></dgm:constrLst>` +
        `<dgm:ruleLst/>` +
        `<dgm:forEach name="Name3" axis="self" ptType="sibTrans">` +
        `<dgm:layoutNode name="connectorText"><dgm:alg type="tx"><dgm:param type="autoTxRot" val="grav"/></dgm:alg>` +
        `<dgm:shape r:blip="" hideGeom="1"><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/>` +
        `<dgm:constrLst><dgm:constr type="lMarg"/><dgm:constr type="rMarg"/><dgm:constr type="tMarg"/><dgm:constr type="bMarg"/></dgm:constrLst>` +
        `<dgm:ruleLst/></dgm:layoutNode></dgm:forEach>` +
        `</dgm:layoutNode></dgm:forEach></dgm:forEach></dgm:layoutNode>`
      );
    case "basicCycle":
      return (
        `<dgm:layoutNode name="cycle"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>` +
        `<dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg>` +
        `<dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/>` +
        `<dgm:constrLst>` +
        `<dgm:constr type="w" for="ch" forName="node" refType="w"/>` +
        `<dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node"/>` +
        `<dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.15"/>` +
        `<dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="0.15"/>` +
        `<dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>` +
        `</dgm:constrLst><dgm:ruleLst/>` +
        `<dgm:forEach name="Name0" axis="ch" ptType="node">` +
        textNode(
          "node",
          "ellipse",
          `axis="desOrSelf" ptType="node"`,
          `<dgm:constr type="h" refType="w"/>`,
        ) +
        `<dgm:forEach name="Name1" axis="followSib" ptType="sibTrans" cnt="1">` +
        `<dgm:layoutNode name="sibTrans" styleLbl="sibTrans2D1"><dgm:alg type="conn"><dgm:param type="srcNode" val="node"/>` +
        `<dgm:param type="dstNode" val="node"/></dgm:alg><dgm:shape type="rightArrow" r:blip=""><dgm:adjLst/></dgm:shape>` +
        `<dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="h" refType="w" fact="0.6"/></dgm:constrLst><dgm:ruleLst/>` +
        `</dgm:layoutNode></dgm:forEach></dgm:forEach></dgm:layoutNode>`
      );
    case "hierarchy":
      return (
        `<dgm:layoutNode name="hierChild1"><dgm:varLst><dgm:chPref val="1"/><dgm:dir/><dgm:animOne val="branch"/>` +
        `<dgm:animLvl val="lvl"/><dgm:resizeHandles/></dgm:varLst>` +
        `<dgm:alg type="hierChild"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/>` +
        `<dgm:constrLst><dgm:constr type="w" for="des" forName="node" refType="w"/>` +
        `<dgm:constr type="h" for="des" forName="node" refType="w" refFor="des" refForName="node" fact="0.635"/>` +
        `<dgm:constr type="sp" for="des" op="equ"/>` +
        `<dgm:constr type="primFontSz" for="des" forName="node" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/>` +
        `<dgm:forEach name="rootForEach" axis="ch" ptType="node">` +
        `<dgm:layoutNode name="hierRoot"><dgm:alg type="hierRoot"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape>` +
        `<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>` +
        textNode("node", "roundRect", `axis="self"`) +
        `<dgm:layoutNode name="hierChild2"><dgm:alg type="hierChild"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape>` +
        `<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>` +
        `<dgm:forEach name="childForEach" axis="ch" ptType="node">` +
        `<dgm:forEach name="lineForEach" axis="precedSib" ptType="parTrans" cnt="1">` +
        `<dgm:layoutNode name="parTrans" styleLbl="parChTrans1D2"><dgm:alg type="conn"><dgm:param type="dim" val="1D"/>` +
        `<dgm:param type="endSty" val="noArr"/><dgm:param type="connRout" val="bend"/></dgm:alg>` +
        `<dgm:shape type="conn" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/>` +
        `<dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach>` +
        `<dgm:forEach name="childRepeat" ref="rootForEach"/>` +
        `</dgm:forEach></dgm:layoutNode></dgm:layoutNode></dgm:forEach></dgm:layoutNode>`
      );
  }
}

export function layoutDefinitionXml(layout: SmartArtLayout): string {
  const def = SMARTART_LAYOUTS[layout];
  return (
    HEADER +
    `<dgm:layoutDef xmlns:dgm="${DGM}" xmlns:a="${A}" xmlns:r="${R}" uniqueId="${def.id}">` +
    `<dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="${def.category}" pri="1000"/></dgm:catLst>` +
    layoutBody(layout) +
    `</dgm:layoutDef>`
  );
}

// The style labels a definition styles, covering every label the layouts above use.
const STYLE_LABELS = [
  "node0",
  "node1",
  "lnNode1",
  "vennNode1",
  "alignNode1",
  "trAlignAcc1",
  "sibTrans2D1",
  "sibTrans1D1",
  "parChTrans1D1",
  "parChTrans1D2",
  "bgShp",
  "revTx",
] as const;

export function styleDefinitionXml(style: SmartArtStyle): string {
  const s = SMARTART_STYLES[style];
  const label = (name: string): string => {
    // Text-only labels and connectors keep thin lines whatever the style.
    const textOnly = name === "revTx";
    return (
      `<dgm:styleLbl name="${name}"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>` +
      `<dgm:sp3d/><dgm:txPr/><dgm:style>` +
      `<a:lnRef idx="${textOnly ? 0 : s.line}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>` +
      `<a:fillRef idx="${textOnly ? 0 : s.fill}"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>` +
      `<a:effectRef idx="${textOnly ? 0 : s.effect}"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>` +
      `<a:fontRef idx="minor"/></dgm:style></dgm:styleLbl>`
    );
  };
  return (
    HEADER +
    `<dgm:styleDef xmlns:dgm="${DGM}" xmlns:a="${A}" uniqueId="${styleUniqueId(style)}">` +
    `<dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>` +
    `<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>` +
    `<dgm:styleLbl name="node0"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>` +
    `<dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="${s.line}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>` +
    `<a:fillRef idx="${s.fill}"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="${s.effect}"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>` +
    `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl>` +
    STYLE_LABELS.filter((n) => n !== "node0")
      .map(label)
      .join("") +
    `</dgm:styleDef>`
  );
}

const schemeColors = (vals: readonly string[]): string =>
  vals.map((v) => `<a:schemeClr val="${v}"/>`).join("");

export function colorsDefinitionXml(colors: SmartArtColors): string {
  const def = SMARTART_COLORS[colors];
  const fills = schemeColors(def.fill);
  const method = def.fill.length > 1 ? "repeat" : "none";
  const label = (name: string): string => {
    const connector = name.startsWith("sibTrans") || name.startsWith("parChTrans");
    const textOnly = name === "revTx";
    const fill = textOnly ? "" : connector ? schemeColors([def.fill[0] ?? "accent1"]) : fills;
    const line =
      connector || textOnly ? schemeColors([def.fill[0] ?? "accent1"]) : schemeColors(["lt1"]);
    const text = textOnly ? schemeColors(["dk1"]) : schemeColors(["lt1"]);
    return (
      `<dgm:styleLbl name="${name}">` +
      `<dgm:fillClrLst meth="${method}">${fill}</dgm:fillClrLst>` +
      `<dgm:linClrLst meth="${method}">${line}</dgm:linClrLst>` +
      `<dgm:effectClrLst/><dgm:txLinClrLst/>` +
      `<dgm:txFillClrLst meth="repeat">${text}</dgm:txFillClrLst><dgm:txEffectClrLst/>` +
      `</dgm:styleLbl>`
    );
  };
  return (
    HEADER +
    `<dgm:colorsDef xmlns:dgm="${DGM}" xmlns:a="${A}" uniqueId="${colorsUniqueId(colors)}">` +
    `<dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="${def.category}" pri="11200"/></dgm:catLst>` +
    STYLE_LABELS.map(label).join("") +
    `</dgm:colorsDef>`
  );
}
