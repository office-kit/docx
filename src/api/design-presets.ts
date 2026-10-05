/**
 * Word's built-in theme building blocks (Design ▸ Themes / Colors / Fonts /
 * Effects) and document style sets, as data for the design functions.
 *
 * The colour and font values follow the schemes Word ships; effect schemes
 * and style sets are reproduced from Word's output where known and
 * approximated otherwise (see each list's comment).
 */

/** A theme colour scheme (`<a:clrScheme>`, ECMA-376 Part 1 §20.1.6.2), as hex RGB. */
export interface ThemeColorScheme {
  readonly name: string;
  readonly dk1: string;
  readonly lt1: string;
  readonly dk2: string;
  readonly lt2: string;
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
  readonly hlink: string;
  readonly folHlink: string;
}

/** A theme font scheme (`<a:fontScheme>`, §20.1.4.1.18): heading (major) and body (minor) fonts. */
export interface ThemeFontScheme {
  readonly name: string;
  /** Latin heading font. */
  readonly major: string;
  /** Latin body font. */
  readonly minor: string;
}

/**
 * A theme effect scheme: the `<a:fmtScheme>` (§20.1.4.1.14) fill, line,
 * effect and background style lists, as DrawingML markup.
 */
export interface ThemeEffectScheme {
  readonly name: string;
  /** Inner markup of `<a:fmtScheme>` (the four style lists). */
  readonly styles: string;
}

/** A complete theme: colours, fonts and effects under one name. */
export interface ThemeDefinition {
  readonly name: string;
  readonly colors: ThemeColorScheme;
  readonly fonts: ThemeFontScheme;
  readonly effects: ThemeEffectScheme;
}

/** The names of the twelve theme colour slots, in `<a:clrScheme>` order. */
export const THEME_COLOR_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

const COLORS_OFFICE: ThemeColorScheme = {
  name: "Office",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "0E2841",
  lt2: "E8E8E8",
  accent1: "156082",
  accent2: "E97132",
  accent3: "196B24",
  accent4: "0F9ED5",
  accent5: "A02B93",
  accent6: "4EA72E",
  hlink: "467886",
  folHlink: "96607D",
};
const COLORS_OFFICE_2013_2022: ThemeColorScheme = {
  name: "Office 2013 - 2022",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  hlink: "0563C1",
  folHlink: "954F72",
};
const COLORS_OFFICE_2007_2010: ThemeColorScheme = {
  name: "Office 2007 - 2010",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "1F497D",
  lt2: "EEECE1",
  accent1: "4F81BD",
  accent2: "C0504D",
  accent3: "9BBB59",
  accent4: "8064A2",
  accent5: "4BACC6",
  accent6: "F79646",
  hlink: "0000FF",
  folHlink: "800080",
};
const COLORS_GRAYSCALE: ThemeColorScheme = {
  name: "Grayscale",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "000000",
  lt2: "F8F8F8",
  accent1: "DDDDDD",
  accent2: "B2B2B2",
  accent3: "969696",
  accent4: "808080",
  accent5: "5F5F5F",
  accent6: "4D4D4D",
  hlink: "5F5F5F",
  folHlink: "919191",
};
const COLORS_BLUE_WARM: ThemeColorScheme = {
  name: "Blue Warm",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "242852",
  lt2: "ACCBF9",
  accent1: "4A66AC",
  accent2: "629DD1",
  accent3: "297FD5",
  accent4: "7F8FA9",
  accent5: "5AA2AE",
  accent6: "9D90A0",
  hlink: "9454C3",
  folHlink: "3EBBF0",
};
const COLORS_BLUE: ThemeColorScheme = {
  name: "Blue",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "17406D",
  lt2: "DBEFF9",
  accent1: "0F6FC6",
  accent2: "009DD9",
  accent3: "0BD0D9",
  accent4: "10CF9B",
  accent5: "7CCA62",
  accent6: "A5C249",
  hlink: "F49100",
  folHlink: "85DFD0",
};
const COLORS_BLUE_II: ThemeColorScheme = {
  name: "Blue II",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "335B74",
  lt2: "DFE3E5",
  accent1: "1CADE4",
  accent2: "2683C6",
  accent3: "27CED7",
  accent4: "42BA97",
  accent5: "3E8853",
  accent6: "62A39F",
  hlink: "6B9F25",
  folHlink: "B26B02",
};
const COLORS_BLUE_GREEN: ThemeColorScheme = {
  name: "Blue Green",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "373545",
  lt2: "CEDBE6",
  accent1: "3494BA",
  accent2: "58B6C0",
  accent3: "75BDA7",
  accent4: "7A8C8E",
  accent5: "84ACB6",
  accent6: "2683C6",
  hlink: "6B9F25",
  folHlink: "9F6715",
};
const COLORS_GREEN: ThemeColorScheme = {
  name: "Green",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "455F51",
  lt2: "E3DED1",
  accent1: "549E39",
  accent2: "8AB833",
  accent3: "C0CF3A",
  accent4: "029676",
  accent5: "4AB5C4",
  accent6: "0989B1",
  hlink: "6B9F25",
  folHlink: "BA6906",
};
const COLORS_GREEN_YELLOW: ThemeColorScheme = {
  name: "Green Yellow",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "455F51",
  lt2: "E3DED1",
  accent1: "99CB38",
  accent2: "63A537",
  accent3: "37A76F",
  accent4: "44C1A3",
  accent5: "4EB3CF",
  accent6: "51C3F9",
  hlink: "EE7B08",
  folHlink: "977B2D",
};
const COLORS_YELLOW: ThemeColorScheme = {
  name: "Yellow",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "39302A",
  lt2: "E5DEDB",
  accent1: "FFCA08",
  accent2: "F8931D",
  accent3: "CE8D3E",
  accent4: "EC7016",
  accent5: "E64823",
  accent6: "9C6A6A",
  hlink: "2998E3",
  folHlink: "7F723D",
};
const COLORS_YELLOW_ORANGE: ThemeColorScheme = {
  name: "Yellow Orange",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "4E3B30",
  lt2: "FBEEC9",
  accent1: "F0A22E",
  accent2: "A5644E",
  accent3: "B58B80",
  accent4: "C3986D",
  accent5: "A19574",
  accent6: "C17529",
  hlink: "AD1F1F",
  folHlink: "FFC42F",
};
const COLORS_ORANGE: ThemeColorScheme = {
  name: "Orange",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "637052",
  lt2: "CCDDEA",
  accent1: "E48312",
  accent2: "BD582C",
  accent3: "865640",
  accent4: "9B8357",
  accent5: "C2BC80",
  accent6: "94A088",
  hlink: "2998E3",
  folHlink: "8C8C8C",
};
const COLORS_ORANGE_RED: ThemeColorScheme = {
  name: "Orange Red",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "696464",
  lt2: "E9E5DC",
  accent1: "D34817",
  accent2: "9B2D1F",
  accent3: "A28E6A",
  accent4: "956251",
  accent5: "918485",
  accent6: "855D5D",
  hlink: "CC9900",
  folHlink: "96A9A9",
};
const COLORS_RED_ORANGE: ThemeColorScheme = {
  name: "Red Orange",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "505046",
  lt2: "EEECE1",
  accent1: "E84C22",
  accent2: "FFBD47",
  accent3: "B64926",
  accent4: "FF8427",
  accent5: "CC9900",
  accent6: "B22600",
  hlink: "CC9900",
  folHlink: "666699",
};
const COLORS_RED: ThemeColorScheme = {
  name: "Red",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "323232",
  lt2: "E5C243",
  accent1: "A5300F",
  accent2: "D55816",
  accent3: "E19825",
  accent4: "B19C7D",
  accent5: "7F5F52",
  accent6: "B27D49",
  hlink: "6B9F25",
  folHlink: "B26B02",
};
const COLORS_RED_VIOLET: ThemeColorScheme = {
  name: "Red Violet",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "454551",
  lt2: "D8D9DC",
  accent1: "E32D91",
  accent2: "C830CC",
  accent3: "4EA6DC",
  accent4: "4775E7",
  accent5: "8971E1",
  accent6: "D54773",
  hlink: "6B9F25",
  folHlink: "8C8C8C",
};
const COLORS_VIOLET: ThemeColorScheme = {
  name: "Violet",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "373545",
  lt2: "DCD8DC",
  accent1: "AD84C6",
  accent2: "8784C7",
  accent3: "5D739A",
  accent4: "6997AF",
  accent5: "84ACB6",
  accent6: "6F8183",
  hlink: "69A020",
  folHlink: "8C8C8C",
};
const COLORS_VIOLET_II: ThemeColorScheme = {
  name: "Violet II",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "632E62",
  lt2: "EAE5EB",
  accent1: "92278F",
  accent2: "9B57D3",
  accent3: "755DD9",
  accent4: "665EB8",
  accent5: "45A5ED",
  accent6: "5982DB",
  hlink: "0066FF",
  folHlink: "666699",
};
const COLORS_MEDIAN: ThemeColorScheme = {
  name: "Median",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "775F55",
  lt2: "EBDDC3",
  accent1: "94B6D2",
  accent2: "DD8047",
  accent3: "A5AB81",
  accent4: "D8B25C",
  accent5: "7BA79D",
  accent6: "968C8C",
  hlink: "F7B615",
  folHlink: "704404",
};
const COLORS_PAPER: ThemeColorScheme = {
  name: "Paper",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "444D26",
  lt2: "FEFAC9",
  accent1: "A5B592",
  accent2: "F3A447",
  accent3: "E7BC29",
  accent4: "D092A7",
  accent5: "9C85C0",
  accent6: "809EC2",
  hlink: "8E58B6",
  folHlink: "7F6F6F",
};
const COLORS_MARQUEE: ThemeColorScheme = {
  name: "Marquee",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "5E5E5E",
  lt2: "DDDDDD",
  accent1: "418AB3",
  accent2: "A6B727",
  accent3: "F69200",
  accent4: "838383",
  accent5: "FEC306",
  accent6: "DF5327",
  hlink: "F59E00",
  folHlink: "B2B2B2",
};
const COLORS_SLIPSTREAM: ThemeColorScheme = {
  name: "Slipstream",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "212745",
  lt2: "B4DCFA",
  accent1: "4E67C8",
  accent2: "5ECCF3",
  accent3: "A7EA52",
  accent4: "5DCEAF",
  accent5: "FF8021",
  accent6: "F14124",
  hlink: "56C7AA",
  folHlink: "59A8D1",
};
const COLORS_ASPECT: ThemeColorScheme = {
  name: "Aspect",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "323232",
  lt2: "E3DED1",
  accent1: "F07F09",
  accent2: "9F2936",
  accent3: "1B587C",
  accent4: "4E8542",
  accent5: "604878",
  accent6: "C19859",
  hlink: "6B9F25",
  folHlink: "B26B02",
};

/** Design ▸ Colors, in Word's menu order. */
export const THEME_COLOR_SCHEMES: readonly ThemeColorScheme[] = [
  COLORS_OFFICE,
  COLORS_OFFICE_2013_2022,
  COLORS_OFFICE_2007_2010,
  COLORS_GRAYSCALE,
  COLORS_BLUE_WARM,
  COLORS_BLUE,
  COLORS_BLUE_II,
  COLORS_BLUE_GREEN,
  COLORS_GREEN,
  COLORS_GREEN_YELLOW,
  COLORS_YELLOW,
  COLORS_YELLOW_ORANGE,
  COLORS_ORANGE,
  COLORS_ORANGE_RED,
  COLORS_RED_ORANGE,
  COLORS_RED,
  COLORS_RED_VIOLET,
  COLORS_VIOLET,
  COLORS_VIOLET_II,
  COLORS_MEDIAN,
  COLORS_PAPER,
  COLORS_MARQUEE,
  COLORS_SLIPSTREAM,
  COLORS_ASPECT,
];

/** Design ▸ Fonts, in Word's menu order (heading font, body font). */
export const THEME_FONT_SCHEMES: readonly ThemeFontScheme[] = [
  { name: "Office", major: "Aptos Display", minor: "Aptos" },
  { name: "Office 2013 - 2022", major: "Calibri Light", minor: "Calibri" },
  { name: "Office 2007 - 2010", major: "Cambria", minor: "Calibri" },
  { name: "Calibri", major: "Calibri", minor: "Calibri" },
  { name: "Arial", major: "Arial", minor: "Arial" },
  { name: "Corbel", major: "Corbel", minor: "Corbel" },
  { name: "Candara", major: "Candara", minor: "Candara" },
  { name: "Franklin Gothic", major: "Franklin Gothic Medium", minor: "Franklin Gothic Book" },
  { name: "Century Gothic", major: "Century Gothic", minor: "Century Gothic" },
  { name: "Tw Cen MT", major: "Tw Cen MT", minor: "Tw Cen MT" },
  { name: "Cambria", major: "Cambria", minor: "Cambria" },
  { name: "Garamond", major: "Garamond", minor: "Garamond" },
  { name: "Times New Roman", major: "Times New Roman", minor: "Times New Roman" },
  { name: "Trebuchet MS", major: "Trebuchet MS", minor: "Trebuchet MS" },
  { name: "Verdana", major: "Verdana", minor: "Verdana" },
  { name: "Georgia", major: "Georgia", minor: "Georgia" },
  { name: "Tw Cen MT-Rockwell", major: "Tw Cen MT Condensed", minor: "Tw Cen MT" },
  { name: "Gill Sans MT", major: "Gill Sans MT", minor: "Gill Sans MT" },
  { name: "Rockwell", major: "Rockwell Condensed", minor: "Rockwell" },
  { name: "Palatino Linotype", major: "Palatino Linotype", minor: "Palatino Linotype" },
  { name: "Century Schoolbook", major: "Century Schoolbook", minor: "Century Schoolbook" },
  { name: "Consolas-Verdana", major: "Consolas", minor: "Verdana" },
  { name: "Calibri-Cambria", major: "Calibri", minor: "Cambria" },
];

// Effect schemes are written out in full (rather than built from helper
// calls) so bundlers can drop them when unused. "Office" and "Office 2007 -
// 2010" are Word's own fmtSchemes; "Subtle Solids" and "Top Shadow" follow
// the look of Word's schemes of those names without matching them byte for byte.
const EFFECTS_OFFICE: ThemeEffectScheme = {
  name: "Office",
  styles:
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst>',
};
const EFFECTS_OFFICE_2013_2022: ThemeEffectScheme = {
  name: "Office 2013 - 2022",
  styles:
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst>',
};
const EFFECTS_OFFICE_2007_2010: ThemeEffectScheme = {
  name: "Office 2007 - 2010",
  styles:
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst>',
};
const EFFECTS_SUBTLE_SOLIDS: ThemeEffectScheme = {
  name: "Subtle Solids",
  styles:
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="90000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>',
};
const EFFECTS_TOP_SHADOW: ThemeEffectScheme = {
  name: "Top Shadow",
  styles:
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="38100" dist="12700" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="63500" dist="25400" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst>',
};

/** Design ▸ Effects. */
export const THEME_EFFECT_SCHEMES: readonly ThemeEffectScheme[] = [
  EFFECTS_OFFICE,
  EFFECTS_OFFICE_2013_2022,
  EFFECTS_OFFICE_2007_2010,
  EFFECTS_SUBTLE_SOLIDS,
  EFFECTS_TOP_SHADOW,
];

/**
 * Design ▸ Themes: Office plus a selection of Word's built-in themes, with
 * Word's colours and fonts and the closest scheme in {@link THEME_EFFECT_SCHEMES}.
 */
export const THEMES: readonly ThemeDefinition[] = [
  {
    name: "Office Theme",
    colors: COLORS_OFFICE,
    fonts: {
      name: "Office",
      major: "Aptos Display",
      minor: "Aptos",
    },
    effects: EFFECTS_OFFICE,
  },
  {
    name: "Office 2013 - 2022 Theme",
    colors: COLORS_OFFICE_2013_2022,
    fonts: {
      name: "Office 2013 - 2022",
      major: "Calibri Light",
      minor: "Calibri",
    },
    effects: EFFECTS_OFFICE_2013_2022,
  },
  {
    name: "Facet",
    colors: {
      name: "Facet",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "2C3C43",
      lt2: "EBEBEB",
      accent1: "90C226",
      accent2: "54A021",
      accent3: "E6B91E",
      accent4: "E76618",
      accent5: "C42F1A",
      accent6: "918655",
      hlink: "99CA3C",
      folHlink: "B9D181",
    },
    fonts: {
      name: "Trebuchet MS",
      major: "Trebuchet MS",
      minor: "Trebuchet MS",
    },
    effects: EFFECTS_SUBTLE_SOLIDS,
  },
  {
    name: "Integral",
    colors: {
      name: "Integral",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "335B74",
      lt2: "DFE3E5",
      accent1: "1CADE4",
      accent2: "2683C6",
      accent3: "27CED7",
      accent4: "42BA97",
      accent5: "3E8853",
      accent6: "62A39F",
      hlink: "6B9F25",
      folHlink: "B26B02",
    },
    fonts: {
      name: "Integral",
      major: "Tw Cen MT Condensed",
      minor: "Tw Cen MT",
    },
    effects: EFFECTS_OFFICE_2013_2022,
  },
  {
    name: "Ion Boardroom",
    colors: {
      name: "Ion Boardroom",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "3D3D3D",
      lt2: "EBEBEB",
      accent1: "B01513",
      accent2: "EA6312",
      accent3: "E6B729",
      accent4: "6AAC90",
      accent5: "5F9C9D",
      accent6: "9E5E9B",
      hlink: "58C1BA",
      folHlink: "9DFFCB",
    },
    fonts: {
      name: "Century Gothic",
      major: "Century Gothic",
      minor: "Century Gothic",
    },
    effects: EFFECTS_TOP_SHADOW,
  },
  {
    name: "Organic",
    colors: {
      name: "Organic",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "212121",
      lt2: "DDDDDD",
      accent1: "83992A",
      accent2: "3C9770",
      accent3: "44709D",
      accent4: "A23C33",
      accent5: "D97828",
      accent6: "DEB340",
      hlink: "A8BF4D",
      folHlink: "B4CA80",
    },
    fonts: {
      name: "Garamond",
      major: "Garamond",
      minor: "Garamond",
    },
    effects: EFFECTS_SUBTLE_SOLIDS,
  },
  {
    name: "Retrospect",
    colors: {
      name: "Retrospect",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "637052",
      lt2: "CCDDEA",
      accent1: "E48312",
      accent2: "BD582C",
      accent3: "865640",
      accent4: "9B8357",
      accent5: "C2BC80",
      accent6: "94A088",
      hlink: "2998E3",
      folHlink: "8C8C8C",
    },
    fonts: {
      name: "Office 2013 - 2022",
      major: "Calibri Light",
      minor: "Calibri",
    },
    effects: EFFECTS_OFFICE_2013_2022,
  },
  {
    name: "Slice",
    colors: {
      name: "Slice",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "146194",
      lt2: "48FFD5",
      accent1: "052F61",
      accent2: "A50E82",
      accent3: "14967C",
      accent4: "6A9E1F",
      accent5: "E87D37",
      accent6: "C62324",
      hlink: "0D2E46",
      folHlink: "356A95",
    },
    fonts: {
      name: "Century Gothic",
      major: "Century Gothic",
      minor: "Century Gothic",
    },
    effects: EFFECTS_SUBTLE_SOLIDS,
  },
  {
    name: "Wisp",
    colors: {
      name: "Wisp",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "4E3B30",
      lt2: "FBEEC9",
      accent1: "A53010",
      accent2: "DE7E18",
      accent3: "9F8351",
      accent4: "728653",
      accent5: "92AA4C",
      accent6: "6AAC91",
      hlink: "FB4A18",
      folHlink: "FB9318",
    },
    fonts: {
      name: "Century Gothic",
      major: "Century Gothic",
      minor: "Century Gothic",
    },
    effects: EFFECTS_TOP_SHADOW,
  },
  {
    name: "Banded",
    colors: {
      name: "Banded",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "3F3F3F",
      lt2: "E8E8E8",
      accent1: "FFC000",
      accent2: "A5D028",
      accent3: "08CC78",
      accent4: "F24099",
      accent5: "828288",
      accent6: "F56617",
      hlink: "005DBA",
      folHlink: "6C606A",
    },
    fonts: {
      name: "Corbel",
      major: "Corbel",
      minor: "Corbel",
    },
    effects: EFFECTS_SUBTLE_SOLIDS,
  },
  {
    name: "Gallery",
    colors: {
      name: "Gallery",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "454545",
      lt2: "DFDFDF",
      accent1: "B71E42",
      accent2: "DE478E",
      accent3: "BC72F0",
      accent4: "795FAF",
      accent5: "586EA6",
      accent6: "6892A0",
      hlink: "FF3399",
      folHlink: "A8A8A8",
    },
    fonts: {
      name: "Gill Sans MT",
      major: "Gill Sans MT",
      minor: "Gill Sans MT",
    },
    effects: EFFECTS_SUBTLE_SOLIDS,
  },
  {
    name: "Dividend",
    colors: {
      name: "Dividend",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "3A3A3A",
      lt2: "DEDEDE",
      accent1: "4D1434",
      accent2: "903163",
      accent3: "B2324B",
      accent4: "969FA7",
      accent5: "66B1CE",
      accent6: "40619D",
      hlink: "828288",
      folHlink: "A5A5A5",
    },
    fonts: {
      name: "Gill Sans MT",
      major: "Gill Sans MT",
      minor: "Gill Sans MT",
    },
    effects: EFFECTS_OFFICE_2013_2022,
  },
  {
    name: "View",
    colors: {
      name: "View",
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "212121",
      lt2: "EBE9E5",
      accent1: "6F6F74",
      accent2: "A7B789",
      accent3: "BEAE98",
      accent4: "92A9B9",
      accent5: "9C8265",
      accent6: "8D6974",
      hlink: "67AABF",
      folHlink: "B1B5AB",
    },
    fonts: {
      name: "Century Schoolbook",
      major: "Century Schoolbook",
      minor: "Century Schoolbook",
    },
    effects: EFFECTS_SUBTLE_SOLIDS,
  },
];

/** A theme colour a style can reference (`w:themeColor`, §17.18.97). */
export type StyleThemeColor =
  | "text1"
  | "text2"
  | "background1"
  | "background2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6";

/** A colour as Word writes it on styles: RGB plus the theme slot it came from. */
export interface StyleColor {
  readonly rgb: string;
  readonly theme?: StyleThemeColor;
  /** `w:themeShade` / `w:themeTint`, hex 00–FF. */
  readonly shade?: string;
  readonly tint?: string;
}

/** A one-sided paragraph border on a style-set style. */
export interface StyleSetBorder {
  readonly style: "single" | "double" | "thick" | "thinThickSmallGap";
  /** Eighths of a point. */
  readonly size: number;
  /** Points between text and border. */
  readonly space: number;
  readonly color: StyleColor;
}

/** The formatting a style set gives one paragraph style. */
export interface StyleSetStyle {
  /** Theme font role; Normal-based styles default to the body font. */
  readonly font?: "major" | "minor";
  readonly sizeHalfPoints?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly caps?: boolean;
  readonly smallCaps?: boolean;
  readonly color?: StyleColor;
  /** Character spacing in twips (`w:spacing` in rPr). */
  readonly characterSpacing?: number;
  readonly kern?: number;
  readonly alignment?: "left" | "center" | "right";
  readonly before?: number;
  readonly after?: number;
  /** 240ths of a line. */
  readonly line?: number;
  readonly contextualSpacing?: boolean;
  readonly borderTop?: StyleSetBorder;
  readonly borderBottom?: StyleSetBorder;
  readonly shading?: StyleColor;
}

/** The styles a style set defines, by style id. */
export type StyleSetStyleId =
  | "Normal"
  | "Title"
  | "Subtitle"
  | "Heading1"
  | "Heading2"
  | "Heading3";

/** A Design ▸ Document Formatting style set. */
export interface StyleSetDefinition {
  readonly name: string;
  /** Body text size (document default), half-points. */
  readonly bodySizeHalfPoints: number;
  /** Default paragraph spacing after, twips. */
  readonly after: number;
  /** Default line spacing, 240ths of a line. */
  readonly line: number;
  readonly styles: Readonly<Partial<Record<StyleSetStyleId, StyleSetStyle>>>;
}

const ACCENT1_DARK: StyleColor = { rgb: "0F4761", theme: "accent1", shade: "BF" };
const ACCENT1: StyleColor = { rgb: "156082", theme: "accent1" };
const ACCENT2: StyleColor = { rgb: "E97132", theme: "accent2" };
const TEXT1: StyleColor = { rgb: "000000", theme: "text1" };
const TEXT1_LIGHT: StyleColor = { rgb: "595959", theme: "text1", tint: "A6" };
const TEXT2: StyleColor = { rgb: "0E2841", theme: "text2" };
const GRAY_FILL: StyleColor = { rgb: "D9D9D9", theme: "background1", shade: "D9" };
const ACCENT1_FILL: StyleColor = { rgb: "156082", theme: "accent1" };
const ACCENT1_PALE: StyleColor = { rgb: "C1E4F5", theme: "accent1", tint: "33" };

/**
 * Design ▸ Document Formatting. "Default" is the Word 365 Normal template's
 * styles; the others recreate the look of Word's style sets of the same
 * names (heading fonts, caps, rules and shading) for the styles listed in
 * {@link StyleSetStyleId}.
 */
export const STYLE_SETS: readonly StyleSetDefinition[] = [
  {
    name: "Default",
    bodySizeHalfPoints: 24,
    after: 160,
    line: 278,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 56,
        kern: 28,
        characterSpacing: -10,
        after: 80,
        line: 240,
        contextualSpacing: true,
      },
      Subtitle: { sizeHalfPoints: 28, color: TEXT1_LIGHT, characterSpacing: 15, after: 160 },
      Heading1: { font: "major", sizeHalfPoints: 40, color: ACCENT1_DARK, before: 360, after: 80 },
      Heading2: { font: "major", sizeHalfPoints: 32, color: ACCENT1_DARK, before: 160, after: 80 },
      Heading3: { sizeHalfPoints: 28, color: ACCENT1_DARK, before: 160, after: 80 },
    },
  },
  {
    name: "Basic (Elegant)",
    bodySizeHalfPoints: 20,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 52,
        caps: true,
        characterSpacing: 20,
        alignment: "right",
        after: 240,
        line: 240,
        contextualSpacing: true,
      },
      Subtitle: {
        caps: true,
        color: TEXT1_LIGHT,
        characterSpacing: 20,
        alignment: "right",
        after: 480,
      },
      Heading1: {
        font: "major",
        sizeHalfPoints: 24,
        caps: true,
        characterSpacing: 20,
        before: 480,
        after: 120,
        borderBottom: { style: "single", size: 4, space: 1, color: ACCENT1 },
      },
      Heading2: {
        font: "major",
        sizeHalfPoints: 22,
        caps: true,
        characterSpacing: 20,
        before: 240,
        after: 80,
      },
      Heading3: {
        font: "major",
        caps: true,
        color: TEXT1_LIGHT,
        characterSpacing: 20,
        before: 240,
        after: 80,
      },
    },
  },
  {
    name: "Basic (Simple)",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 48,
        caps: true,
        color: TEXT1,
        characterSpacing: 10,
        after: 200,
        contextualSpacing: true,
      },
      Subtitle: { caps: true, color: TEXT1_LIGHT, characterSpacing: 10, after: 240 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 26,
        caps: true,
        color: ACCENT1_DARK,
        characterSpacing: 10,
        before: 400,
        after: 80,
      },
      Heading2: {
        font: "major",
        sizeHalfPoints: 22,
        caps: true,
        color: ACCENT1_DARK,
        characterSpacing: 10,
        before: 240,
        after: 40,
      },
      Heading3: { font: "major", caps: true, color: TEXT1_LIGHT, before: 240, after: 40 },
    },
  },
  {
    name: "Basic (Stylish)",
    bodySizeHalfPoints: 20,
    after: 160,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 40,
        italic: true,
        color: { rgb: "FFFFFF", theme: "background1" },
        alignment: "center",
        after: 160,
        shading: ACCENT2,
      },
      Subtitle: { italic: true, color: ACCENT2, alignment: "center", after: 240 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 24,
        color: ACCENT2,
        before: 360,
        after: 120,
        borderBottom: { style: "single", size: 8, space: 1, color: ACCENT2 },
      },
      Heading2: { font: "major", sizeHalfPoints: 22, color: ACCENT2, before: 240, after: 80 },
      Heading3: { font: "major", italic: true, color: ACCENT2, before: 200, after: 80 },
    },
  },
  {
    name: "Black & White (Capitals)",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 52,
        bold: true,
        caps: true,
        color: TEXT1,
        characterSpacing: 40,
        alignment: "center",
        after: 300,
        contextualSpacing: true,
      },
      Subtitle: { caps: true, color: TEXT1, characterSpacing: 20, alignment: "center", after: 360 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 28,
        bold: true,
        caps: true,
        color: TEXT1,
        characterSpacing: 20,
        before: 480,
        after: 120,
      },
      Heading2: {
        font: "major",
        sizeHalfPoints: 24,
        bold: true,
        caps: true,
        color: TEXT1,
        before: 240,
        after: 80,
      },
      Heading3: { font: "major", caps: true, color: TEXT1, before: 240, after: 80 },
    },
  },
  {
    name: "Black & White (Classic)",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 56,
        bold: true,
        italic: true,
        color: TEXT1,
        after: 300,
        contextualSpacing: true,
      },
      Subtitle: { italic: true, color: TEXT1, after: 240 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 28,
        bold: true,
        italic: true,
        color: TEXT1,
        before: 480,
        after: 120,
      },
      Heading2: {
        font: "major",
        sizeHalfPoints: 24,
        bold: true,
        color: TEXT1,
        before: 240,
        after: 80,
      },
      Heading3: { font: "major", bold: true, italic: true, color: TEXT1, before: 240, after: 80 },
    },
  },
  {
    name: "Black & White (Word 2013)",
    bodySizeHalfPoints: 22,
    after: 160,
    line: 259,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 56,
        kern: 28,
        color: TEXT1,
        characterSpacing: -10,
        line: 240,
        contextualSpacing: true,
      },
      Subtitle: { color: TEXT1_LIGHT, characterSpacing: 15, after: 160 },
      Heading1: { font: "major", sizeHalfPoints: 32, color: TEXT1, before: 240 },
      Heading2: { font: "major", sizeHalfPoints: 26, color: TEXT1, before: 40 },
      Heading3: { font: "major", sizeHalfPoints: 24, color: TEXT1_LIGHT, before: 40 },
    },
  },
  {
    name: "Casual",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 300,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 64,
        color: ACCENT1,
        after: 200,
        contextualSpacing: true,
      },
      Subtitle: { sizeHalfPoints: 28, color: ACCENT2, after: 280 },
      Heading1: { font: "major", sizeHalfPoints: 36, color: ACCENT1, before: 360, after: 120 },
      Heading2: { font: "major", sizeHalfPoints: 28, color: ACCENT2, before: 240, after: 80 },
      Heading3: { font: "major", sizeHalfPoints: 24, color: ACCENT1, before: 200, after: 80 },
    },
  },
  {
    name: "Centered",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 44,
        caps: true,
        color: TEXT2,
        characterSpacing: 30,
        alignment: "center",
        after: 240,
        contextualSpacing: true,
      },
      Subtitle: { caps: true, color: TEXT1_LIGHT, alignment: "center", after: 360 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 22,
        caps: true,
        color: { rgb: "FFFFFF", theme: "background1" },
        characterSpacing: 20,
        alignment: "center",
        before: 360,
        after: 120,
        shading: { rgb: "7F7F7F", theme: "text1", tint: "80" },
      },
      Heading2: {
        font: "major",
        caps: true,
        color: TEXT2,
        alignment: "center",
        before: 240,
        after: 80,
        shading: GRAY_FILL,
      },
      Heading3: {
        font: "major",
        caps: true,
        color: TEXT2,
        alignment: "center",
        before: 240,
        after: 80,
      },
    },
  },
  {
    name: "Lines (Distinctive)",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 52,
        color: TEXT1,
        after: 240,
        contextualSpacing: true,
        borderBottom: { style: "thinThickSmallGap", size: 24, space: 4, color: ACCENT1 },
      },
      Subtitle: { italic: true, color: TEXT1_LIGHT, after: 240 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 32,
        bold: true,
        color: ACCENT1_DARK,
        before: 480,
        after: 120,
        borderBottom: { style: "thick", size: 12, space: 1, color: ACCENT1 },
      },
      Heading2: {
        font: "major",
        sizeHalfPoints: 26,
        bold: true,
        color: ACCENT1_DARK,
        before: 240,
        after: 80,
        borderBottom: { style: "single", size: 4, space: 1, color: ACCENT1 },
      },
      Heading3: { font: "major", bold: true, color: ACCENT1, before: 240, after: 80 },
    },
  },
  {
    name: "Lines (Simple)",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 48,
        color: TEXT1,
        after: 240,
        contextualSpacing: true,
        borderBottom: { style: "single", size: 6, space: 4, color: ACCENT1 },
      },
      Subtitle: { color: TEXT1_LIGHT, after: 240 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 28,
        color: ACCENT1_DARK,
        before: 400,
        after: 120,
        borderBottom: { style: "single", size: 4, space: 1, color: ACCENT1 },
      },
      Heading2: { font: "major", sizeHalfPoints: 24, color: ACCENT1_DARK, before: 240, after: 80 },
      Heading3: { font: "major", color: ACCENT1_DARK, before: 240, after: 80 },
    },
  },
  {
    name: "Lines (Stylish)",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 48,
        italic: true,
        color: TEXT2,
        alignment: "center",
        after: 240,
        contextualSpacing: true,
        borderBottom: {
          style: "single",
          size: 12,
          space: 6,
          color: { rgb: "4EA72E", theme: "accent6" },
        },
      },
      Subtitle: { italic: true, color: TEXT1_LIGHT, alignment: "center", after: 240 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 28,
        italic: true,
        color: TEXT2,
        before: 400,
        after: 120,
        borderTop: {
          style: "single",
          size: 4,
          space: 4,
          color: { rgb: "4EA72E", theme: "accent6" },
        },
      },
      Heading2: {
        font: "major",
        sizeHalfPoints: 24,
        italic: true,
        color: TEXT2,
        before: 240,
        after: 80,
      },
      Heading3: { font: "major", italic: true, color: TEXT1_LIGHT, before: 240, after: 80 },
    },
  },
  {
    name: "Minimalist",
    bodySizeHalfPoints: 21,
    after: 240,
    line: 288,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 44,
        color: TEXT1_LIGHT,
        characterSpacing: 10,
        after: 360,
        contextualSpacing: true,
      },
      Subtitle: { color: TEXT1_LIGHT, after: 360 },
      Heading1: { font: "major", sizeHalfPoints: 26, color: TEXT1, before: 480, after: 160 },
      Heading2: { font: "major", sizeHalfPoints: 22, color: TEXT1_LIGHT, before: 320, after: 120 },
      Heading3: { font: "major", color: TEXT1_LIGHT, before: 240, after: 80 },
    },
  },
  {
    name: "Shaded",
    bodySizeHalfPoints: 22,
    after: 200,
    line: 276,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 48,
        color: { rgb: "FFFFFF", theme: "background1" },
        after: 200,
        shading: ACCENT1_FILL,
      },
      Subtitle: { color: ACCENT1_DARK, after: 240 },
      Heading1: {
        font: "major",
        sizeHalfPoints: 26,
        bold: true,
        color: { rgb: "FFFFFF", theme: "background1" },
        before: 360,
        after: 120,
        shading: ACCENT1_FILL,
      },
      Heading2: {
        font: "major",
        sizeHalfPoints: 24,
        bold: true,
        color: ACCENT1_DARK,
        before: 240,
        after: 80,
        shading: ACCENT1_PALE,
      },
      Heading3: { font: "major", bold: true, color: ACCENT1_DARK, before: 240, after: 80 },
    },
  },
  {
    name: "Title",
    bodySizeHalfPoints: 22,
    after: 160,
    line: 259,
    styles: {
      Title: {
        font: "major",
        sizeHalfPoints: 72,
        color: ACCENT1_DARK,
        kern: 28,
        after: 160,
        contextualSpacing: true,
      },
      Subtitle: { sizeHalfPoints: 32, color: TEXT1_LIGHT, after: 240 },
      Heading1: { font: "major", sizeHalfPoints: 32, color: ACCENT1_DARK, before: 360, after: 80 },
      Heading2: { font: "major", sizeHalfPoints: 26, color: ACCENT1_DARK, before: 160, after: 40 },
      Heading3: { font: "major", sizeHalfPoints: 24, color: ACCENT1, before: 160, after: 40 },
    },
  },
];

/** A Design ▸ Paragraph Spacing preset: default paragraph spacing, twips / 240ths of a line. */
export interface ParagraphSpacingPreset {
  readonly name: string;
  readonly before: number;
  readonly after: number;
  readonly line: number;
}

/** Design ▸ Paragraph Spacing, as Word's tooltips list them. */
export const PARAGRAPH_SPACING_PRESETS: readonly ParagraphSpacingPreset[] = [
  { name: "No Paragraph Space", before: 0, after: 0, line: 240 },
  { name: "Compact", before: 0, after: 80, line: 240 },
  { name: "Tight", before: 0, after: 120, line: 276 },
  { name: "Open", before: 0, after: 200, line: 276 },
  { name: "Relaxed", before: 0, after: 120, line: 360 },
  { name: "Double", before: 0, after: 160, line: 480 },
];
