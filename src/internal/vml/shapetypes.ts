/**
 * VML shape type definitions (`v:shapetype`, Part 4 §19.1.2.20) that Word
 * needs spelled out: for a text path (WordArt, text watermarks) Word draws
 * nothing unless the shape refers to the definition with its formulas.
 */

import { parseXml, type XmlElement } from "../xml/index.js";
import { OFFICE_NS, VML_NS } from "./namespaces.js";

/** Shape type 136, "plain text" WordArt; shapes refer to it as `type="#_x0000_t136"`. */
export const TEXT_PATH_SHAPETYPE_XML =
  '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">' +
  '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>' +
  '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/>' +
  '<v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>' +
  '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>' +
  '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>' +
  '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>' +
  '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';

/** The id a shape's `type` attribute names, `#` included. */
export const TEXT_PATH_SHAPETYPE_REF = "#_x0000_t136";

/** {@link TEXT_PATH_SHAPETYPE_XML} as a fresh element tree. */
export function textPathShapetype(): XmlElement {
  const wrapper = parseXml(
    `<wrap xmlns:v="${VML_NS}" xmlns:o="${OFFICE_NS}">${TEXT_PATH_SHAPETYPE_XML}</wrap>`,
  ).root;
  const shapetype = wrapper.children.find((c): c is XmlElement => c.kind === "element");
  if (!shapetype) throw new Error("The text path shape type did not parse.");
  return shapetype;
}
