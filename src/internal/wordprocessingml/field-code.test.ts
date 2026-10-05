import { describe, expect, it } from "vitest";
import {
  applyFormatSwitch,
  evaluateComparison,
  evaluateFormula,
  formatNumericPicture,
  parseFieldInstruction,
  quoteFieldArgument,
} from "./field-code.js";

describe("parseFieldInstruction", () => {
  it("splits type, quoted arguments and switches", () => {
    expect(parseFieldInstruction(' DATE \\@ "dddd, MMMM d" \\* MERGEFORMAT ')).toEqual({
      type: "DATE",
      args: [],
      switches: [
        { name: "@", arg: "dddd, MMMM d" },
        { name: "*", arg: "MERGEFORMAT" },
      ],
    });
    expect(parseFieldInstruction('HYPERLINK "https://a.b/" \\l "top" \\o "Tip \\"x\\""')).toEqual({
      type: "HYPERLINK",
      args: ["https://a.b/"],
      switches: [
        { name: "l", arg: "top" },
        { name: "o", arg: 'Tip "x"' },
      ],
    });
    expect(parseFieldInstruction("REF _Ref1 \\h \\p")).toEqual({
      type: "REF",
      args: ["_Ref1"],
      switches: [{ name: "h" }, { name: "p" }],
    });
    expect(parseFieldInstruction("=SUM(1,2) \\# 0.0").args).toEqual(["SUM(1,2)"]);
  });

  it("quotes arguments that need it", () => {
    expect(quoteFieldArgument("Title")).toBe("Title");
    expect(quoteFieldArgument('a "b"')).toBe('"a \\"b\\""');
  });
});

describe("format switches", () => {
  it.each([
    ["7", "ROMAN", "VII"],
    ["7", "roman", "vii"],
    ["28", "ALPHABETIC", "BB"],
    ["3", "alphabetic", "c"],
    ["22", "Ordinal", "22nd"],
    ["113", "Ordinal", "113th"],
    ["42", "CardText", "forty-two"],
    ["21", "OrdText", "twenty-first"],
    ["255", "Hex", "FF"],
    ["hello world", "Caps", "Hello World"],
    ["hello", "FirstCap", "Hello"],
    ["Hi", "Upper", "HI"],
  ])("%s \\* %s → %s", (value, format, expected) => {
    expect(applyFormatSwitch(value, format)).toBe(expected);
  });

  it("applies numeric pictures", () => {
    expect(formatNumericPicture(1234.5, "#,##0.00")).toBe("1,234.50");
    expect(formatNumericPicture(0.5, "0.0")).toBe("0.5");
    expect(formatNumericPicture(7, "00")).toBe("07");
    expect(formatNumericPicture(-3, "0;(0)")).toBe("(3)");
  });
});

describe("formulas and comparisons", () => {
  it("evaluates arithmetic, functions and bookmarks", () => {
    expect(evaluateFormula("2+3*4")).toBe(14);
    expect(evaluateFormula("(2+3)*4^2")).toBe(80);
    expect(evaluateFormula("SUM(1,2,3) + MAX(4;9)")).toBe(15);
    expect(evaluateFormula("50%*10")).toBe(5);
    expect(evaluateFormula("Price*2", (n) => (n === "PRICE" ? 21 : undefined))).toBe(42);
    expect(evaluateFormula("3>2")).toBe(1);
    expect(() => evaluateFormula("1/0")).toThrow("!Zero Divide");
  });

  it("compares numbers numerically and text with wildcards", () => {
    expect(evaluateComparison("10", ">", "9")).toBe(true);
    expect(evaluateComparison("apple", "=", "a*e")).toBe(true);
    expect(evaluateComparison("apple", "<>", "b?")).toBe(true);
  });
});
