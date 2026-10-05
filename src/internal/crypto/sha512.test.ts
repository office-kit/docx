import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import { sha512 } from "./sha512.js";

const hex = (bytes: Uint8Array): string => Buffer.from(bytes).toString("hex");

describe("sha512", () => {
  it("matches the FIPS 180-2 'abc' test vector", () => {
    expect(hex(sha512(new TextEncoder().encode("abc")))).toBe(
      "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a" +
        "2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f",
    );
  });

  it("matches Node's digest across block-padding boundaries", () => {
    for (const length of [0, 1, 55, 111, 112, 127, 128, 129, 300]) {
      const data = Uint8Array.from({ length }, (_, i) => (i * 37 + 11) & 0xff);
      expect(hex(sha512(data))).toBe(createHash("sha512").update(data).digest("hex"));
    }
  });
});
