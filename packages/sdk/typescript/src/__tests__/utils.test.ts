import { columnNumberFromName } from "../utils";

describe("utils", () => {
  test("columnNumberFromName", () => {
    expect(columnNumberFromName("A")).toBe(1);
    expect(columnNumberFromName("Z")).toBe(26);
    expect(columnNumberFromName("AA")).toBe(27);
    expect(columnNumberFromName("AB")).toBe(28);
    expect(columnNumberFromName("ZZ")).toBe(702);
    expect(columnNumberFromName("AAA")).toBe(703);
    expect(columnNumberFromName("AZZ")).toBe(1378);
    expect(columnNumberFromName("BAA")).toBe(1379);
    expect(columnNumberFromName("AZ")).toBe(52);
    expect(columnNumberFromName("SZ")).toBe(520);
    expect(columnNumberFromName("CUZ")).toBe(2600);
    expect(columnNumberFromName("NTP")).toBe(10_000);
    expect(columnNumberFromName("XFD")).toBe(16_384);
  });
});
