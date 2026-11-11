import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { getCellId, sanitize, getDateFormat, getNumberFormat } from "../utils.js";

describe("The getCellId function", () => {
  it("returns correct ID for first cell and row", () => {
    const expectedResult = "A1";
    const result = getCellId(0, 0);
    assert.equal(result, expectedResult);
  });
  it("returns correct ID for 25th cell and 30th row", () => {
    const expectedResult = "Y30";
    const result = getCellId(29, 24);
    assert.equal(result, expectedResult);
  });
  it("returns correct ID for 26th cell and 30th row", () => {
    const expectedResult = "Z30";
    const result = getCellId(29, 25);
    assert.equal(result, expectedResult);
  });
  it("returns correct ID for 27th cell and 30th row", () => {
    const expectedResult = "AA30";
    const result = getCellId(29, 26);
    assert.equal(result, expectedResult);
  });
  it("returns correct ID for 28th cell and 30th row", () => {
    const expectedResult = "AB30";
    const result = getCellId(29, 27);
    assert.equal(result, expectedResult);
  });
});

describe("The sanitize function", () => {
  it("escapes XML entities &, >, <", () => {
    const expectedResult = "&amp;&lt;&gt;";
    const result = sanitize("&<>");
    assert.equal(result, expectedResult);
  });

  it("removes invalid XML characters", () => {
    const expectedResult = "";
    const result = sanitize("\u001A");
    assert.equal(result, expectedResult);
  });
});

describe("Number formats function", () => {
  it("returns 5 for integer length 11+", () => {
    const expected = 5;
    assert.equal(getNumberFormat(10000000000), expected);
    assert.equal(getNumberFormat(1000000000000), expected);
    assert.equal(getNumberFormat(-10000000000), expected);
    assert.equal(getNumberFormat(-1000000000000), expected);
  });

  it("returns 1 for integer from 0 to 999", () => {
    const expected = 1;
    assert.equal(getNumberFormat(0), expected);
    assert.equal(getNumberFormat(-10), expected);
    assert.equal(getNumberFormat(999), expected);
    assert.equal(getNumberFormat(-1), expected);
  });

  it("returns 2 for float from 0 to 999", () => {
    const expected = 2;
    assert.equal(getNumberFormat(1.22), expected);
    assert.equal(getNumberFormat(-10.1231), expected);
    assert.equal(getNumberFormat(999.1), expected);
    assert.equal(getNumberFormat(-1.2222222), expected);
  });

  it("returns 3 for integer from 1000", () => {
    const expected = 3;
    assert.equal(getNumberFormat(1000), expected);
    assert.equal(getNumberFormat(1000000000), expected);
    assert.equal(getNumberFormat(-2500), expected);
    assert.equal(getNumberFormat(-38000), expected);
  });

  it("returns 4 for floats from 1000", () => {
    const expected = 4;
    assert.equal(getNumberFormat(1000.12), expected);
    assert.equal(getNumberFormat(10000000.1), expected);
    assert.equal(getNumberFormat(-2500.4232), expected);
    assert.equal(getNumberFormat(-3800.10012), expected);
  });
});

describe("Date formats function", () => {
  it("Should return 6 for all dates", () => {
    const expected = 6;
    assert.equal(getDateFormat(), expected);
  });
});
