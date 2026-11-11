import { describe, it } from "node:test";
import assert from "node:assert/strict";

import { Readable } from "stream";
import { SheetHeader, SheetFooter } from "../templates/index.js";
import { transformXlsxRow } from "../createXlsxStream.js";

describe("The XLSXRowTransform", () => {
  it("Correctly transforms an array of data into a XLSX row format with header and footer", async () => {
    const expectedResult = `${SheetHeader}
        <row r="1" spans="1:2" x14ac:dyDescent="0.2">
            <c r="A1" t="inlineStr"><is><t>test</t></is></c>
			<c r="B1" t="n"><v>123</v></c>
        </row>${SheetFooter}`;

    const inputStream = Readable.from([["test", 123]]);
    const result = Buffer.concat(await Array.fromAsync(transformXlsxRow(inputStream, false)));

    assert.equal(result.toString("utf8"), expectedResult);
  });
});
