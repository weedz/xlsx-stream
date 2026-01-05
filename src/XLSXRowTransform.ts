import { Transform, type TransformCallback } from "node:stream";
import { Row, SheetHeader, SheetFooter } from "./templates/index.js";

/** Class representing a XLSX Row transformation from array to Row. Also adds the necessary XLSX header and footer. */
export default class XLSXRowTransform extends Transform {
  shouldFormat: boolean;
  rowCount: number;
  constructor(shouldFormat: boolean) {
    super({ objectMode: true });
    this.rowCount = 0;
    this.shouldFormat = shouldFormat;
    this.push(SheetHeader);
  }
  /**
   * Transform array to row string
   */
  _transform(row: any, _encoding: BufferEncoding, callback: TransformCallback) {
    if (!Array.isArray(row)) return callback();

    const xlsxRow = Row(this.rowCount, row, this.shouldFormat);
    this.rowCount++;
    callback(null, xlsxRow);
  }

  _flush(callback: TransformCallback) {
    this.push(SheetFooter);
    callback();
  }
}
