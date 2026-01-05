import { Transform, type TransformCallback } from "node:stream";
import XLSXRowTransform from "./XLSXRowTransform.js";
import * as templates from "./templates/index.js";
import Archiver from "archiver";

interface Options {
  shouldFormat?: boolean;
}
/** Class representing a XLSX Transform Stream */
export class XLSXTransformStream extends Transform {
  options: Options;
  rowTransform: XLSXRowTransform;
  zip: Archiver.Archiver;
  /**
   * Create a new Stream
   */
  constructor(options: Options = {}) {
    super({ objectMode: true });
    this.options = options;
    this.rowTransform = new XLSXRowTransform(this.options.shouldFormat || false);

    this.zip = Archiver("zip");

    this.initializeArchiver();
    this.zip.append(this.rowTransform, {
      name: "xl/worksheets/sheet1.xml",
    });
  }

  initializeArchiver() {
    this.zip.on("data", (data) => {
      this.push(data);
    });

    this.zip.append(templates.ContentTypes, {
      name: "[Content_Types].xml",
    });

    this.zip.append(templates.Rels, {
      name: "_rels/.rels",
    });

    this.zip.append(templates.Workbook, {
      name: "xl/workbook.xml",
    });

    this.zip.append(templates.Styles, {
      name: "xl/styles.xml",
    });

    this.zip.append(templates.WorkbookRels, {
      name: "xl/_rels/workbook.xml.rels",
    });

    this.zip.on("warning", (err) => {
      console.warn(err);
    });

    this.zip.on("error", (err) => {
      console.error(err);
    });
  }
  _transform(row: any, _encoding: BufferEncoding, callback: TransformCallback) {
    if (this.rowTransform.write(row)) {
      process.nextTick(callback);
    } else {
      this.rowTransform.once("drain", callback);
    }
  }
  _flush(callback: TransformCallback) {
    this.rowTransform.end();
    this.zip.finalize().then(() => callback());
  }
}
