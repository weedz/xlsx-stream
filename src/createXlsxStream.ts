import { Readable } from "node:stream";
import { Buffer } from "node:buffer";
import { createZipArchive, ZipEntry } from "node:zlib";

import * as templates from "./templates/index.js";

interface Options {
  shouldFormat?: boolean;
}

export function createXLSXStream(options: Options = {}) {
  const shouldFormat = options.shouldFormat ?? false;
  return function (data: AsyncIterable<unknown[]>) {
    const entries = Readable.from([
      ZipEntry.create("[Content_Types].xml", Buffer.from(templates.ContentTypes)),
      ZipEntry.create("_rels/.rels", Buffer.from(templates.Rels)),
      ZipEntry.create("xl/workbook.xml", Buffer.from(templates.Workbook)),
      ZipEntry.create("xl/styles.xml", Buffer.from(templates.Styles)),
      ZipEntry.create("xl/_rels/workbook.xml.rels", Buffer.from(templates.WorkbookRels)),

      ZipEntry.createStream("xl/worksheets/sheet1.xml", transformXlsxRow(data, shouldFormat)),
    ]);
    return createZipArchive(entries, "created by node:zlib");
  };
}

export async function* transformXlsxRow(data: AsyncIterable<unknown[]>, shouldFormat: boolean) {
  yield Buffer.from(templates.SheetHeader);

  let rowCount = 0;
  for await (const row of data) {
    yield Buffer.from(templates.Row(rowCount, row, shouldFormat));
    rowCount++;
  }

  yield Buffer.from(templates.SheetFooter);
}
