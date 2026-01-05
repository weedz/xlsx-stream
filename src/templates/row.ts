import cell from "./cell.js";
import { getCellId } from "../utils.js";

function row(index: number, values: unknown[], shouldFormat: boolean) {
  return `
        <row r="${index + 1}" spans="1:${values.length}" x14ac:dyDescent="0.2">
            ${values.map((cellValue, cellIndex) => cell(cellValue, getCellId(index, cellIndex), shouldFormat)).join("\n\t\t\t")}
        </row>`;
}

export default row;
