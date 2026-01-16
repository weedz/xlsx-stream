export function getCellId(rowIndex: number, x: number) {
  let name = "";
  do {
    name = String.fromCodePoint((x % 26) + 65) + name;
    x = Math.floor(x / 26) - 1;
  } while (x >= 0);
  return `${name}${rowIndex + 1}`;
}

export function sanitize(text: string) {
  const escaped = text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

  let str = "";
  for (let i = 0, len = escaped.length; i < len; ++i) {
    const letter = escaped[i]!;
    if (
      letter === "\x09" ||
      letter === "\x0A" ||
      letter === "\x0D" ||
      (letter >= "\x20" && letter <= "\uD7FF") ||
      (letter >= "\uE000" && letter <= "\uFFFD")
    ) {
      str += letter;
    }
  }
  return str;
}

export function getNumberFormat(n: number) {
  const number = Math.abs(n);
  if (number.toString().length >= 11) {
    return 5;
  }
  if (number % 1 === 0) {
    return number >= 1000 ? 3 : 1;
  }
  return number >= 1000 ? 4 : 2;
}

export function getDateFormat() {
  return 6;
}
