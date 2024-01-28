import { readFile, WorkBook } from "../deps.ts";

export function tryReadXLSX(filename: string): WorkBook {
  try {
    return readFile(filename);
  } catch {
    console.error(
      `指定したファイル[${filename}]が正しく読み込めませんでした。`,
    );
    Deno.exit(1);
  }
}
