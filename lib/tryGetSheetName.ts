import { WorkBook } from "../deps.ts";

export function tryGetSheetName(workbook: WorkBook, sheet?: string): string {
  const sheetNameList = workbook.SheetNames;
  if (sheet !== undefined && sheetNameList.includes(sheet)) {
    return sheet;
  }
  if (sheetNameList.length === 0) {
    console.error(`指定したExcelファイルにシートがありません。`);
    Deno.exit(1);
  }
  return sheetNameList[0];
}
