import { Command, utils } from "./deps.ts";
import { tryReadXLSX } from "./lib/tryReadXLSX.ts";
import { tryGetSheetName } from "./lib/tryGetSheetName.ts";

await new Command()
  .name("excel-formula-to-csv")
  .description("Excelの式をCSVに直します")
  .version("v1.0.0")
  .option("-r, --row <row:number>", "縦に何行分にするか", {
    default: 1000,
  })
  .option("-c, --column <column:number>", "横に何列分にするか", {
    default: 1000,
  })
  .option("-s, --sheet <sheet:string>", "シート名", {
    default: undefined,
  })
  .arguments("[filename]")
  .action(async ({ row, column, sheet }, filename) => {
    if (filename === undefined) {
      console.error("ファイル名を指定してください");
      Deno.exit(1);
    }
    const workbook = tryReadXLSX(filename);
    const sheetName = tryGetSheetName(workbook, sheet);

    const cells: string[] = [`"行番号"`];
    for (let c, i = 0; i < column; i++) {
      c = utils.encode_col(i);
      cells.push(`"${c}"`);
    }
    console.log(cells.join(","));

    for (let r = 0; r < row; r++) {
      let cells: string[] = [`${utils.encode_row(r)}`];
      for (let c = 0; c < column; c++) {
        const cell = utils.encode_cell({ c, r });
        const cellValue:
          | { v?: string; t?: string; f?: string; w?: string }
          | undefined = workbook.Sheets[sheetName][cell];
        if (cellValue === undefined) {
          cells.push("");
          continue;
        }
        if (cellValue.f !== undefined) {
          cells.push(`${cellValue.f}`);
          continue;
        }
        cells.push(cellValue.w ?? "");
      }
      console.log(cells.map((cell) => `"${cell}"`).join(","));
    }
  })
  .parse();
