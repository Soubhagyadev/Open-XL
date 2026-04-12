import { runExcel } from "@/utils/office";

/**
 * Read all values from the active worksheet's used range.
 * @returns {Promise<any[][]>}
 */
export async function getUsedRangeValues() {
  let values;
  await runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();
    values = range.values;
  });
  return values;
}

/**
 * Write a 2D array of values to a given address on the active sheet.
 * @param {string} address  e.g. "A1"
 * @param {any[][]} values
 */
export async function writeValues(address, values) {
  await runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.values = values;
    await context.sync();
  });
}
