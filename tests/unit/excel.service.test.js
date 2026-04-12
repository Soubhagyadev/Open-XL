import { getUsedRangeValues, writeValues } from "@/services/excel.service";

describe("excel.service", () => {
  beforeEach(() => {
    global.Excel = {
      run: jest.fn((cb) => cb({ sync: jest.fn(), workbook: mockWorkbook() })),
    };
  });

  function mockWorkbook() {
    const range = { load: jest.fn(), values: [["A", "B"]] };
    return {
      worksheets: {
        getActiveWorksheet: () => ({
          getUsedRange: () => range,
          getRange: () => range,
        }),
      },
    };
  }

  it("getUsedRangeValues returns values", async () => {
    const values = await getUsedRangeValues();
    expect(values).toEqual([["A", "B"]]);
  });

  it("writeValues calls Excel.run", async () => {
    await writeValues("A1", [["Hello"]]);
    expect(Excel.run).toHaveBeenCalled();
  });
});
