// Update the URL *or* replace it with a key from the xlwings.conf sheet
const url = "https://yoururl.com/endpoint";

// xlwings dev
// (c) 2022-present by Zoomer Analytics GmbH
// This is licensed under a commercial license and
// must be used with a valid license key.
// You will find the license under
// https://github.com/xlwings/xlwings/blob/master/LICENSE_PRO.txt

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Read config from sheet
  let base_url: string;
  let config = {};
  let configSheet = workbook.getWorksheet("xlwings.conf");
  if (configSheet) {
    const configValues = workbook
      .getWorksheet("xlwings.conf")
      .getRange("A1")
      .getExtendedRange(ExcelScript.KeyboardDirection.down)
      .getExtendedRange(ExcelScript.KeyboardDirection.right)
      .getValues();
    configValues.forEach((el) => (config[el[0].toString()] = el[1].toString()));
  }
  // Prepare config values
  if (url.includes("://")) {
    base_url = url;
  } else if (configSheet) {
    base_url = config[url];
  } else {
    throw "Missing URL!";
  }

  let exclude_sheets: [];
  if ("EXCLUDE_SHEETS" in config) {
    exclude_sheets = config["EXCLUDE_SHEETS"]
      .split(",")
      .map((item: string) => item.trim());
  } else {
    exclude_sheets = [];
  }

  // Payload
  let sheets = workbook.getWorksheets();
  let payload: {} = {};
  payload["version"] = "dev";
  payload["book"] = {
    name: workbook.getName(),
    active_sheet_index: workbook.getActiveWorksheet().getPosition(),
  };
  payload["sheets"] = [];
  let lastCellCol: number;
  let lastCellRow: number;
  sheets.forEach((sheet) => {
    let values: [][];
    let categories: [][];
    if (sheet.getUsedRange() !== undefined) {
      lastCellCol = sheet.getUsedRange().getLastCell().getColumnIndex();
      lastCellRow = sheet.getUsedRange().getLastCell().getRowIndex();
    } else {
      lastCellCol = 0;
      lastCellRow = 0;
    }
    if (exclude_sheets.includes(sheet.getName())) {
      values = [[]];
    } else {
      values = sheet
        .getRangeByIndexes(0, 0, lastCellRow + 1, lastCellCol + 1)
        .getValues();
      categories = sheet
        .getRangeByIndexes(0, 0, lastCellRow + 1, lastCellCol + 1)
        .getNumberFormatCategories();
      // Handle dates
      values.forEach((valueRow: [], rowIndex: number) => {
        const categoryRow = categories[rowIndex];
        valueRow.forEach((value, colIndex: number) => {
          const category = categoryRow[colIndex];
          if (category.toString() === "Date" && typeof value === "number") {
            values[rowIndex][colIndex] = new Date(
              Math.round((value - 25569) * 86400 * 1000)
            ).toISOString();
          }
        });
      });
    }
    // Update payload
    payload["sheets"].push({
      name: sheet.getName(),
      values: values,
    });
  });

  // console.log(payload);

  // Headers
  let headers = { "Content-Type": "application/json" };
  for (const property in config) {
    if (property.toLowerCase().startsWith("header_")) {
      headers[property.substring(7)] = config[property];
    }
  }

  // API call
  let response = await fetch(base_url, {
    method: "POST",
    headers: headers,
    body: JSON.stringify(payload),
  });

  // Parse JSON response
  let rawData: [
    {
      func: string;
      args: [];
      data: [][];
      sheet_position: number;
      start_row: number;
      start_column: number;
      row_count: number;
      column_count: number;
    }
  ];
  if (response.status !== 200) {
    throw `Error while contacting server: Error ${response.status}`;
  } else {
    rawData = await response.json();
  }

  // console.log(rawData);

  // Functions map
  let funcs = {
    setValues: setValues,
    clearContents: clearContents,
    addSheet: addSheet,
    setSheetName: setSheetName,
  };

  // Run Functions
  rawData.forEach((result) => {
    funcs[result.func](workbook, result);
  });
}

function setValues(workbook: ExcelScript.Workbook, result: {}) {
  // Handle DateTime
  let dt: Date;
  result.data.forEach((valueRow, rowIndex) => {
    valueRow.forEach((value, colIndex) => {
      if (typeof value === "string") {
        dt = new Date(Date.parse(value));
        let dtstr: string;
        dtstr = dt.toLocaleDateString();
        if (dtstr !== "Invalid Date") {
          if (
            value.length > 10 &&
            dt.getHours() +
              dt.getMinutes() +
              dt.getSeconds() +
              dt.getMilliseconds() !==
              0
          ) {
            dtstr += " " + dt.toLocaleTimeString();
          }
          result.data[rowIndex][colIndex] = dtstr;
        }
      }
    });
  });

  workbook
    .getWorksheets()
    [result.sheet_position].getRangeByIndexes(
      result.start_row,
      result.start_column,
      result.row_count,
      result.column_count
    )
    .setValues(result.data);
}

function clearContents(workbook: ExcelScript.Workbook, result: {}) {
  workbook
    .getWorksheets()
    [result.sheet_position].getRangeByIndexes(
      result.start_row,
      result.start_column,
      result.row_count,
      result.column_count
    )
    .clear(ExcelScript.ClearApplyTo.contents);
}

function addSheet(workbook: ExcelScript.Workbook, result: {}) {
  workbook.addWorksheet();
}

function setSheetName(workbook: ExcelScript.Workbook, result: {}) {
  workbook.getWorksheets()[result.sheet_position].setName(result.args[0]);
}
