// Config (actual values or keys in optional xlwings.conf sheet)
const url = "URL";
const apiKey = "API_KEY";

/**
 * xlwings dev
 * (c) 2022-present by Zoomer Analytics GmbH
 * This file is licensed under a commercial license and
 * must be used in connection with a valid license key.
 * You will find the license under
 * https://github.com/xlwings/xlwings/blob/master/LICENSE_PRO.txt
 */

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Read config from optional xlwings.conf sheet
  let configSheet = workbook.getWorksheet("xlwings.conf");
  let config = {};
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
  let base_url: string;
  if (url in config) {
    base_url = config[url];
  } else {
    base_url = url;
  }

  let headerApiKey: string;
  if (apiKey in config) {
    headerApiKey = config[apiKey];
  } else {
    headerApiKey = apiKey;
  }

  let exclude_sheets: string[] = [];
  if ("EXCLUDE_SHEETS" in config) {
    exclude_sheets = config["EXCLUDE_SHEETS"]
      .split(",")
      .map((item: string) => item.trim());
  }

  // Request payload
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
    let values: (string | number | boolean)[][];
    let categories: ExcelScript.NumberFormatCategory[][];
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
      values.forEach(
        (valueRow: (string | number | boolean)[], rowIndex: number) => {
          const categoryRow = categories[rowIndex];
          valueRow.forEach((value, colIndex: number) => {
            const category = categoryRow[colIndex];
            if (category.toString() === "Date" && typeof value === "number") {
              values[rowIndex][colIndex] = new Date(
                Math.round((value - 25569) * 86400 * 1000)
              ).toISOString();
            }
          });
        }
      );
    }
    payload["sheets"].push({
      name: sheet.getName(),
      values: values,
    });
  });

  // console.log(payload);

  // Headers
  let headers = {
    "Content-Type": "application/json",
    Authorization: headerApiKey,
  };
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
  let rawData: { actions: Action[] };
  if (response.status !== 200) {
    throw `Server responded with error ${response.status}`;
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
  rawData["actions"].forEach((action) => {
    funcs[action.func](workbook, action);
  });
}

// Helpers
interface Action {
  func: string;
  args: (string | number | boolean)[];
  values: (string | number | boolean)[][];
  sheet_position: number;
  start_row: number;
  start_column: number;
  row_count: number;
  column_count: number;
}

function getRange(workbook: ExcelScript.Workbook, action: Action) {
  return workbook
    .getWorksheets()
  [action.sheet_position].getRangeByIndexes(
    action.start_row,
    action.start_column,
    action.row_count,
    action.column_count
  );
}

// Functions
function setValues(workbook: ExcelScript.Workbook, action: Action) {
  // Handle DateTime
  let dt: Date;
  let dtString: string;
  action.values.forEach((valueRow, rowIndex) => {
    valueRow.forEach((value: string | number | boolean, colIndex) => {
      if (typeof value === "string") {
        dt = new Date(Date.parse(value));
        dtString = dt.toLocaleDateString();
        if (dtString !== "Invalid Date") {
          if (
            value.length > 10 &&
            dt.getHours() +
            dt.getMinutes() +
            dt.getSeconds() +
            dt.getMilliseconds() !==
            0
          ) {
            dtString += " " + dt.toLocaleTimeString();
          }
          action.values[rowIndex][colIndex] = dtString;
        }
      }
    });
  });
  getRange(workbook, action).setValues(action.values);
}

function clearContents(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).clear(ExcelScript.ClearApplyTo.contents);
}

function addSheet(workbook: ExcelScript.Workbook, action: Action) {
  workbook.addWorksheet();
}

function setSheetName(workbook: ExcelScript.Workbook, action: Action) {
  workbook.getWorksheets()[action.sheet_position].setName(action.args[0]);
}
