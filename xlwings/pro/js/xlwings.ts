// Config (actual values or keys in xlwings.conf sheet)
const url = "URL"; // required
const apiKey = "API_KEY"; // required
const excludeSheets = "EXCLUDE_SHEETS"; // optional

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
  let url_: string = getConfig(url, config);
  let headerApiKey: string = getConfig(apiKey, config);

  let excludeSheetsString: string = getConfig(excludeSheets, config);
  let excludeSheetsArray: string[] = [];
  excludeSheetsArray = excludeSheetsString
    .split(",")
    .map((item: string) => item.trim());

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
  let values: (string | number | boolean)[][];
  let categories: ExcelScript.NumberFormatCategory[][];
  sheets.forEach((sheet) => {
    if (sheet.getUsedRange() !== undefined) {
      let lastCell = sheet.getUsedRange().getLastCell();
      lastCellCol = lastCell.getColumnIndex();
      lastCellRow = lastCell.getRowIndex();
    } else {
      lastCellCol = 0;
      lastCellRow = 0;
    }
    if (excludeSheetsArray.includes(sheet.getName())) {
      values = [[]];
    } else {
      let range = sheet.getRangeByIndexes(
        0,
        0,
        lastCellRow + 1,
        lastCellCol + 1
      );
      values = range.getValues();
      categories = range.getNumberFormatCategories();
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
  let response = await fetch(url_, {
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

  // Run Functions
  const forceSync = ["sheet"];
  rawData["actions"].forEach((action) => {
    if (forceSync.some(el => action.func.toLowerCase().includes(el))) {
      console.log(); // Force sync to prevent writing to wrong sheet
    }
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

function getConfig(keyOrValue: string, config: {}) {
  if (keyOrValue in config) {
    return config[keyOrValue];
  } else {
    return keyOrValue;
  }
}

// Functions map
let funcs = {
  setValues: setValues,
  clearContents: clearContents,
  addSheet: addSheet,
  setSheetName: setSheetName,
  setAutofit: setAutofit,
  setRangeColor: setRangeColor,
  activateSheet: activateSheet,
};

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
  let sheet = workbook.addWorksheet();
  sheet.setPosition(action.args[0]);
}

function setSheetName(workbook: ExcelScript.Workbook, action: Action) {
  workbook.getWorksheets()[action.sheet_position].setName(action.args[0]);
}

function setAutofit(workbook: ExcelScript.Workbook, action: Action) {
  if (action.args[0] === 'columns') {
    getRange(workbook, action).getFormat().autofitColumns();
  } else {
    getRange(workbook, action).getFormat().autofitRows();
  }
}

function setRangeColor(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).getFormat().getFill().setColor(action.args[0])
}

function activateSheet(workbook: ExcelScript.Workbook, action: Action) {
  workbook.getWorksheets()[action.args[0]].activate()
}
