// Config (actual values or keys in xlwings.conf sheet)
const url = "URL"; // required
const apiKey = "DEVELOPMENT"; // required
const excludeSheets = "EXCLUDE_SHEETS"; // optional

/**
 * xlwings dev
 * (c) 2022-present by Zoomer Analytics GmbH
 * This file is licensed under a commercial license and
 * must be used in connection with a valid license key.
 * You will find the license under
 * https://github.com/xlwings/xlwings/blob/master/LICENSE_PRO.txt
 */

function main() {
  workbook = SpreadsheetApp.getActive();
  let configSheet = workbook.getSheetByName("xlwings.conf");
  let config = {};
  let configValues = {};
  if (configSheet) {
    configValues = workbook
      .getSheetByName("xlwings.conf")
      .getRange("A1")
      .getDataRegion()
      .getValues();
    configValues.forEach((el) => (config[el[0].toString()] = el[1].toString()));
  }

  // Prepare config values
  let url_ = getConfig(url, config);
  let headerApiKey = getConfig(apiKey, config);

  let excludeSheetsString = getConfig(excludeSheets, config);
  let excludeSheetsArray = [];
  excludeSheetsArray = excludeSheetsString
    .split(",")
    .map((item) => item.trim());

  // Request payload
  let sheets = workbook.getSheets();
  let payload = {};
  payload["version"] = "dev";
  payload["book"] = {
    name: workbook.getName(),
    active_sheet_index: workbook.getActiveSheet().getIndex() - 1,
  };
  payload["sheets"] = [];
  let lastCellCol;
  let lastCellRow;
  let values;
  sheets.forEach((sheet) => {
    lastCellCol = sheet.getLastColumn();
    lastCellRow = sheet.getLastRow();
    if (excludeSheetsArray.includes(sheet.getName())) {
      values = [[]];
    } else {
      let range = sheet.getRange(1, 1, lastCellRow > 0 ? lastCellRow : 1, lastCellCol > 0 ? lastCellCol : 1);
      values = range.getValues();
    }
    payload["sheets"].push({
      name: sheet.getName(),
      values: values,
    });
  });

  // console.log(payload);

  // Headers
  let headers = {
    Authorization: headerApiKey,
  };
  for (const property in config) {
    if (property.toLowerCase().startsWith("header_")) {
      headers[property.substring(7)] = config[property];
    }
  }

  // API call
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: headers,
  };

  // Parse JSON response
  // TODO: handle non-200 status
  let response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var rawData = JSON.parse(json);

  // console.log(rawData);

  // Run Functions
  if (rawData !== null) {
    const forceSync = ["sheet"];
    rawData["actions"].forEach((action) => {
      if (forceSync.some((el) => action.func.toLowerCase().includes(el))) {
        SpreadsheetApp.flush(); // Force sync to prevent writing to wrong sheet
      }
      funcs[action.func](workbook, action);
    });
  }
}

// Helpers
function getRange(workbook, action) {
  return workbook
    .getSheets()
    [action.sheet_position].getRange(
      action.start_row + 1,
      action.start_column + 1,
      action.row_count,
      action.column_count
    );
}

function getConfig(keyOrValue, config) {
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
function setValues(workbook, action) {
  // Handle DateTime (TODO: backend should deliver indices with datetime obj)
  let dt;
  let dtString;
  action.values.forEach((valueRow, rowIndex) => {
    valueRow.forEach((value, colIndex) => {
      if (typeof value === "string") {
        dt = new Date(Date.parse(value));
        dtString = dt.toLocaleDateString();
        if (dtString !== "Invalid Date") {
          if (
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

function clearContents(workbook, action) {
  getRange(workbook, action).clearContent();
}

function addSheet(workbook, action) {
  let sheet = workbook.insertSheet(action.args[0]);
}

function setSheetName(workbook, action) {
  workbook
    .getSheets()
    [action.sheet_position].setName(action.args[0].toString());
}

function setAutofit(workbook, action) {
  if (action.args[0] === "columns") {
    workbook
      .getSheets()
      [action.sheet_position].autoResizeColumns(
        action.start_column + 1,
        action.start_column + action.column_count
      );
  } else {
    workbook
      .getSheets()
      [action.sheet_position].autoResizeRows(
        action.start_row + 1,
        action.start_row + action.row_count
      );
  }
}

function setRangeColor(workbook, action) {
  getRange(workbook, action).setBackground(action.args[0]);
}

function activateSheet(workbook, action) {
  workbook.getSheets()[parseInt(action.args[0])].activate();
}
