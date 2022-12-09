async function main(workbook: ExcelScript.Workbook) {
  await runPython(workbook, "url", { auth: "..." });
}

/**
 * xlwings dev (for Microsoft Office Scripts)
 * Copyright (C) 2014 - present, Zoomer Analytics GmbH.
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without modification,
 * are permitted provided that the following conditions are met:
 *
 * * Redistributions of source code must retain the above copyright notice, this
 *   list of conditions and the following disclaimer.
 *
 * * Redistributions in binary form must reproduce the above copyright notice, this
 *   list of conditions and the following disclaimer in the documentation and/or
 *   other materials provided with the distribution.
 *
 * * Neither the name of the copyright holder nor the names of its
 *   contributors may be used to endorse or promote products derived from
 *   this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
 * ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR
 * ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 * ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

async function runPython(
  workbook: ExcelScript.Workbook,
  url = "",
  { auth = "", apiKey = "", include = "", exclude = "", headers = {} }: Options = {}
): Promise<void> {
  const sheets = workbook.getWorksheets();
  // Config
  let configSheet = workbook.getWorksheet("xlwings.conf");
  let config = {};
  if (configSheet) {
    const configValues = workbook
      .getWorksheet("xlwings.conf")
      .getRange("A1")
      .getSurroundingRegion()
      .getValues();
    configValues.forEach((el) => (config[el[0].toString()] = el[1].toString()));
  }

  if (apiKey === "") {
    apiKey = config["API_KEY"] || "";
  }
  if (auth === "") {
    auth = config["AUTH"] || "";
  }

  if (include === "") {
    include = config["INCLUDE"] || "";
  }
  let includeArray: string[] = [];
  if (include !== "") {
    includeArray = include.split(",").map((item) => item.trim());
  }

  if (exclude === "") {
    exclude = config["EXCLUDE"] || "";
  }
  let excludeArray: string[] = [];
  if (exclude !== "") {
    excludeArray = exclude.split(",").map((item) => item.trim());
  }
  if (includeArray.length > 0 && excludeArray.length > 0) {
    throw "Either use 'include' or 'exclude', but not both!";
  }
  if (includeArray.length > 0) {
    sheets.forEach((sheet) => {
      if (!includeArray.includes(sheet.getName())) {
        excludeArray.push(sheet.getName());
      }
    });
  }

  if (Object.keys(headers).length === 0) {
    for (const property in config) {
      if (property.toLowerCase().startsWith("header_")) {
        headers[property.substring(7)] = config[property];
      }
    }
  }
  // Deprecated: replaced by "auth"
  if (!("Authorization" in headers) && apiKey.length > 0) {
    headers["Authorization"] = apiKey;
  }
  if (!("Authorization" in headers) && auth.length > 0) {
    headers["Authorization"] = auth;
  }

  // Standard headers
  headers["Content-Type"] = "application/json";

  // Request payload
  let payload: {} = {};
  payload["client"] = "Microsoft Office Scripts";
  payload["version"] = "dev";
  payload["book"] = {
    name: workbook.getName(),
    active_sheet_index: workbook.getActiveWorksheet().getPosition(),
    selection: workbook.getSelectedRange().getAddress().split("!").pop(),
  };

  // Names
  let names: Names[] = [];
  workbook.getNames().forEach((namedItem, ix) => {
    // Currently filtering to named ranges
    // Sheet scope names don't seem to come through despite the existence of getScope()
    let itemType: ExcelScript.NamedItemType = namedItem.getType();
    if (itemType === ExcelScript.NamedItemType.range) {
      names[ix] = {
        name: namedItem.getName(),
        sheet_index: namedItem.getRange().getWorksheet().getPosition(),
        address: namedItem.getRange().getAddress().split("!").pop(),
        book_scope:
          namedItem.getScope() === ExcelScript.NamedItemScope.workbook,
      };
    }
  });
  payload["names"] = names;

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
    if (excludeArray.includes(sheet.getName())) {
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
            if (
              (category.toString() === "Date" ||
                category.toString() === "Time") &&
              typeof value === "number"
            ) {
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
      pictures: [], // TODO: NotImplemented
    });
  });

  // console.log(payload);

  // API call
  let response = await fetch(url, {
    method: "POST",
    headers: headers,
    body: JSON.stringify(payload),
  });

  // Parse JSON response
  let rawData: { actions: Action[] };
  if (response.status !== 200) {
    throw await response.text();
  } else {
    rawData = await response.json();
  }

  // console.log(rawData);

  // Run Functions
  if (rawData !== null) {
    const forceSync = ["sheet"];
    rawData["actions"].forEach((action) => {
      if (forceSync.some((el) => action.func.toLowerCase().includes(el))) {
        console.log(); // Force sync to prevent writing to wrong sheet
      }
      funcs[action.func](workbook, action);
    });
  }
}

// Helpers
interface Options {
  auth?: string;
  apiKey?: string;
  include?: string;
  exclude?: string;
  headers?: {};
}

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

interface Names {
  name: string;
  sheet_index: number;
  address: string;
  book_scope: boolean;
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

// Functions map
let funcs = {
  setValues: setValues,
  clearContents: clearContents,
  addSheet: addSheet,
  setSheetName: setSheetName,
  setAutofit: setAutofit,
  setRangeColor: setRangeColor,
  activateSheet: activateSheet,
  addHyperlink: addHyperlink,
  setNumberFormat: setNumberFormat,
  setPictureName: setPictureName,
  setPictureWidth: setPictureWidth,
  setPictureHeight: setPictureHeight,
  deletePicture: deletePicture,
  addPicture: addPicture,
  updatePicture: updatePicture,
  alert: alert,
};

// Functions
function setValues(workbook: ExcelScript.Workbook, action: Action) {
  // Handle DateTime (TODO: backend should deliver indices with datetime obj)
  let dt: Date;
  let dtString: string;
  action.values.forEach((valueRow, rowIndex) => {
    valueRow.forEach((value: string | number | boolean, colIndex) => {
      if (
        typeof value === "string" &&
        value.length > 18 &&
        value.includes("T")
      ) {
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

function clearContents(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).clear(ExcelScript.ClearApplyTo.contents);
}

function addSheet(workbook: ExcelScript.Workbook, action: Action) {
  let sheet = workbook.addWorksheet();
  sheet.setPosition(parseInt(action.args[0].toString()));
}

function setSheetName(workbook: ExcelScript.Workbook, action: Action) {
  workbook
    .getWorksheets()
    [action.sheet_position].setName(action.args[0].toString());
}

function setAutofit(workbook: ExcelScript.Workbook, action: Action) {
  if (action.args[0] === "columns") {
    getRange(workbook, action).getFormat().autofitColumns();
  } else {
    getRange(workbook, action).getFormat().autofitRows();
  }
}

function setRangeColor(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action)
    .getFormat()
    .getFill()
    .setColor(action.args[0].toString());
}

function activateSheet(workbook: ExcelScript.Workbook, action: Action) {
  workbook.getWorksheets()[parseInt(action.args[0].toString())].activate();
}

function addHyperlink(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).setHyperlink({
    address: action.args[0].toString(),
    textToDisplay: action.args[1].toString(),
    screenTip: action.args[2].toString(),
  });
}

function setNumberFormat(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).setNumberFormat(action.args[0].toString());
}

function setPictureName(workbook: ExcelScript.Workbook, action: Action) {
  throw "Not Implemented: setPictureName";
}

function setPictureHeight(workbook: ExcelScript.Workbook, action: Action) {
  throw "Not Implemented: setPictureHeight";
}

function setPictureWidth(workbook: ExcelScript.Workbook, action: Action) {
  throw "Not Implemented: setPictureWidth";
}

function deletePicture(workbook: ExcelScript.Workbook, action: Action) {
  throw "Not Implemented: deletePicture";
}

function addPicture(workbook: ExcelScript.Workbook, action: Action) {
  throw "Not Implemented: addPicture";
}

function updatePicture(workbook: ExcelScript.Workbook, action: Action) {
  throw "Not Implemented: updatePicture";
}

function alert(workbook: ExcelScript.Workbook, action: Action) {
  throw "Not Implemented: alert";
}
