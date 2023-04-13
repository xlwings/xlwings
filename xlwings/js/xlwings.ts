async function main(workbook: ExcelScript.Workbook) {
  await runPython(workbook, "url", { auth: "DEVELOPMENT" });
}

/**
 * xlwings for Microsoft Office Scripts
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

globalThis.callbacks = {};
async function runPython(
  workbook: ExcelScript.Workbook,
  url = "",
  {
    auth = "",
    apiKey = "",
    include = "",
    exclude = "",
    headers = {},
  }: Options = {}
): Promise<void> {
  const version = "dev";
  const sheets = workbook.getWorksheets();
  // Config
  let configSheet = workbook.getWorksheet("xlwings.conf");
  let config = {};
  if (configSheet) {
    // @ts-ignore
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
  payload["version"] = version;
  payload["book"] = {
    name: workbook.getName(),
    active_sheet_index: workbook.getActiveWorksheet().getPosition(),
    selection: workbook.getSelectedRange().getAddress().split("!").pop(),
  };

  // Names (book scope only)
  let names: Names[] = [];
  workbook.getNames().forEach((namedItem, ix) => {
    // Currently filtering to named ranges
    // TODO: add sheet scoped named ranges via sheets as in officejs
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
  let values: (string | number | boolean)[][] = [[]];
  let categories: ExcelScript.NumberFormatCategory[][];
  sheets.forEach((sheet) => {
    let isSheetIncluded = !excludeArray.includes(sheet.getName());
    if (sheet.getUsedRange() !== undefined) {
      let lastCell = sheet.getUsedRange().getLastCell();
      lastCellCol = lastCell.getColumnIndex();
      lastCellRow = lastCell.getRowIndex();
    } else {
      lastCellCol = 0;
      lastCellRow = 0;
    }
    if (isSheetIncluded) {
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
    // Tables
    let tables: Tables[] = [];
    if (isSheetIncluded) {
      for (let table of sheet.getTables()) {
        tables.push({
          name: table.getName(),
          range_address: table.getRange().getAddress().split("!").pop(),
          header_row_range_address: table.getShowHeaders()
            ? table.getHeaderRowRange().getAddress().split("!").pop()
            : null,
          data_body_range_address: table
            .getRangeBetweenHeaderAndTotal()
            .getAddress()
            .split("!")
            .pop(),
          total_row_range_address: table.getShowTotals()
            ? table.getTotalRowRange().getAddress().split("!").pop()
            : null,
          show_headers: table.getShowHeaders(),
          show_totals: table.getShowTotals(),
          table_style: table.getPredefinedTableStyle(),
          show_autofilter: table.getShowFilterButton(),
        });
      }
    }

    // Pictures
    let pictures: Pictures[] = [];
    if (isSheetIncluded) {
      for (let shape of sheet.getShapes())
        if (shape.getType() === ExcelScript.ShapeType.image) {
          pictures.push({
            name: shape.getName(),
            width: shape.getWidth(),
            height: shape.getHeight(),
          });
        }
    }

    payload["sheets"].push({
      name: sheet.getName(),
      values: values,
      pictures: pictures,
      tables: tables,
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
    const forceSync = ["sheet", "table", "copy", "picture"];
    rawData["actions"].forEach((action) => {
      if (action.func === "addPicture") {
        // addPicture doesn't manage to pull both top and left from anchorCell otherwise
        addPicture(workbook, action);
      } else if (action.func === "updatePicture") {
        updatePicture(workbook, action);
      } else {
        globalThis.callbacks[action.func](workbook, action);
      }
      if (forceSync.some((el) => action.func.toLowerCase().includes(el))) {
        console.log(); // Force sync
      }
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
  address: string | undefined;
  book_scope: boolean;
}

interface Tables {
  name: string;
  range_address: string | undefined;
  header_row_range_address: string | undefined | null;
  data_body_range_address: string | undefined;
  total_row_range_address: string | undefined | null;
  show_headers: boolean;
  show_totals: boolean;
  table_style: string;
  show_autofilter: boolean;
}

interface Pictures {
  name: string;
  height: number;
  width: number;
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

function getShapeByType(
  workbook: ExcelScript.Workbook,
  sheetPosition: number,
  shapeIndex: number,
  shapeType: ExcelScript.ShapeType
) {
  const myshapes = workbook
    .getWorksheets()
    [sheetPosition].getShapes()
    .filter((shape: ExcelScript.Shape) => shape.getType() === shapeType);
  return myshapes[shapeIndex];
}

function registerCallback(callback: Function) {
  globalThis.callbacks[callback.name] = callback;
}

// Callbacks
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
registerCallback(setValues);

function clearContents(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).clear(ExcelScript.ClearApplyTo.contents);
}
registerCallback(clearContents);

function addSheet(workbook: ExcelScript.Workbook, action: Action) {
  let sheet: ExcelScript.Worksheet;
  if (action.args[1] !== null) {
    sheet = workbook.addWorksheet(action.args[1].toString());
  } else {
    sheet = workbook.addWorksheet();
  }
  sheet.setPosition(parseInt(action.args[0].toString()));
}
registerCallback(addSheet);

function setSheetName(workbook: ExcelScript.Workbook, action: Action) {
  workbook
    .getWorksheets()
    [action.sheet_position].setName(action.args[0].toString());
}
registerCallback(setSheetName);

function setAutofit(workbook: ExcelScript.Workbook, action: Action) {
  if (action.args[0] === "columns") {
    getRange(workbook, action).getFormat().autofitColumns();
  } else {
    getRange(workbook, action).getFormat().autofitRows();
  }
}
registerCallback(setAutofit);

function setRangeColor(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action)
    .getFormat()
    .getFill()
    .setColor(action.args[0].toString());
}
registerCallback(setRangeColor);

function activateSheet(workbook: ExcelScript.Workbook, action: Action) {
  workbook.getWorksheets()[parseInt(action.args[0].toString())].activate();
}
registerCallback(activateSheet);

function addHyperlink(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).setHyperlink({
    address: action.args[0].toString(),
    textToDisplay: action.args[1].toString(),
    screenTip: action.args[2].toString(),
  });
}
registerCallback(addHyperlink);

function setNumberFormat(workbook: ExcelScript.Workbook, action: Action) {
  getRange(workbook, action).setNumberFormat(action.args[0].toString());
}
registerCallback(setNumberFormat);

function setPictureName(workbook: ExcelScript.Workbook, action: Action) {
  const myshape = getShapeByType(
    workbook,
    action.sheet_position,
    Number(action.args[0]),
    ExcelScript.ShapeType.image
  );
  myshape.setName(action.args[1].toString());
}
registerCallback(setPictureName);

function setPictureHeight(workbook: ExcelScript.Workbook, action: Action) {
  const myshape = getShapeByType(
    workbook,
    action.sheet_position,
    Number(action.args[0]),
    ExcelScript.ShapeType.image
  );
  myshape.setHeight(Number(action.args[1]));
}
registerCallback(setPictureHeight);

function setPictureWidth(workbook: ExcelScript.Workbook, action: Action) {
  const myshape = getShapeByType(
    workbook,
    action.sheet_position,
    Number(action.args[0]),
    ExcelScript.ShapeType.image
  );
  myshape.setWidth(Number(action.args[1]));
}
registerCallback(setPictureWidth);

function deletePicture(workbook: ExcelScript.Workbook, action: Action) {
  const myshape = getShapeByType(
    workbook,
    action.sheet_position,
    Number(action.args[0]),
    ExcelScript.ShapeType.image
  );
  myshape.delete();
}
registerCallback(deletePicture);

function addPicture(workbook: ExcelScript.Workbook, action: Action) {
  const selection = workbook.getSelectedRange();
  const imageBase64 = action["args"][0].toString();
  const colIndex = Number(action["args"][1]);
  const rowIndex = Number(action["args"][2]);
  let left = Number(action["args"][3]);
  let top = Number(action["args"][4]);

  const sheet = workbook.getWorksheets()[action.sheet_position];
  let anchorCell = sheet.getRangeByIndexes(rowIndex, colIndex, 1, 1);
  left = Math.max(left, anchorCell.getLeft());
  top = Math.max(top, anchorCell.getTop());
  const image = sheet.addImage(imageBase64);
  image.setLeft(left);
  image.setTop(top);
  selection.select();
}
registerCallback(addPicture);

function updatePicture(workbook: ExcelScript.Workbook, action: Action) {
  const selection = workbook.getSelectedRange();
  const imageBase64 = action["args"][0].toString();
  const sheet = workbook.getWorksheets()[action.sheet_position];
  let image = getShapeByType(
    workbook,
    action.sheet_position,
    Number(action.args[1]),
    ExcelScript.ShapeType.image
  );
  let imgName = image.getName();
  let imgLeft = image.getLeft();
  let imgTop = image.getTop();
  let imgHeight = image.getHeight();
  let imgWidth = image.getWidth();
  image.delete();

  const newImage = sheet.addImage(imageBase64);
  newImage.setName(imgName);
  newImage.setLeft(imgLeft);
  newImage.setTop(imgTop);
  newImage.setHeight(imgHeight);
  newImage.setWidth(imgWidth);
  selection.select();
}
registerCallback(updatePicture);

function alert(workbook: ExcelScript.Workbook, action: Action) {
  // OfficeScripts doesn't have an any alert outside of DataValidation...
  let myPrompt = action.args[0];
  let myTitle = action.args[1]; // ignored
  let myButtons = action.args[2]; // ignored
  let myMode = action.args[3]; // ignored
  let myCallback = action.args[4]; // ignored
  throw myPrompt;
}
registerCallback(alert);

function setRangeName(workbook: ExcelScript.Workbook, action: Action) {
  throw "NotImplemented: setRangeName";
}
registerCallback(setRangeName);

function namesAdd(workbook: ExcelScript.Workbook, action: Action) {
  throw "NotImplemented: namesAdd";
}
registerCallback(namesAdd);

function nameDelete(workbook: ExcelScript.Workbook, action: Action) {
  throw "NotImplemented: deleteName";
}
registerCallback(nameDelete);

function runMacro(workbook: ExcelScript.Workbook, action: Action) {
  globalThis.callbacks[action.args[0].toString()](
    workbook,
    ...action.args.slice(1)
  );
}
registerCallback(runMacro);

function rangeDelete(workbook: ExcelScript.Workbook, action: Action) {
  let shift = action.args[0].toString();
  if (shift === "up") {
    getRange(workbook, action).delete(ExcelScript.DeleteShiftDirection.up);
  } else if (shift === "left") {
    getRange(workbook, action).delete(ExcelScript.DeleteShiftDirection.left);
  }
}
registerCallback(rangeDelete);

function rangeInsert(workbook: ExcelScript.Workbook, action: Action) {
  let shift = action.args[0].toString();
  if (shift === "down") {
    getRange(workbook, action).insert(ExcelScript.InsertShiftDirection.down);
  } else if (shift === "right") {
    getRange(workbook, action).insert(ExcelScript.InsertShiftDirection.right);
  }
}
registerCallback(rangeInsert);

function addTable(workbook: ExcelScript.Workbook, action: Action) {
  let mytable = workbook
    .getWorksheets()
    [action.sheet_position].addTable(
      action.args[0].toString(),
      Boolean(action.args[1])
    );
  if (action.args[2] !== null) {
    mytable.setPredefinedTableStyle(action.args[2].toString());
  }
  if (action.args[3] !== null) {
    mytable.setName(action.args[3].toString());
  }
}
registerCallback(addTable);

function setTableName(workbook: ExcelScript.Workbook, action: Action) {
  const mytable = workbook.getWorksheets()[action.sheet_position].getTables()[
    parseInt(action.args[0].toString())
  ];
  mytable.setName(action.args[1].toString());
}
registerCallback(setTableName);

function resizeTable(workbook: ExcelScript.Workbook, action: Action) {
  const mytable = workbook.getWorksheets()[action.sheet_position].getTables()[
    parseInt(action.args[0].toString())
  ];
  mytable.resize(action.args[1].toString());
}
registerCallback(resizeTable);

function showAutofilterTable(workbook: ExcelScript.Workbook, action: Action) {
  const mytable = workbook.getWorksheets()[action.sheet_position].getTables()[
    parseInt(action.args[0].toString())
  ];
  mytable.setShowFilterButton(Boolean(action.args[1]));
}
registerCallback(showAutofilterTable);

function showHeadersTable(workbook: ExcelScript.Workbook, action: Action) {
  const mytable = workbook.getWorksheets()[action.sheet_position].getTables()[
    parseInt(action.args[0].toString())
  ];
  mytable.setShowHeaders(Boolean(action.args[1]));
}
registerCallback(showHeadersTable);

function showTotalsTable(workbook: ExcelScript.Workbook, action: Action) {
  const mytable = workbook.getWorksheets()[action.sheet_position].getTables()[
    parseInt(action.args[0].toString())
  ];
  mytable.setShowTotals(Boolean(action.args[1]));
}
registerCallback(showTotalsTable);

function setTableStyle(workbook: ExcelScript.Workbook, action: Action) {
  const mytable = workbook.getWorksheets()[action.sheet_position].getTables()[
    parseInt(action.args[0].toString())
  ];
  mytable.setPredefinedTableStyle(action.args[1].toString());
}
registerCallback(setTableStyle);

function copyRange(workbook: ExcelScript.Workbook, action: Action) {
  const destination = workbook
    .getWorksheets()
    [parseInt(action.args[0].toString())].getRange(action.args[1].toString());
  destination.copyFrom(getRange(workbook, action));
}
registerCallback(copyRange);
