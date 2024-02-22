function hello() {
  runPython("url", { auth: "DEVELOPMENT" });
}

/**
 * xlwings for Google Apps Script
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

/**
 * @OnlyCurrentDoc
 */

function runPython(
  url,
  { auth = "", apiKey = "", include = "", exclude = "", headers = {} } = {}
) {
  const version = "dev";
  const workbook = SpreadsheetApp.getActive();
  const sheets = workbook.getSheets();

  // Only used to request permission for proper OAuth Scope
  Session.getActiveUser().getEmail();

  // Config
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

  if (apiKey === "") {
    apiKey = config["API_KEY"] || "";
  }

  if (auth === "") {
    auth = config["AUTH"] || "";
  }

  if (include === "") {
    include = config["INCLUDE"] || "";
  }
  let includeArray = [];
  if (include !== "") {
    includeArray = include.split(",").map((item) => item.trim());
  }

  if (exclude === "") {
    exclude = config["EXCLUDE"] || "";
  }
  let excludeArray = [];
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

  // Request payload
  let payload = {};
  payload["client"] = "Google Apps Script";
  payload["version"] = version;
  payload["book"] = {
    name: workbook.getName(),
    active_sheet_index: workbook.getActiveSheet().getIndex() - 1,
    selection: workbook.getActiveRange().getA1Notation(),
  };

  // Names
  let names = [];
  workbook.getNamedRanges().forEach((namedRange, ix) => {
    let name = namedRange.getName().includes(" ")
      ? namedRange.getName()
      : namedRange.getName().replace("'", "").replace("'", "");
    names[ix] = {
      name: name,
      sheet_index: namedRange.getRange().getSheet().getIndex() - 1,
      address: namedRange.getRange().getA1Notation(),
      // Sheet scope can only be created by copying a sheet (?)
      scope_sheet_name: namedRange.getName().includes("!")
        ? namedRange.getRange().getSheet().getName()
        : null,
      scope_sheet_index: namedRange.getName().includes("!")
        ? namedRange.getRange().getSheet().getIndex() - 1
        : null,
      book_scope: !namedRange.getName().includes("!"),
    };
  });
  payload["names"] = names;

  payload["sheets"] = [];
  let lastCellCol;
  let lastCellRow;
  let values;
  sheets.forEach((sheet) => {
    lastCellCol = sheet.getLastColumn();
    lastCellRow = sheet.getLastRow();
    if (excludeArray.includes(sheet.getName())) {
      values = [[]];
    } else {
      let range = sheet.getRange(
        1,
        1,
        lastCellRow > 0 ? lastCellRow : 1,
        lastCellCol > 0 ? lastCellCol : 1
      );
      values = range.getValues();
      // Handle dates
      values.forEach((valueRow, rowIndex) => {
        valueRow.forEach((value, colIndex) => {
          if (value instanceof Date) {
            // Convert from script timezone to spreadsheet timezone
            let tzDate = new Date(
              value
                .toLocaleString("en-US", {
                  timeZone: workbook.getSpreadsheetTimeZone(),
                })
                .replace(/\u202F/, " ") // https://bugs.chromium.org/p/v8/issues/detail?id=13494
            );
            // toISOString transforms to UTC, so we need to correct for offset
            values[rowIndex][colIndex] = new Date(
              tzDate.getTime() - tzDate.getTimezoneOffset() * 60 * 1000
            ).toISOString();
          }
        });
      });
    }

    let pictures = [];
    if (excludeArray.includes(sheet.getName())) {
      pictures = [];
    } else {
      sheet.getImages().forEach((image, ix) => {
        pictures[ix] = {
          name: image.getAltTextTitle(),
          height: image.getHeight(),
          width: image.getWidth(),
        };
      });
    }

    payload["sheets"].push({
      name: sheet.getName(),
      values: values,
      pictures: pictures,
      tables: [],
    });
  });

  // console.log(payload);

  // API call
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: headers,
    muteHttpExceptions: true,
  };

  // Parse JSON response
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw response.getContentText();
  }
  const json = response.getContentText();
  const rawData = JSON.parse(json);

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

// Functions map
let funcs = this;

// Functions
function setValues(workbook, action) {
  // Handle DateTime (TODO: backend should deliver indices with datetime obj)
  let dt;
  let dtString;
  let locale = workbook.getSpreadsheetLocale().replace("_", "-");
  action.values.forEach((valueRow, rowIndex) => {
    valueRow.forEach((value, colIndex) => {
      if (
        typeof value === "string" &&
        value.length > 18 &&
        value.includes("T")
      ) {
        dt = new Date(Date.parse(value));
        dtString = dt.toLocaleDateString(locale);
        if (dtString !== "Invalid Date") {
          let hours = dt.getHours();
          let minutes = dt.getMinutes();
          let seconds = dt.getSeconds();
          let milliseconds = dt.getMilliseconds();
          if (hours + minutes + seconds + milliseconds !== 0) {
            // The time doesn't follow the locale in the Date Time combination!
            dtString +=
              " " + hours + ":" + minutes + ":" + seconds + "." + milliseconds;
          }
          action.values[rowIndex][colIndex] = dtString;
        }
      }
    });
  });
  getRange(workbook, action).setValues(action.values);
}

function rangeClearContents(workbook, action) {
  getRange(workbook, action).clearContent();
}

function rangeClearFormats(workbook, action) {
  getRange(workbook, action).clearFormat();
}

function rangeClear(workbook, action) {
  getRange(workbook, action).clear();
}

function addSheet(workbook, action) {
  // insertSheet(sheetName, sheetIndex)
  let sheet = workbook.insertSheet(action.args[1], parseInt(action.args[0]));
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
        action.column_count
      );
  } else {
    workbook
      .getSheets()
      [action.sheet_position].autoResizeRows(
        action.start_row + 1,
        action.row_count
      );
  }
}

function setRangeColor(workbook, action) {
  getRange(workbook, action).setBackground(action.args[0]);
}

function activateSheet(workbook, action) {
  workbook.getSheets()[parseInt(action.args[0])].activate();
}

function addHyperlink(workbook, action) {
  let value = SpreadsheetApp.newRichTextValue()
    .setText(action.args[1])
    .setLinkUrl(action.args[0])
    .build();
  getRange(workbook, action).setRichTextValue(value);
}

function setNumberFormat(workbook, action) {
  getRange(workbook, action).setNumberFormat(action.args[0]);
}

function setPictureName(workbook, action) {
  workbook
    .getSheets()
    [action.sheet_position].getImages()
    [action.args[0]].setAltTextTitle(action.args[1]);
}

function setPictureHeight(workbook, action) {
  workbook
    .getSheets()
    [action.sheet_position].getImages()
    [action.args[0]].setHeight(action.args[1]);
}

function setPictureWidth(workbook, action) {
  workbook
    .getSheets()
    [action.sheet_position].getImages()
    [action.args[0]].setWidth(action.args[1]);
}

function deletePicture(workbook, action) {
  workbook
    .getSheets()
    [action.sheet_position].getImages()
    [action.args[0]].remove();
}

function addPicture(workbook, action) {
  let imageBlob = Utilities.newBlob(
    Utilities.base64Decode(action.args[0]),
    "image/png",
    "MyImageName"
  );
  workbook
    .getSheets()
    [action.sheet_position].insertImage(
      imageBlob,
      action.args[1] + 1,
      action.args[2] + 1
    );
  SpreadsheetApp.flush();
}

function updatePicture(workbook, action) {
  // Workaround as img.replace() doesn't manage to refresh the screen half of the time
  let imageBlob = Utilities.newBlob(
    Utilities.base64Decode(action.args[0]),
    "image/png",
    "MyImageName"
  );
  let img = workbook.getSheets()[action.sheet_position].getImages()[
    action.args[1]
  ];
  let altTextTitle = img.getAltTextTitle();
  let rowIndex = img.getAnchorCell().getRowIndex();
  let colIndex = img.getAnchorCell().getColumnIndex();
  let xOffset = img.getAnchorCellXOffset();
  let yOffset = img.getAnchorCellYOffset();
  let width = img.getWidth();
  let height = img.getHeight();
  // Seems to help if the new image is inserted first before deleting the old one
  imgNew = workbook
    .getSheets()
    [action.sheet_position].insertImage(
      imageBlob,
      colIndex,
      rowIndex,
      xOffset,
      yOffset
    );
  img.remove();
  imgNew.setAltTextTitle(altTextTitle);
  imgNew.setWidth(width);
  imgNew.setHeight(height);
  SpreadsheetApp.flush();
}

function alert(workbook, action) {
  let ui = SpreadsheetApp.getUi();

  let myPrompt = action.args[0];
  let myTitle = action.args[1];
  let myButtons = action.args[2];
  let myMode = action.args[3]; // ignored
  let myCallback = action.args[4];

  if (myButtons == "ok") {
    myButtons = ui.ButtonSet.OK;
  } else if (myButtons == "ok_cancel") {
    myButtons = ui.ButtonSet.OK_CANCEL;
  } else if (myButtons == "yes_no") {
    myButtons = ui.ButtonSet.YES_NO;
  } else if (myButtons == "yes_no_cancel") {
    myButtons = ui.ButtonSet.YES_NO_CANCEL;
  }

  let rv = ui.alert(myTitle, myPrompt, myButtons);

  let buttonResult;
  if (rv == ui.Button.OK) {
    buttonResult = "ok";
  } else if (rv == ui.Button.CANCEL) {
    buttonResult = "cancel";
  } else if (rv == ui.Button.YES) {
    buttonResult = "yes";
  } else if (rv == ui.Button.NO) {
    buttonResult = "no";
  }

  if (myCallback != "") {
    funcs[myCallback](buttonResult);
  }
}

function setRangeName(workbook, action) {
  let range = getRange(workbook, action);
  range.getSheet().getParent().setNamedRange(action.args[0], range);
}

function namesAdd(workbook, action) {
  let name = action.args[0];
  if (name.includes("!")) {
    throw "NotImplemented: sheet scoped names";
  }
  let refersTo = action.args[1];
  const parts = refersTo.split("!");
  const address = parts[1];
  let sheetName = parts[0];
  if (sheetName.charAt(0) === "=") {
    sheetName = sheetName.substring(1);
  }
  if (sheetName.includes(" ")) {
    sheetName = sheetName.replace("'", "").replace("'", "");
  }
  let range = workbook.getSheetByName(sheetName).getRange(address);
  range.getSheet().getParent().setNamedRange(name, range);
}

function nameDelete(workbook, action) {
  // workbook.removeNamedRange(name) doesn't work with sheet scoped names
  function processName(name) {
    if (name.includes("!")) {
      const [sheetName, definedName] = name.split("!");
      if (!sheetName.startsWith("'")) {
        return `'${sheetName}'!${definedName}`;
      }
    }
    return name;
  }
  workbook.getNamedRanges().forEach((namedRange) => {
    if (namedRange.getName() === processName(action.args[0])) {
      namedRange.remove();
      return;
    }
  });
}

function runMacro(workbook, action) {
  funcs[action.args[0]](workbook, ...action.args.slice(1));
}

function rangeDelete(workbook, action) {
  if (action.args[0] === "up") {
    getRange(workbook, action).deleteCells(SpreadsheetApp.Dimension.ROWS);
  } else {
    getRange(workbook, action).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  }
}

function rangeInsert(workbook, action) {
  if (action.args[0] === "down") {
    getRange(workbook, action).insertCells(SpreadsheetApp.Dimension.ROWS);
  } else {
    getRange(workbook, action).insertCells(SpreadsheetApp.Dimension.COLUMNS);
  }
}

function addTable(workbook, action) {
  throw "NotImplemented: addTable";
}

function setTableName(workbook, action) {
  throw "NotImplemented: setTableName";
}

function resizeTable(workbook, action) {
  throw "NotImplemented: resizeTable";
}

function showAutofilterTable(workbook, action) {
  throw "NotImplemented: showAutofilterTable";
}

function showHeadersTable(workbook, action) {
  throw "NotImplemented: showHeadersTable";
}

function showTotalsTable(workbook, action) {
  throw "NotImplemented: showTotalsTable";
}

function setTableStyle(workbook, action) {
  throw "NotImplemented: setTableStyle";
}

function copyRange(workbook, action) {
  const destination = workbook
    .getSheets()
    [parseInt(action.args[0])].getRange(action.args[1].toString());
  getRange(workbook, action).copyTo(destination);
}
