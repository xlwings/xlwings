function hello() {
  runPython("url", { apiKey: "API_KEY" });
}

/**
 * xlwings dev (for Google Apps Script)
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
  { apiKey = "", include = "", exclude = "", headers = {} } = {}
) {
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
  if (!("Authorization" in headers)) {
    headers["Authorization"] = apiKey;
  }

  // Request payload
  let payload = {};
  payload["client"] = "Google Apps Script";
  payload["version"] = "dev";
  payload["book"] = {
    name: workbook.getName(),
    active_sheet_index: workbook.getActiveSheet().getIndex() - 1,
    selection: workbook.getActiveRange().getA1Notation(),
  };
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
              value.toLocaleString("en-US", {
                timeZone: workbook.getSpreadsheetTimeZone(),
              })
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
    workbook.getImages().forEach((image, ix) => {
      pictures[ix] = {
        name: image.getAltTextTitle(),
        height: image.getHeight(),
        width: image.getWidth(),
      };
    });

    payload["sheets"].push({
      name: sheet.getName(),
      values: values,
      pictures: pictures,
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
  deletePicture: deletePicture,
  addPicture: addPicture,
};

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
