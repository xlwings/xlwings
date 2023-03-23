// core-js polyfills for ie11
import "core-js/actual/object/assign";
import "core-js/actual/array/includes";
import "core-js/actual/global-this";
import "core-js/actual/function/name";
import { xlAlert } from "./alert";

const version = "dev";
globalThis.callbacks = {};
export async function runPython(
  url = "",
  { auth = "", include = "", exclude = "", headers = {} }: Options = {}
) {
  try {
    await Excel.run(async (context) => {
      // workbook
      const workbook = context.workbook;
      workbook.load("name");

      // sheets
      let worksheets = workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();
      let sheets = worksheets.items;

      // Config
      let configSheet = worksheets.getItemOrNullObject("xlwings.conf");
      await context.sync();
      let config = {};
      if (!configSheet.isNullObject) {
        const configRange = configSheet
          .getRange("A1")
          .getSurroundingRegion()
          .load("values");
        await context.sync();
        const configValues = configRange.values;
        configValues.forEach(
          (el) => (config[el[0].toString()] = el[1].toString())
        );
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
          if (!includeArray.includes(sheet.name)) {
            excludeArray.push(sheet.name);
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
      if (!("Authorization" in headers) && auth.length > 0) {
        headers["Authorization"] = auth;
      }

      // Standard headers
      headers["Content-Type"] = "application/json";

      // Request payload
      let payload: {} = {};
      payload["client"] = "Office.js";
      payload["version"] = version;
      let activeSheet = worksheets.getActiveWorksheet().load("position");
      let selection = workbook.getSelectedRange().load("address");
      await context.sync();
      payload["book"] = {
        name: workbook.name,
        active_sheet_index: activeSheet.position,
        selection: selection.address.split("!").pop(),
      };

      // Names (book scope)
      let names: Names[] = [];
      const namedItems = context.workbook.names.load("name, type");
      await context.sync();

      namedItems.items.forEach((namedItem, ix) => {
        // Currently filtering to named ranges
        if (namedItem.type === "Range") {
          names[ix] = {
            name: namedItem.name,
            sheet: namedItem.getRange().worksheet.load("position"),
            range: namedItem.getRange().load("address"),
            book_scope: true, // workbook.names contains only workbook scope!
          };
        }
      });

      await context.sync();

      names.forEach((namedItem, ix) => {
        names[ix] = {
          name: namedItem.name,
          sheet_index: namedItem.sheet.position,
          address: namedItem.range.address.split("!").pop(),
          book_scope: namedItem.book_scope,
        };
      });

      payload["names"] = names;

      // Sheets
      payload["sheets"] = [];
      let sheetsLoader = [];
      sheets.forEach((sheet) => {
        sheet.load("name names");
        let lastCell: Excel.Range;
        if (sheet.getUsedRange() !== undefined) {
          lastCell = sheet.getUsedRange().getLastCell().load("address");
        } else {
          lastCell = sheet.getRange("A1").load("address");
        }
        sheetsLoader.push({
          sheet: sheet,
          lastCell: lastCell,
        });
      });

      await context.sync();

      sheetsLoader.forEach((item, ix) => {
        let range: Excel.Range;
        range = item["sheet"]
          .getRange(`A1:${item["lastCell"].address}`)
          .load("values, numberFormatCategories");
        sheetsLoader[ix]["range"] = range;
        // Names (sheet scope)
        sheetsLoader[ix]["names"] = item["sheet"].names.load("name, type");
      });

      await context.sync();

      // Names (sheet scope)
      let namesSheetScope: Names[] = [];
      sheetsLoader.forEach((item) => {
        item["names"].items.forEach((namedItem, ix) => {
          namesSheetScope[ix] = {
            name: namedItem.name,
            sheet: namedItem.getRange().worksheet.load("position"),
            range: namedItem.getRange().load("address"),
            book_scope: false,
          };
        });
      });

      await context.sync();

      let namesSheetsScope2: Names[] = [];
      namesSheetScope.forEach((namedItem, ix) => {
        namesSheetsScope2[ix] = {
          name: namedItem.name,
          sheet_index: namedItem.sheet.position,
          address: namedItem.range.address.split("!").pop(),
          book_scope: namedItem.book_scope,
        };
      });

      // Add sheet scoped names to book scoped names
      payload["names"] = payload["names"].concat(namesSheetsScope2);

      // values
      for (let item of sheetsLoader) {
        let sheet = item["sheet"]; // TODO: replace item["sheet"] with sheet
        let values;
        if (excludeArray.includes(item["sheet"].name)) {
          values = [[]];
        } else {
          values = item["range"].values;
          if (Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
            // numberFormatCategories requires Excel 2021/365
            // i.e., dates aren't transformed to Python's datetime in Excel <=2019
            let categories = item["range"].numberFormatCategories;
            // Handle dates
            // https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/excel-samples#dates
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
        }
        // Tables
        let tablesArray: Tables[] = [];
        if (!excludeArray.includes(item["sheet"].name)) {
          const tables = sheet.tables.load([
            "name",
            "showHeaders",
            "dataBodyRange",
            "showTotals",
            "style",
            "showFilterButton",
          ]);
          await context.sync();
          let tablesLoader = [];
          for (let table of sheet.tables.items) {
            tablesLoader.push({
              name: table.name,
              showHeaders: table.showHeaders,
              showTotals: table.showTotals,
              style: table.style,
              showFilterButton: table.showFilterButton,
              range: table.getRange().load("address"),
              dataBodyRange: table.getDataBodyRange().load("address"),
              headerRowRange: table.showHeaders
                ? table.getHeaderRowRange().load("address")
                : null,
              totalRowRange: table.showTotals
                ? table.getTotalRowRange().load("address")
                : null,
            });
          }
          await context.sync();
          for (let table of tablesLoader) {
            tablesArray.push({
              name: table.name,
              range_address: table.range.address.split("!").pop(),
              header_row_range_address: table.showHeaders
                ? table.headerRowRange.address.split("!").pop()
                : null,
              data_body_range_address: table.dataBodyRange.address.split("!").pop(),
              total_row_range_address: table.showTotals
                ? table.totalRowRange.address.split("!").pop()
                : null,
              show_headers: table.showHeaders,
              show_totals: table.showTotals,
              table_style: table.style,
              show_autofilter: table.showFilterButton,
            });
          }
        }
        payload["sheets"].push({
          name: item["sheet"].name,
          values: values,
          pictures: [], // TODO
          tables: tablesArray,
        });
      }

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
        for (let action of rawData["actions"]) {
          await globalThis.callbacks[action.func](context, action);
          if (forceSync.some((el) => action.func.toLowerCase().includes(el))) {
            await context.sync();
          }
        }
      }
    });
  } catch (error) {
    console.error(error);
    xlAlert(error, "Error", "ok", "critical", "");
  }
}

// Helpers
interface Options {
  auth?: string;
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
  sheet_index?: number;
  sheet?: Excel.Worksheet;
  range?: Excel.Range;
  address?: string;
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

async function getRange(context: Excel.RequestContext, action: Action) {
  let sheets = context.workbook.worksheets.load("items");
  await context.sync();
  return sheets.items[action["sheet_position"]].getRangeByIndexes(
    action.start_row,
    action.start_column,
    action.row_count,
    action.column_count
  );
}

export function registerCallback(callback: Function) {
  globalThis.callbacks[callback.name] = callback;
}

// Callbacks
async function setValues(context: Excel.RequestContext, action: Action) {
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
        // Excel on macOS doesn't use proper locale if not passed explicitly
        // https://learn.microsoft.com/en-us/office/dev/add-ins/develop/localization#match-datetime-format-with-client-locale
        dtString = dt.toLocaleDateString(Office.context.contentLanguage);
        // Note that adding the time will format the cell as Custom instead of Date/Time
        // which xlwings currently doesn't translate to datetime when reading
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
  let range = await getRange(context, action);
  range.values = action.values;
  await context.sync();
}
registerCallback(setValues);

async function clearContents(context: Excel.RequestContext, action: Action) {
  let range = await getRange(context, action);
  range.clear(Excel.ClearApplyTo.contents);
  await context.sync();
}
registerCallback(clearContents);

async function addSheet(context: Excel.RequestContext, action: Action) {
  let sheet: Excel.Worksheet;
  if (action.args[1] !== null) {
    sheet = context.workbook.worksheets.add(action.args[1].toString());
  } else {
    sheet = context.workbook.worksheets.add();
  }
  sheet.position = parseInt(action.args[0].toString());
}
registerCallback(addSheet);

async function setSheetName(context: Excel.RequestContext, action: Action) {
  let sheets = context.workbook.worksheets.load("items");
  sheets.items[action.sheet_position].name = action.args[0].toString();
}
registerCallback(setSheetName);

async function setAutofit(context: Excel.RequestContext, action: Action) {
  if (action.args[0] === "columns") {
    let range = await getRange(context, action);
    range.format.autofitColumns();
  } else {
    let range = await getRange(context, action);
    range.format.autofitRows();
  }
}
registerCallback(setAutofit);

async function setRangeColor(context: Excel.RequestContext, action: Action) {
  let range = await getRange(context, action);
  range.format.fill.color = action.args[0].toString();
  await context.sync();
}
registerCallback(setRangeColor);

async function activateSheet(context: Excel.RequestContext, action: Action) {
  let worksheets = context.workbook.worksheets;
  worksheets.load("items");
  await context.sync();
  worksheets.items[parseInt(action.args[0].toString())].activate();
}
registerCallback(activateSheet);

async function addHyperlink(context: Excel.RequestContext, action: Action) {
  let range = await getRange(context, action);
  let hyperlink = {
    textToDisplay: action.args[1].toString(),
    screenTip: action.args[2].toString(),
    address: action.args[0].toString(),
  };
  range.hyperlink = hyperlink;
  await context.sync();
}
registerCallback(addHyperlink);

async function setNumberFormat(context: Excel.RequestContext, action: Action) {
  let range = await getRange(context, action);
  range.numberFormat = [[action.args[0].toString()]];
}
registerCallback(setNumberFormat);

async function setPictureName(context: Excel.RequestContext, action: Action) {
  throw "Not Implemented: setPictureName";
}
registerCallback(setPictureName);

async function setPictureHeight(context: Excel.RequestContext, action: Action) {
  throw "Not Implemented: setPictureHeight";
}
registerCallback(setPictureHeight);

async function setPictureWidth(context: Excel.RequestContext, action: Action) {
  throw "Not Implemented: setPictureWidth";
}
registerCallback(setPictureWidth);

async function deletePicture(context: Excel.RequestContext, action: Action) {
  throw "Not Implemented: deletePicture";
}
registerCallback(deletePicture);

async function addPicture(context: Excel.RequestContext, action: Action) {
  throw "Not Implemented: addPicture";
}
registerCallback(addPicture);

async function updatePicture(context: Excel.RequestContext, action: Action) {
  throw "Not Implemented: updatePicture";
}
registerCallback(updatePicture);

async function alert(context: Excel.RequestContext, action: Action) {
  let myPrompt = action.args[0].toString();
  let myTitle = action.args[1].toString();
  let myButtons = action.args[2].toString();
  let myMode = action.args[3].toString();
  let myCallback = action.args[4].toString();
  xlAlert(myPrompt, myTitle, myButtons, myMode, myCallback);
}
registerCallback(alert);

async function setRangeName(context: Excel.RequestContext, action: Action) {
  throw "NotImplemented: setRangeName";
}
registerCallback(setRangeName);

async function namesAdd(context: Excel.RequestContext, action: Action) {
  throw "NotImplemented: namesAdd";
}
registerCallback(namesAdd);

async function nameDelete(context: Excel.RequestContext, action: Action) {
  throw "NotImplemented: deleteName";
}
registerCallback(nameDelete);

async function runMacro(context: Excel.RequestContext, action: Action) {
  await globalThis.callbacks[action.args[0].toString()](
    context,
    ...action.args.slice(1)
  );
}
registerCallback(runMacro);

async function rangeDelete(context: Excel.RequestContext, action: Action) {
  let range = await getRange(context, action);
  let shift = action.args[0].toString();
  if (shift === "up") {
    range.delete(Excel.DeleteShiftDirection.up);
  } else if (shift === "left") {
    range.delete(Excel.DeleteShiftDirection.left);
  }
}
registerCallback(rangeDelete);
