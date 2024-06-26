/**
* Copyright (C) 2014 - present, Zoomer Analytics GmbH. All rights reserved.
* Licensed under BSD-3-Clause license, see: https://docs.xlwings.org/en/stable/license.html
*
* This file also contains code from core-js
* Copyright (C) 2014-2023 Denis Pushkarev, Licensed under MIT license, see https://raw.githubusercontent.com/zloirock/core-js/master/LICENSE
* This file also contains code from Webpack
* Copyright (C) JS Foundation and other contributors, Licensed under MIT license, see https://raw.githubusercontent.com/webpack/webpack/main/LICENSE
*/
var xlwings;
/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/alert.ts":
/*!**********************!*\
  !*** ./src/alert.ts ***!
  \**********************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   xlAlert: function() { return /* binding */ xlAlert; }
/* harmony export */ });
// https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var dialog;
function dialogCallback(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("".concat(asyncResult.error.message, " [").concat(asyncResult.error.code, "]"));
    }
    else {
        dialog = asyncResult.value;
        // Handle messages and events
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
}
function processMessage(arg) {
    dialog.close();
    var _a = arg.message.split("|"), selection = _a[0], callback = _a[1];
    if (callback !== "" && callback in globalThis.callbacks) {
        globalThis.callbacks[callback](selection);
    }
    else {
        if (callback !== "" && !(callback in globalThis.callbacks)) {
            throw new Error("Didn't find callback '".concat(callback, "'! Make sure to run xlwings.registerCallback(").concat(callback, ") before calling runPython."));
        }
    }
}
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            console.log("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            console.log("HTTPS is required.");
            break;
        case 12006:
            console.log("Dialog closed by user");
            break;
        default:
            console.log("Unknown error in dialog box");
            break;
    }
}
function xlAlert(prompt, title, buttons, mode, callback) {
    return __awaiter(this, void 0, void 0, function () {
        var width, height, appPathElement, appPath;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Office.onReady()];
                case 1:
                    _a.sent();
                    if (Office.context.platform.toString() === "OfficeOnline") {
                        width = 28;
                        height = 36;
                    }
                    else if (Office.context.platform.toString() === "PC") {
                        width = 28; // seems to have a wider min width
                        height = 40;
                    }
                    else {
                        width = 32;
                        height = 30;
                    }
                    appPathElement = document.getElementById("app-path");
                    appPath = appPathElement ? JSON.parse(appPathElement.textContent) : null;
                    Office.context.ui.displayDialogAsync(window.location.origin +
                        (appPath && appPath.appPath !== "" ? "/".concat(appPath.appPath) : "") +
                        "/xlwings/alert?prompt=" +
                        encodeURIComponent("".concat(prompt)) +
                        "&title=" +
                        encodeURIComponent("".concat(title)) +
                        "&buttons=".concat(buttons, "&mode=").concat(mode, "&callback=").concat(callback), { height: height, width: width, displayInIframe: true }, dialogCallback);
                    return [2 /*return*/];
            }
        });
    });
}


/***/ }),

/***/ "./src/auth.ts":
/*!*********************!*\
  !*** ./src/auth.ts ***!
  \*********************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getAccessToken: function() { return /* binding */ getAccessToken; }
/* harmony export */ });
// Office.auth.getAccessToken claims that it does everything that this module does,
// only it doesn't: https://github.com/OfficeDev/office-js/issues/3298
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var accessToken = null;
var isRenewingToken = false;
var tokenLock = false;
var tokenExpiry = null;
function hasKeyExpired() {
    if (!tokenExpiry) {
        return true;
    }
    var currentTime = Math.floor(Date.now() / 1000); // Convert to seconds
    // Renew 15 minutes before expiry
    return currentTime >= tokenExpiry - 15 * 60;
}
function renewAccessToken() {
    return __awaiter(this, void 0, void 0, function () {
        var payload, base64, decodedPayload, error_1, token_error;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("Renewing access token");
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, Office.auth.getAccessToken({
                            allowSignInPrompt: true,
                            allowConsentPrompt: true,
                        })];
                case 2:
                    accessToken = _a.sent();
                    payload = accessToken.split(".")[1];
                    base64 = payload.replace(/-/g, "+").replace(/_/g, "/");
                    while (base64.length % 4) {
                        base64 += "=";
                    }
                    decodedPayload = JSON.parse(window.atob(base64));
                    tokenExpiry = decodedPayload.exp;
                    accessToken = "Bearer " + accessToken;
                    return [3 /*break*/, 5];
                case 3:
                    error_1 = _a.sent();
                    token_error = "Error ".concat(error_1.code, ": ").concat(error_1.message);
                    console.log(token_error);
                    // return token error so it can be logged on backend
                    accessToken = token_error;
                    return [3 /*break*/, 5];
                case 4:
                    tokenLock = false;
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    });
}
function getAccessToken() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Office.onReady()];
                case 1:
                    _a.sent();
                    if (!(!accessToken || hasKeyExpired())) return [3 /*break*/, 5];
                    if (!!tokenLock) return [3 /*break*/, 3];
                    tokenLock = true;
                    isRenewingToken = true;
                    return [4 /*yield*/, renewAccessToken()];
                case 2:
                    _a.sent();
                    isRenewingToken = false;
                    return [3 /*break*/, 5];
                case 3:
                    if (!isRenewingToken) return [3 /*break*/, 5];
                    return [4 /*yield*/, new Promise(function (resolve) { return setTimeout(resolve, 100); })];
                case 4:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 5: return [2 /*return*/, accessToken];
            }
        });
    });
}


/***/ }),

/***/ "./src/utils.ts":
/*!**********************!*\
  !*** ./src/utils.ts ***!
  \**********************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getActiveBookName: function() { return /* binding */ getActiveBookName; }
/* harmony export */ });
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
function getActiveBookName() {
    return __awaiter(this, void 0, void 0, function () {
        var error_1;
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 3, , 4]);
                    return [4 /*yield*/, Office.onReady()];
                case 1:
                    _a.sent();
                    return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                            var workbook;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        workbook = context.workbook;
                                        workbook.load("name");
                                        return [4 /*yield*/, context.sync()];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/, workbook.name];
                                }
                            });
                        }); })];
                case 2: return [2 /*return*/, _a.sent()];
                case 3:
                    error_1 = _a.sent();
                    console.error(error_1);
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    });
}


/***/ }),

/***/ "./node_modules/core-js/actual/array/includes.js":
/*!*******************************************************!*\
  !*** ./node_modules/core-js/actual/array/includes.js ***!
  \*******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../../stable/array/includes */ "./node_modules/core-js/stable/array/includes.js");

module.exports = parent;


/***/ }),

/***/ "./node_modules/core-js/actual/function/name.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/actual/function/name.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../../stable/function/name */ "./node_modules/core-js/stable/function/name.js");

module.exports = parent;


/***/ }),

/***/ "./node_modules/core-js/actual/global-this.js":
/*!****************************************************!*\
  !*** ./node_modules/core-js/actual/global-this.js ***!
  \****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../stable/global-this */ "./node_modules/core-js/stable/global-this.js");

module.exports = parent;


/***/ }),

/***/ "./node_modules/core-js/actual/object/assign.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/actual/object/assign.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../../stable/object/assign */ "./node_modules/core-js/stable/object/assign.js");

module.exports = parent;


/***/ }),

/***/ "./node_modules/core-js/es/array/includes.js":
/*!***************************************************!*\
  !*** ./node_modules/core-js/es/array/includes.js ***!
  \***************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


__webpack_require__(/*! ../../modules/es.array.includes */ "./node_modules/core-js/modules/es.array.includes.js");
var entryUnbind = __webpack_require__(/*! ../../internals/entry-unbind */ "./node_modules/core-js/internals/entry-unbind.js");

module.exports = entryUnbind('Array', 'includes');


/***/ }),

/***/ "./node_modules/core-js/es/function/name.js":
/*!**************************************************!*\
  !*** ./node_modules/core-js/es/function/name.js ***!
  \**************************************************/
/***/ (function(__unused_webpack_module, __unused_webpack_exports, __webpack_require__) {


__webpack_require__(/*! ../../modules/es.function.name */ "./node_modules/core-js/modules/es.function.name.js");


/***/ }),

/***/ "./node_modules/core-js/es/global-this.js":
/*!************************************************!*\
  !*** ./node_modules/core-js/es/global-this.js ***!
  \************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


__webpack_require__(/*! ../modules/es.global-this */ "./node_modules/core-js/modules/es.global-this.js");

module.exports = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");


/***/ }),

/***/ "./node_modules/core-js/es/object/assign.js":
/*!**************************************************!*\
  !*** ./node_modules/core-js/es/object/assign.js ***!
  \**************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


__webpack_require__(/*! ../../modules/es.object.assign */ "./node_modules/core-js/modules/es.object.assign.js");
var path = __webpack_require__(/*! ../../internals/path */ "./node_modules/core-js/internals/path.js");

module.exports = path.Object.assign;


/***/ }),

/***/ "./node_modules/core-js/internals/a-callable.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/internals/a-callable.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");
var tryToString = __webpack_require__(/*! ../internals/try-to-string */ "./node_modules/core-js/internals/try-to-string.js");

var $TypeError = TypeError;

// `Assert: IsCallable(argument) is true`
module.exports = function (argument) {
  if (isCallable(argument)) return argument;
  throw new $TypeError(tryToString(argument) + ' is not a function');
};


/***/ }),

/***/ "./node_modules/core-js/internals/add-to-unscopables.js":
/*!**************************************************************!*\
  !*** ./node_modules/core-js/internals/add-to-unscopables.js ***!
  \**************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var wellKnownSymbol = __webpack_require__(/*! ../internals/well-known-symbol */ "./node_modules/core-js/internals/well-known-symbol.js");
var create = __webpack_require__(/*! ../internals/object-create */ "./node_modules/core-js/internals/object-create.js");
var defineProperty = (__webpack_require__(/*! ../internals/object-define-property */ "./node_modules/core-js/internals/object-define-property.js").f);

var UNSCOPABLES = wellKnownSymbol('unscopables');
var ArrayPrototype = Array.prototype;

// Array.prototype[@@unscopables]
// https://tc39.es/ecma262/#sec-array.prototype-@@unscopables
if (ArrayPrototype[UNSCOPABLES] === undefined) {
  defineProperty(ArrayPrototype, UNSCOPABLES, {
    configurable: true,
    value: create(null)
  });
}

// add a key to Array.prototype[@@unscopables]
module.exports = function (key) {
  ArrayPrototype[UNSCOPABLES][key] = true;
};


/***/ }),

/***/ "./node_modules/core-js/internals/an-object.js":
/*!*****************************************************!*\
  !*** ./node_modules/core-js/internals/an-object.js ***!
  \*****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var isObject = __webpack_require__(/*! ../internals/is-object */ "./node_modules/core-js/internals/is-object.js");

var $String = String;
var $TypeError = TypeError;

// `Assert: Type(argument) is Object`
module.exports = function (argument) {
  if (isObject(argument)) return argument;
  throw new $TypeError($String(argument) + ' is not an object');
};


/***/ }),

/***/ "./node_modules/core-js/internals/array-includes.js":
/*!**********************************************************!*\
  !*** ./node_modules/core-js/internals/array-includes.js ***!
  \**********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var toIndexedObject = __webpack_require__(/*! ../internals/to-indexed-object */ "./node_modules/core-js/internals/to-indexed-object.js");
var toAbsoluteIndex = __webpack_require__(/*! ../internals/to-absolute-index */ "./node_modules/core-js/internals/to-absolute-index.js");
var lengthOfArrayLike = __webpack_require__(/*! ../internals/length-of-array-like */ "./node_modules/core-js/internals/length-of-array-like.js");

// `Array.prototype.{ indexOf, includes }` methods implementation
var createMethod = function (IS_INCLUDES) {
  return function ($this, el, fromIndex) {
    var O = toIndexedObject($this);
    var length = lengthOfArrayLike(O);
    if (length === 0) return !IS_INCLUDES && -1;
    var index = toAbsoluteIndex(fromIndex, length);
    var value;
    // Array#includes uses SameValueZero equality algorithm
    // eslint-disable-next-line no-self-compare -- NaN check
    if (IS_INCLUDES && el !== el) while (length > index) {
      value = O[index++];
      // eslint-disable-next-line no-self-compare -- NaN check
      if (value !== value) return true;
    // Array#indexOf ignores holes, Array#includes - not
    } else for (;length > index; index++) {
      if ((IS_INCLUDES || index in O) && O[index] === el) return IS_INCLUDES || index || 0;
    } return !IS_INCLUDES && -1;
  };
};

module.exports = {
  // `Array.prototype.includes` method
  // https://tc39.es/ecma262/#sec-array.prototype.includes
  includes: createMethod(true),
  // `Array.prototype.indexOf` method
  // https://tc39.es/ecma262/#sec-array.prototype.indexof
  indexOf: createMethod(false)
};


/***/ }),

/***/ "./node_modules/core-js/internals/classof-raw.js":
/*!*******************************************************!*\
  !*** ./node_modules/core-js/internals/classof-raw.js ***!
  \*******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");

var toString = uncurryThis({}.toString);
var stringSlice = uncurryThis(''.slice);

module.exports = function (it) {
  return stringSlice(toString(it), 8, -1);
};


/***/ }),

/***/ "./node_modules/core-js/internals/copy-constructor-properties.js":
/*!***********************************************************************!*\
  !*** ./node_modules/core-js/internals/copy-constructor-properties.js ***!
  \***********************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var hasOwn = __webpack_require__(/*! ../internals/has-own-property */ "./node_modules/core-js/internals/has-own-property.js");
var ownKeys = __webpack_require__(/*! ../internals/own-keys */ "./node_modules/core-js/internals/own-keys.js");
var getOwnPropertyDescriptorModule = __webpack_require__(/*! ../internals/object-get-own-property-descriptor */ "./node_modules/core-js/internals/object-get-own-property-descriptor.js");
var definePropertyModule = __webpack_require__(/*! ../internals/object-define-property */ "./node_modules/core-js/internals/object-define-property.js");

module.exports = function (target, source, exceptions) {
  var keys = ownKeys(source);
  var defineProperty = definePropertyModule.f;
  var getOwnPropertyDescriptor = getOwnPropertyDescriptorModule.f;
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (!hasOwn(target, key) && !(exceptions && hasOwn(exceptions, key))) {
      defineProperty(target, key, getOwnPropertyDescriptor(source, key));
    }
  }
};


/***/ }),

/***/ "./node_modules/core-js/internals/create-non-enumerable-property.js":
/*!**************************************************************************!*\
  !*** ./node_modules/core-js/internals/create-non-enumerable-property.js ***!
  \**************************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var definePropertyModule = __webpack_require__(/*! ../internals/object-define-property */ "./node_modules/core-js/internals/object-define-property.js");
var createPropertyDescriptor = __webpack_require__(/*! ../internals/create-property-descriptor */ "./node_modules/core-js/internals/create-property-descriptor.js");

module.exports = DESCRIPTORS ? function (object, key, value) {
  return definePropertyModule.f(object, key, createPropertyDescriptor(1, value));
} : function (object, key, value) {
  object[key] = value;
  return object;
};


/***/ }),

/***/ "./node_modules/core-js/internals/create-property-descriptor.js":
/*!**********************************************************************!*\
  !*** ./node_modules/core-js/internals/create-property-descriptor.js ***!
  \**********************************************************************/
/***/ (function(module) {


module.exports = function (bitmap, value) {
  return {
    enumerable: !(bitmap & 1),
    configurable: !(bitmap & 2),
    writable: !(bitmap & 4),
    value: value
  };
};


/***/ }),

/***/ "./node_modules/core-js/internals/define-built-in-accessor.js":
/*!********************************************************************!*\
  !*** ./node_modules/core-js/internals/define-built-in-accessor.js ***!
  \********************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var makeBuiltIn = __webpack_require__(/*! ../internals/make-built-in */ "./node_modules/core-js/internals/make-built-in.js");
var defineProperty = __webpack_require__(/*! ../internals/object-define-property */ "./node_modules/core-js/internals/object-define-property.js");

module.exports = function (target, name, descriptor) {
  if (descriptor.get) makeBuiltIn(descriptor.get, name, { getter: true });
  if (descriptor.set) makeBuiltIn(descriptor.set, name, { setter: true });
  return defineProperty.f(target, name, descriptor);
};


/***/ }),

/***/ "./node_modules/core-js/internals/define-built-in.js":
/*!***********************************************************!*\
  !*** ./node_modules/core-js/internals/define-built-in.js ***!
  \***********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");
var definePropertyModule = __webpack_require__(/*! ../internals/object-define-property */ "./node_modules/core-js/internals/object-define-property.js");
var makeBuiltIn = __webpack_require__(/*! ../internals/make-built-in */ "./node_modules/core-js/internals/make-built-in.js");
var defineGlobalProperty = __webpack_require__(/*! ../internals/define-global-property */ "./node_modules/core-js/internals/define-global-property.js");

module.exports = function (O, key, value, options) {
  if (!options) options = {};
  var simple = options.enumerable;
  var name = options.name !== undefined ? options.name : key;
  if (isCallable(value)) makeBuiltIn(value, name, options);
  if (options.global) {
    if (simple) O[key] = value;
    else defineGlobalProperty(key, value);
  } else {
    try {
      if (!options.unsafe) delete O[key];
      else if (O[key]) simple = true;
    } catch (error) { /* empty */ }
    if (simple) O[key] = value;
    else definePropertyModule.f(O, key, {
      value: value,
      enumerable: false,
      configurable: !options.nonConfigurable,
      writable: !options.nonWritable
    });
  } return O;
};


/***/ }),

/***/ "./node_modules/core-js/internals/define-global-property.js":
/*!******************************************************************!*\
  !*** ./node_modules/core-js/internals/define-global-property.js ***!
  \******************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");

// eslint-disable-next-line es/no-object-defineproperty -- safe
var defineProperty = Object.defineProperty;

module.exports = function (key, value) {
  try {
    defineProperty(global, key, { value: value, configurable: true, writable: true });
  } catch (error) {
    global[key] = value;
  } return value;
};


/***/ }),

/***/ "./node_modules/core-js/internals/descriptors.js":
/*!*******************************************************!*\
  !*** ./node_modules/core-js/internals/descriptors.js ***!
  \*******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");

// Detect IE8's incomplete defineProperty implementation
module.exports = !fails(function () {
  // eslint-disable-next-line es/no-object-defineproperty -- required for testing
  return Object.defineProperty({}, 1, { get: function () { return 7; } })[1] !== 7;
});


/***/ }),

/***/ "./node_modules/core-js/internals/document-create-element.js":
/*!*******************************************************************!*\
  !*** ./node_modules/core-js/internals/document-create-element.js ***!
  \*******************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var isObject = __webpack_require__(/*! ../internals/is-object */ "./node_modules/core-js/internals/is-object.js");

var document = global.document;
// typeof document.createElement is 'object' in old IE
var EXISTS = isObject(document) && isObject(document.createElement);

module.exports = function (it) {
  return EXISTS ? document.createElement(it) : {};
};


/***/ }),

/***/ "./node_modules/core-js/internals/engine-user-agent.js":
/*!*************************************************************!*\
  !*** ./node_modules/core-js/internals/engine-user-agent.js ***!
  \*************************************************************/
/***/ (function(module) {


module.exports = typeof navigator != 'undefined' && String(navigator.userAgent) || '';


/***/ }),

/***/ "./node_modules/core-js/internals/engine-v8-version.js":
/*!*************************************************************!*\
  !*** ./node_modules/core-js/internals/engine-v8-version.js ***!
  \*************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var userAgent = __webpack_require__(/*! ../internals/engine-user-agent */ "./node_modules/core-js/internals/engine-user-agent.js");

var process = global.process;
var Deno = global.Deno;
var versions = process && process.versions || Deno && Deno.version;
var v8 = versions && versions.v8;
var match, version;

if (v8) {
  match = v8.split('.');
  // in old Chrome, versions of V8 isn't V8 = Chrome / 10
  // but their correct versions are not interesting for us
  version = match[0] > 0 && match[0] < 4 ? 1 : +(match[0] + match[1]);
}

// BrowserFS NodeJS `process` polyfill incorrectly set `.v8` to `0.0`
// so check `userAgent` even if `.v8` exists, but 0
if (!version && userAgent) {
  match = userAgent.match(/Edge\/(\d+)/);
  if (!match || match[1] >= 74) {
    match = userAgent.match(/Chrome\/(\d+)/);
    if (match) version = +match[1];
  }
}

module.exports = version;


/***/ }),

/***/ "./node_modules/core-js/internals/entry-unbind.js":
/*!********************************************************!*\
  !*** ./node_modules/core-js/internals/entry-unbind.js ***!
  \********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");

module.exports = function (CONSTRUCTOR, METHOD) {
  return uncurryThis(global[CONSTRUCTOR].prototype[METHOD]);
};


/***/ }),

/***/ "./node_modules/core-js/internals/enum-bug-keys.js":
/*!*********************************************************!*\
  !*** ./node_modules/core-js/internals/enum-bug-keys.js ***!
  \*********************************************************/
/***/ (function(module) {


// IE8- don't enum bug keys
module.exports = [
  'constructor',
  'hasOwnProperty',
  'isPrototypeOf',
  'propertyIsEnumerable',
  'toLocaleString',
  'toString',
  'valueOf'
];


/***/ }),

/***/ "./node_modules/core-js/internals/export.js":
/*!**************************************************!*\
  !*** ./node_modules/core-js/internals/export.js ***!
  \**************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var getOwnPropertyDescriptor = (__webpack_require__(/*! ../internals/object-get-own-property-descriptor */ "./node_modules/core-js/internals/object-get-own-property-descriptor.js").f);
var createNonEnumerableProperty = __webpack_require__(/*! ../internals/create-non-enumerable-property */ "./node_modules/core-js/internals/create-non-enumerable-property.js");
var defineBuiltIn = __webpack_require__(/*! ../internals/define-built-in */ "./node_modules/core-js/internals/define-built-in.js");
var defineGlobalProperty = __webpack_require__(/*! ../internals/define-global-property */ "./node_modules/core-js/internals/define-global-property.js");
var copyConstructorProperties = __webpack_require__(/*! ../internals/copy-constructor-properties */ "./node_modules/core-js/internals/copy-constructor-properties.js");
var isForced = __webpack_require__(/*! ../internals/is-forced */ "./node_modules/core-js/internals/is-forced.js");

/*
  options.target         - name of the target object
  options.global         - target is the global object
  options.stat           - export as static methods of target
  options.proto          - export as prototype methods of target
  options.real           - real prototype method for the `pure` version
  options.forced         - export even if the native feature is available
  options.bind           - bind methods to the target, required for the `pure` version
  options.wrap           - wrap constructors to preventing global pollution, required for the `pure` version
  options.unsafe         - use the simple assignment of property instead of delete + defineProperty
  options.sham           - add a flag to not completely full polyfills
  options.enumerable     - export as enumerable property
  options.dontCallGetSet - prevent calling a getter on target
  options.name           - the .name of the function if it does not match the key
*/
module.exports = function (options, source) {
  var TARGET = options.target;
  var GLOBAL = options.global;
  var STATIC = options.stat;
  var FORCED, target, key, targetProperty, sourceProperty, descriptor;
  if (GLOBAL) {
    target = global;
  } else if (STATIC) {
    target = global[TARGET] || defineGlobalProperty(TARGET, {});
  } else {
    target = global[TARGET] && global[TARGET].prototype;
  }
  if (target) for (key in source) {
    sourceProperty = source[key];
    if (options.dontCallGetSet) {
      descriptor = getOwnPropertyDescriptor(target, key);
      targetProperty = descriptor && descriptor.value;
    } else targetProperty = target[key];
    FORCED = isForced(GLOBAL ? key : TARGET + (STATIC ? '.' : '#') + key, options.forced);
    // contained in target
    if (!FORCED && targetProperty !== undefined) {
      if (typeof sourceProperty == typeof targetProperty) continue;
      copyConstructorProperties(sourceProperty, targetProperty);
    }
    // add a flag to not completely full polyfills
    if (options.sham || (targetProperty && targetProperty.sham)) {
      createNonEnumerableProperty(sourceProperty, 'sham', true);
    }
    defineBuiltIn(target, key, sourceProperty, options);
  }
};


/***/ }),

/***/ "./node_modules/core-js/internals/fails.js":
/*!*************************************************!*\
  !*** ./node_modules/core-js/internals/fails.js ***!
  \*************************************************/
/***/ (function(module) {


module.exports = function (exec) {
  try {
    return !!exec();
  } catch (error) {
    return true;
  }
};


/***/ }),

/***/ "./node_modules/core-js/internals/function-bind-native.js":
/*!****************************************************************!*\
  !*** ./node_modules/core-js/internals/function-bind-native.js ***!
  \****************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");

module.exports = !fails(function () {
  // eslint-disable-next-line es/no-function-prototype-bind -- safe
  var test = (function () { /* empty */ }).bind();
  // eslint-disable-next-line no-prototype-builtins -- safe
  return typeof test != 'function' || test.hasOwnProperty('prototype');
});


/***/ }),

/***/ "./node_modules/core-js/internals/function-call.js":
/*!*********************************************************!*\
  !*** ./node_modules/core-js/internals/function-call.js ***!
  \*********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var NATIVE_BIND = __webpack_require__(/*! ../internals/function-bind-native */ "./node_modules/core-js/internals/function-bind-native.js");

var call = Function.prototype.call;

module.exports = NATIVE_BIND ? call.bind(call) : function () {
  return call.apply(call, arguments);
};


/***/ }),

/***/ "./node_modules/core-js/internals/function-name.js":
/*!*********************************************************!*\
  !*** ./node_modules/core-js/internals/function-name.js ***!
  \*********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var hasOwn = __webpack_require__(/*! ../internals/has-own-property */ "./node_modules/core-js/internals/has-own-property.js");

var FunctionPrototype = Function.prototype;
// eslint-disable-next-line es/no-object-getownpropertydescriptor -- safe
var getDescriptor = DESCRIPTORS && Object.getOwnPropertyDescriptor;

var EXISTS = hasOwn(FunctionPrototype, 'name');
// additional protection from minified / mangled / dropped function names
var PROPER = EXISTS && (function something() { /* empty */ }).name === 'something';
var CONFIGURABLE = EXISTS && (!DESCRIPTORS || (DESCRIPTORS && getDescriptor(FunctionPrototype, 'name').configurable));

module.exports = {
  EXISTS: EXISTS,
  PROPER: PROPER,
  CONFIGURABLE: CONFIGURABLE
};


/***/ }),

/***/ "./node_modules/core-js/internals/function-uncurry-this.js":
/*!*****************************************************************!*\
  !*** ./node_modules/core-js/internals/function-uncurry-this.js ***!
  \*****************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var NATIVE_BIND = __webpack_require__(/*! ../internals/function-bind-native */ "./node_modules/core-js/internals/function-bind-native.js");

var FunctionPrototype = Function.prototype;
var call = FunctionPrototype.call;
var uncurryThisWithBind = NATIVE_BIND && FunctionPrototype.bind.bind(call, call);

module.exports = NATIVE_BIND ? uncurryThisWithBind : function (fn) {
  return function () {
    return call.apply(fn, arguments);
  };
};


/***/ }),

/***/ "./node_modules/core-js/internals/get-built-in.js":
/*!********************************************************!*\
  !*** ./node_modules/core-js/internals/get-built-in.js ***!
  \********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");

var aFunction = function (argument) {
  return isCallable(argument) ? argument : undefined;
};

module.exports = function (namespace, method) {
  return arguments.length < 2 ? aFunction(global[namespace]) : global[namespace] && global[namespace][method];
};


/***/ }),

/***/ "./node_modules/core-js/internals/get-method.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/internals/get-method.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var aCallable = __webpack_require__(/*! ../internals/a-callable */ "./node_modules/core-js/internals/a-callable.js");
var isNullOrUndefined = __webpack_require__(/*! ../internals/is-null-or-undefined */ "./node_modules/core-js/internals/is-null-or-undefined.js");

// `GetMethod` abstract operation
// https://tc39.es/ecma262/#sec-getmethod
module.exports = function (V, P) {
  var func = V[P];
  return isNullOrUndefined(func) ? undefined : aCallable(func);
};


/***/ }),

/***/ "./node_modules/core-js/internals/global.js":
/*!**************************************************!*\
  !*** ./node_modules/core-js/internals/global.js ***!
  \**************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var check = function (it) {
  return it && it.Math === Math && it;
};

// https://github.com/zloirock/core-js/issues/86#issuecomment-115759028
module.exports =
  // eslint-disable-next-line es/no-global-this -- safe
  check(typeof globalThis == 'object' && globalThis) ||
  check(typeof window == 'object' && window) ||
  // eslint-disable-next-line no-restricted-globals -- safe
  check(typeof self == 'object' && self) ||
  check(typeof __webpack_require__.g == 'object' && __webpack_require__.g) ||
  check(typeof this == 'object' && this) ||
  // eslint-disable-next-line no-new-func -- fallback
  (function () { return this; })() || Function('return this')();


/***/ }),

/***/ "./node_modules/core-js/internals/has-own-property.js":
/*!************************************************************!*\
  !*** ./node_modules/core-js/internals/has-own-property.js ***!
  \************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var toObject = __webpack_require__(/*! ../internals/to-object */ "./node_modules/core-js/internals/to-object.js");

var hasOwnProperty = uncurryThis({}.hasOwnProperty);

// `HasOwnProperty` abstract operation
// https://tc39.es/ecma262/#sec-hasownproperty
// eslint-disable-next-line es/no-object-hasown -- safe
module.exports = Object.hasOwn || function hasOwn(it, key) {
  return hasOwnProperty(toObject(it), key);
};


/***/ }),

/***/ "./node_modules/core-js/internals/hidden-keys.js":
/*!*******************************************************!*\
  !*** ./node_modules/core-js/internals/hidden-keys.js ***!
  \*******************************************************/
/***/ (function(module) {


module.exports = {};


/***/ }),

/***/ "./node_modules/core-js/internals/html.js":
/*!************************************************!*\
  !*** ./node_modules/core-js/internals/html.js ***!
  \************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var getBuiltIn = __webpack_require__(/*! ../internals/get-built-in */ "./node_modules/core-js/internals/get-built-in.js");

module.exports = getBuiltIn('document', 'documentElement');


/***/ }),

/***/ "./node_modules/core-js/internals/ie8-dom-define.js":
/*!**********************************************************!*\
  !*** ./node_modules/core-js/internals/ie8-dom-define.js ***!
  \**********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");
var createElement = __webpack_require__(/*! ../internals/document-create-element */ "./node_modules/core-js/internals/document-create-element.js");

// Thanks to IE8 for its funny defineProperty
module.exports = !DESCRIPTORS && !fails(function () {
  // eslint-disable-next-line es/no-object-defineproperty -- required for testing
  return Object.defineProperty(createElement('div'), 'a', {
    get: function () { return 7; }
  }).a !== 7;
});


/***/ }),

/***/ "./node_modules/core-js/internals/indexed-object.js":
/*!**********************************************************!*\
  !*** ./node_modules/core-js/internals/indexed-object.js ***!
  \**********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");
var classof = __webpack_require__(/*! ../internals/classof-raw */ "./node_modules/core-js/internals/classof-raw.js");

var $Object = Object;
var split = uncurryThis(''.split);

// fallback for non-array-like ES3 and non-enumerable old V8 strings
module.exports = fails(function () {
  // throws an error in rhino, see https://github.com/mozilla/rhino/issues/346
  // eslint-disable-next-line no-prototype-builtins -- safe
  return !$Object('z').propertyIsEnumerable(0);
}) ? function (it) {
  return classof(it) === 'String' ? split(it, '') : $Object(it);
} : $Object;


/***/ }),

/***/ "./node_modules/core-js/internals/inspect-source.js":
/*!**********************************************************!*\
  !*** ./node_modules/core-js/internals/inspect-source.js ***!
  \**********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");
var store = __webpack_require__(/*! ../internals/shared-store */ "./node_modules/core-js/internals/shared-store.js");

var functionToString = uncurryThis(Function.toString);

// this helper broken in `core-js@3.4.1-3.4.4`, so we can't use `shared` helper
if (!isCallable(store.inspectSource)) {
  store.inspectSource = function (it) {
    return functionToString(it);
  };
}

module.exports = store.inspectSource;


/***/ }),

/***/ "./node_modules/core-js/internals/internal-state.js":
/*!**********************************************************!*\
  !*** ./node_modules/core-js/internals/internal-state.js ***!
  \**********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var NATIVE_WEAK_MAP = __webpack_require__(/*! ../internals/weak-map-basic-detection */ "./node_modules/core-js/internals/weak-map-basic-detection.js");
var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var isObject = __webpack_require__(/*! ../internals/is-object */ "./node_modules/core-js/internals/is-object.js");
var createNonEnumerableProperty = __webpack_require__(/*! ../internals/create-non-enumerable-property */ "./node_modules/core-js/internals/create-non-enumerable-property.js");
var hasOwn = __webpack_require__(/*! ../internals/has-own-property */ "./node_modules/core-js/internals/has-own-property.js");
var shared = __webpack_require__(/*! ../internals/shared-store */ "./node_modules/core-js/internals/shared-store.js");
var sharedKey = __webpack_require__(/*! ../internals/shared-key */ "./node_modules/core-js/internals/shared-key.js");
var hiddenKeys = __webpack_require__(/*! ../internals/hidden-keys */ "./node_modules/core-js/internals/hidden-keys.js");

var OBJECT_ALREADY_INITIALIZED = 'Object already initialized';
var TypeError = global.TypeError;
var WeakMap = global.WeakMap;
var set, get, has;

var enforce = function (it) {
  return has(it) ? get(it) : set(it, {});
};

var getterFor = function (TYPE) {
  return function (it) {
    var state;
    if (!isObject(it) || (state = get(it)).type !== TYPE) {
      throw new TypeError('Incompatible receiver, ' + TYPE + ' required');
    } return state;
  };
};

if (NATIVE_WEAK_MAP || shared.state) {
  var store = shared.state || (shared.state = new WeakMap());
  /* eslint-disable no-self-assign -- prototype methods protection */
  store.get = store.get;
  store.has = store.has;
  store.set = store.set;
  /* eslint-enable no-self-assign -- prototype methods protection */
  set = function (it, metadata) {
    if (store.has(it)) throw new TypeError(OBJECT_ALREADY_INITIALIZED);
    metadata.facade = it;
    store.set(it, metadata);
    return metadata;
  };
  get = function (it) {
    return store.get(it) || {};
  };
  has = function (it) {
    return store.has(it);
  };
} else {
  var STATE = sharedKey('state');
  hiddenKeys[STATE] = true;
  set = function (it, metadata) {
    if (hasOwn(it, STATE)) throw new TypeError(OBJECT_ALREADY_INITIALIZED);
    metadata.facade = it;
    createNonEnumerableProperty(it, STATE, metadata);
    return metadata;
  };
  get = function (it) {
    return hasOwn(it, STATE) ? it[STATE] : {};
  };
  has = function (it) {
    return hasOwn(it, STATE);
  };
}

module.exports = {
  set: set,
  get: get,
  has: has,
  enforce: enforce,
  getterFor: getterFor
};


/***/ }),

/***/ "./node_modules/core-js/internals/is-callable.js":
/*!*******************************************************!*\
  !*** ./node_modules/core-js/internals/is-callable.js ***!
  \*******************************************************/
/***/ (function(module) {


// https://tc39.es/ecma262/#sec-IsHTMLDDA-internal-slot
var documentAll = typeof document == 'object' && document.all;

// `IsCallable` abstract operation
// https://tc39.es/ecma262/#sec-iscallable
// eslint-disable-next-line unicorn/no-typeof-undefined -- required for testing
module.exports = typeof documentAll == 'undefined' && documentAll !== undefined ? function (argument) {
  return typeof argument == 'function' || argument === documentAll;
} : function (argument) {
  return typeof argument == 'function';
};


/***/ }),

/***/ "./node_modules/core-js/internals/is-forced.js":
/*!*****************************************************!*\
  !*** ./node_modules/core-js/internals/is-forced.js ***!
  \*****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");
var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");

var replacement = /#|\.prototype\./;

var isForced = function (feature, detection) {
  var value = data[normalize(feature)];
  return value === POLYFILL ? true
    : value === NATIVE ? false
    : isCallable(detection) ? fails(detection)
    : !!detection;
};

var normalize = isForced.normalize = function (string) {
  return String(string).replace(replacement, '.').toLowerCase();
};

var data = isForced.data = {};
var NATIVE = isForced.NATIVE = 'N';
var POLYFILL = isForced.POLYFILL = 'P';

module.exports = isForced;


/***/ }),

/***/ "./node_modules/core-js/internals/is-null-or-undefined.js":
/*!****************************************************************!*\
  !*** ./node_modules/core-js/internals/is-null-or-undefined.js ***!
  \****************************************************************/
/***/ (function(module) {


// we can't use just `it == null` since of `document.all` special case
// https://tc39.es/ecma262/#sec-IsHTMLDDA-internal-slot-aec
module.exports = function (it) {
  return it === null || it === undefined;
};


/***/ }),

/***/ "./node_modules/core-js/internals/is-object.js":
/*!*****************************************************!*\
  !*** ./node_modules/core-js/internals/is-object.js ***!
  \*****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");

module.exports = function (it) {
  return typeof it == 'object' ? it !== null : isCallable(it);
};


/***/ }),

/***/ "./node_modules/core-js/internals/is-pure.js":
/*!***************************************************!*\
  !*** ./node_modules/core-js/internals/is-pure.js ***!
  \***************************************************/
/***/ (function(module) {


module.exports = false;


/***/ }),

/***/ "./node_modules/core-js/internals/is-symbol.js":
/*!*****************************************************!*\
  !*** ./node_modules/core-js/internals/is-symbol.js ***!
  \*****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var getBuiltIn = __webpack_require__(/*! ../internals/get-built-in */ "./node_modules/core-js/internals/get-built-in.js");
var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");
var isPrototypeOf = __webpack_require__(/*! ../internals/object-is-prototype-of */ "./node_modules/core-js/internals/object-is-prototype-of.js");
var USE_SYMBOL_AS_UID = __webpack_require__(/*! ../internals/use-symbol-as-uid */ "./node_modules/core-js/internals/use-symbol-as-uid.js");

var $Object = Object;

module.exports = USE_SYMBOL_AS_UID ? function (it) {
  return typeof it == 'symbol';
} : function (it) {
  var $Symbol = getBuiltIn('Symbol');
  return isCallable($Symbol) && isPrototypeOf($Symbol.prototype, $Object(it));
};


/***/ }),

/***/ "./node_modules/core-js/internals/length-of-array-like.js":
/*!****************************************************************!*\
  !*** ./node_modules/core-js/internals/length-of-array-like.js ***!
  \****************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var toLength = __webpack_require__(/*! ../internals/to-length */ "./node_modules/core-js/internals/to-length.js");

// `LengthOfArrayLike` abstract operation
// https://tc39.es/ecma262/#sec-lengthofarraylike
module.exports = function (obj) {
  return toLength(obj.length);
};


/***/ }),

/***/ "./node_modules/core-js/internals/make-built-in.js":
/*!*********************************************************!*\
  !*** ./node_modules/core-js/internals/make-built-in.js ***!
  \*********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");
var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");
var hasOwn = __webpack_require__(/*! ../internals/has-own-property */ "./node_modules/core-js/internals/has-own-property.js");
var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var CONFIGURABLE_FUNCTION_NAME = (__webpack_require__(/*! ../internals/function-name */ "./node_modules/core-js/internals/function-name.js").CONFIGURABLE);
var inspectSource = __webpack_require__(/*! ../internals/inspect-source */ "./node_modules/core-js/internals/inspect-source.js");
var InternalStateModule = __webpack_require__(/*! ../internals/internal-state */ "./node_modules/core-js/internals/internal-state.js");

var enforceInternalState = InternalStateModule.enforce;
var getInternalState = InternalStateModule.get;
var $String = String;
// eslint-disable-next-line es/no-object-defineproperty -- safe
var defineProperty = Object.defineProperty;
var stringSlice = uncurryThis(''.slice);
var replace = uncurryThis(''.replace);
var join = uncurryThis([].join);

var CONFIGURABLE_LENGTH = DESCRIPTORS && !fails(function () {
  return defineProperty(function () { /* empty */ }, 'length', { value: 8 }).length !== 8;
});

var TEMPLATE = String(String).split('String');

var makeBuiltIn = module.exports = function (value, name, options) {
  if (stringSlice($String(name), 0, 7) === 'Symbol(') {
    name = '[' + replace($String(name), /^Symbol\(([^)]*)\).*$/, '$1') + ']';
  }
  if (options && options.getter) name = 'get ' + name;
  if (options && options.setter) name = 'set ' + name;
  if (!hasOwn(value, 'name') || (CONFIGURABLE_FUNCTION_NAME && value.name !== name)) {
    if (DESCRIPTORS) defineProperty(value, 'name', { value: name, configurable: true });
    else value.name = name;
  }
  if (CONFIGURABLE_LENGTH && options && hasOwn(options, 'arity') && value.length !== options.arity) {
    defineProperty(value, 'length', { value: options.arity });
  }
  try {
    if (options && hasOwn(options, 'constructor') && options.constructor) {
      if (DESCRIPTORS) defineProperty(value, 'prototype', { writable: false });
    // in V8 ~ Chrome 53, prototypes of some methods, like `Array.prototype.values`, are non-writable
    } else if (value.prototype) value.prototype = undefined;
  } catch (error) { /* empty */ }
  var state = enforceInternalState(value);
  if (!hasOwn(state, 'source')) {
    state.source = join(TEMPLATE, typeof name == 'string' ? name : '');
  } return value;
};

// add fake Function#toString for correct work wrapped methods / constructors with methods like LoDash isNative
// eslint-disable-next-line no-extend-native -- required
Function.prototype.toString = makeBuiltIn(function toString() {
  return isCallable(this) && getInternalState(this).source || inspectSource(this);
}, 'toString');


/***/ }),

/***/ "./node_modules/core-js/internals/math-trunc.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/internals/math-trunc.js ***!
  \******************************************************/
/***/ (function(module) {


var ceil = Math.ceil;
var floor = Math.floor;

// `Math.trunc` method
// https://tc39.es/ecma262/#sec-math.trunc
// eslint-disable-next-line es/no-math-trunc -- safe
module.exports = Math.trunc || function trunc(x) {
  var n = +x;
  return (n > 0 ? floor : ceil)(n);
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-assign.js":
/*!*********************************************************!*\
  !*** ./node_modules/core-js/internals/object-assign.js ***!
  \*********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var call = __webpack_require__(/*! ../internals/function-call */ "./node_modules/core-js/internals/function-call.js");
var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");
var objectKeys = __webpack_require__(/*! ../internals/object-keys */ "./node_modules/core-js/internals/object-keys.js");
var getOwnPropertySymbolsModule = __webpack_require__(/*! ../internals/object-get-own-property-symbols */ "./node_modules/core-js/internals/object-get-own-property-symbols.js");
var propertyIsEnumerableModule = __webpack_require__(/*! ../internals/object-property-is-enumerable */ "./node_modules/core-js/internals/object-property-is-enumerable.js");
var toObject = __webpack_require__(/*! ../internals/to-object */ "./node_modules/core-js/internals/to-object.js");
var IndexedObject = __webpack_require__(/*! ../internals/indexed-object */ "./node_modules/core-js/internals/indexed-object.js");

// eslint-disable-next-line es/no-object-assign -- safe
var $assign = Object.assign;
// eslint-disable-next-line es/no-object-defineproperty -- required for testing
var defineProperty = Object.defineProperty;
var concat = uncurryThis([].concat);

// `Object.assign` method
// https://tc39.es/ecma262/#sec-object.assign
module.exports = !$assign || fails(function () {
  // should have correct order of operations (Edge bug)
  if (DESCRIPTORS && $assign({ b: 1 }, $assign(defineProperty({}, 'a', {
    enumerable: true,
    get: function () {
      defineProperty(this, 'b', {
        value: 3,
        enumerable: false
      });
    }
  }), { b: 2 })).b !== 1) return true;
  // should work with symbols and should have deterministic property order (V8 bug)
  var A = {};
  var B = {};
  // eslint-disable-next-line es/no-symbol -- safe
  var symbol = Symbol('assign detection');
  var alphabet = 'abcdefghijklmnopqrst';
  A[symbol] = 7;
  alphabet.split('').forEach(function (chr) { B[chr] = chr; });
  return $assign({}, A)[symbol] !== 7 || objectKeys($assign({}, B)).join('') !== alphabet;
}) ? function assign(target, source) { // eslint-disable-line no-unused-vars -- required for `.length`
  var T = toObject(target);
  var argumentsLength = arguments.length;
  var index = 1;
  var getOwnPropertySymbols = getOwnPropertySymbolsModule.f;
  var propertyIsEnumerable = propertyIsEnumerableModule.f;
  while (argumentsLength > index) {
    var S = IndexedObject(arguments[index++]);
    var keys = getOwnPropertySymbols ? concat(objectKeys(S), getOwnPropertySymbols(S)) : objectKeys(S);
    var length = keys.length;
    var j = 0;
    var key;
    while (length > j) {
      key = keys[j++];
      if (!DESCRIPTORS || call(propertyIsEnumerable, S, key)) T[key] = S[key];
    }
  } return T;
} : $assign;


/***/ }),

/***/ "./node_modules/core-js/internals/object-create.js":
/*!*********************************************************!*\
  !*** ./node_modules/core-js/internals/object-create.js ***!
  \*********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


/* global ActiveXObject -- old IE, WSH */
var anObject = __webpack_require__(/*! ../internals/an-object */ "./node_modules/core-js/internals/an-object.js");
var definePropertiesModule = __webpack_require__(/*! ../internals/object-define-properties */ "./node_modules/core-js/internals/object-define-properties.js");
var enumBugKeys = __webpack_require__(/*! ../internals/enum-bug-keys */ "./node_modules/core-js/internals/enum-bug-keys.js");
var hiddenKeys = __webpack_require__(/*! ../internals/hidden-keys */ "./node_modules/core-js/internals/hidden-keys.js");
var html = __webpack_require__(/*! ../internals/html */ "./node_modules/core-js/internals/html.js");
var documentCreateElement = __webpack_require__(/*! ../internals/document-create-element */ "./node_modules/core-js/internals/document-create-element.js");
var sharedKey = __webpack_require__(/*! ../internals/shared-key */ "./node_modules/core-js/internals/shared-key.js");

var GT = '>';
var LT = '<';
var PROTOTYPE = 'prototype';
var SCRIPT = 'script';
var IE_PROTO = sharedKey('IE_PROTO');

var EmptyConstructor = function () { /* empty */ };

var scriptTag = function (content) {
  return LT + SCRIPT + GT + content + LT + '/' + SCRIPT + GT;
};

// Create object with fake `null` prototype: use ActiveX Object with cleared prototype
var NullProtoObjectViaActiveX = function (activeXDocument) {
  activeXDocument.write(scriptTag(''));
  activeXDocument.close();
  var temp = activeXDocument.parentWindow.Object;
  activeXDocument = null; // avoid memory leak
  return temp;
};

// Create object with fake `null` prototype: use iframe Object with cleared prototype
var NullProtoObjectViaIFrame = function () {
  // Thrash, waste and sodomy: IE GC bug
  var iframe = documentCreateElement('iframe');
  var JS = 'java' + SCRIPT + ':';
  var iframeDocument;
  iframe.style.display = 'none';
  html.appendChild(iframe);
  // https://github.com/zloirock/core-js/issues/475
  iframe.src = String(JS);
  iframeDocument = iframe.contentWindow.document;
  iframeDocument.open();
  iframeDocument.write(scriptTag('document.F=Object'));
  iframeDocument.close();
  return iframeDocument.F;
};

// Check for document.domain and active x support
// No need to use active x approach when document.domain is not set
// see https://github.com/es-shims/es5-shim/issues/150
// variation of https://github.com/kitcambridge/es5-shim/commit/4f738ac066346
// avoid IE GC bug
var activeXDocument;
var NullProtoObject = function () {
  try {
    activeXDocument = new ActiveXObject('htmlfile');
  } catch (error) { /* ignore */ }
  NullProtoObject = typeof document != 'undefined'
    ? document.domain && activeXDocument
      ? NullProtoObjectViaActiveX(activeXDocument) // old IE
      : NullProtoObjectViaIFrame()
    : NullProtoObjectViaActiveX(activeXDocument); // WSH
  var length = enumBugKeys.length;
  while (length--) delete NullProtoObject[PROTOTYPE][enumBugKeys[length]];
  return NullProtoObject();
};

hiddenKeys[IE_PROTO] = true;

// `Object.create` method
// https://tc39.es/ecma262/#sec-object.create
// eslint-disable-next-line es/no-object-create -- safe
module.exports = Object.create || function create(O, Properties) {
  var result;
  if (O !== null) {
    EmptyConstructor[PROTOTYPE] = anObject(O);
    result = new EmptyConstructor();
    EmptyConstructor[PROTOTYPE] = null;
    // add "__proto__" for Object.getPrototypeOf polyfill
    result[IE_PROTO] = O;
  } else result = NullProtoObject();
  return Properties === undefined ? result : definePropertiesModule.f(result, Properties);
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-define-properties.js":
/*!********************************************************************!*\
  !*** ./node_modules/core-js/internals/object-define-properties.js ***!
  \********************************************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var V8_PROTOTYPE_DEFINE_BUG = __webpack_require__(/*! ../internals/v8-prototype-define-bug */ "./node_modules/core-js/internals/v8-prototype-define-bug.js");
var definePropertyModule = __webpack_require__(/*! ../internals/object-define-property */ "./node_modules/core-js/internals/object-define-property.js");
var anObject = __webpack_require__(/*! ../internals/an-object */ "./node_modules/core-js/internals/an-object.js");
var toIndexedObject = __webpack_require__(/*! ../internals/to-indexed-object */ "./node_modules/core-js/internals/to-indexed-object.js");
var objectKeys = __webpack_require__(/*! ../internals/object-keys */ "./node_modules/core-js/internals/object-keys.js");

// `Object.defineProperties` method
// https://tc39.es/ecma262/#sec-object.defineproperties
// eslint-disable-next-line es/no-object-defineproperties -- safe
exports.f = DESCRIPTORS && !V8_PROTOTYPE_DEFINE_BUG ? Object.defineProperties : function defineProperties(O, Properties) {
  anObject(O);
  var props = toIndexedObject(Properties);
  var keys = objectKeys(Properties);
  var length = keys.length;
  var index = 0;
  var key;
  while (length > index) definePropertyModule.f(O, key = keys[index++], props[key]);
  return O;
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-define-property.js":
/*!******************************************************************!*\
  !*** ./node_modules/core-js/internals/object-define-property.js ***!
  \******************************************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var IE8_DOM_DEFINE = __webpack_require__(/*! ../internals/ie8-dom-define */ "./node_modules/core-js/internals/ie8-dom-define.js");
var V8_PROTOTYPE_DEFINE_BUG = __webpack_require__(/*! ../internals/v8-prototype-define-bug */ "./node_modules/core-js/internals/v8-prototype-define-bug.js");
var anObject = __webpack_require__(/*! ../internals/an-object */ "./node_modules/core-js/internals/an-object.js");
var toPropertyKey = __webpack_require__(/*! ../internals/to-property-key */ "./node_modules/core-js/internals/to-property-key.js");

var $TypeError = TypeError;
// eslint-disable-next-line es/no-object-defineproperty -- safe
var $defineProperty = Object.defineProperty;
// eslint-disable-next-line es/no-object-getownpropertydescriptor -- safe
var $getOwnPropertyDescriptor = Object.getOwnPropertyDescriptor;
var ENUMERABLE = 'enumerable';
var CONFIGURABLE = 'configurable';
var WRITABLE = 'writable';

// `Object.defineProperty` method
// https://tc39.es/ecma262/#sec-object.defineproperty
exports.f = DESCRIPTORS ? V8_PROTOTYPE_DEFINE_BUG ? function defineProperty(O, P, Attributes) {
  anObject(O);
  P = toPropertyKey(P);
  anObject(Attributes);
  if (typeof O === 'function' && P === 'prototype' && 'value' in Attributes && WRITABLE in Attributes && !Attributes[WRITABLE]) {
    var current = $getOwnPropertyDescriptor(O, P);
    if (current && current[WRITABLE]) {
      O[P] = Attributes.value;
      Attributes = {
        configurable: CONFIGURABLE in Attributes ? Attributes[CONFIGURABLE] : current[CONFIGURABLE],
        enumerable: ENUMERABLE in Attributes ? Attributes[ENUMERABLE] : current[ENUMERABLE],
        writable: false
      };
    }
  } return $defineProperty(O, P, Attributes);
} : $defineProperty : function defineProperty(O, P, Attributes) {
  anObject(O);
  P = toPropertyKey(P);
  anObject(Attributes);
  if (IE8_DOM_DEFINE) try {
    return $defineProperty(O, P, Attributes);
  } catch (error) { /* empty */ }
  if ('get' in Attributes || 'set' in Attributes) throw new $TypeError('Accessors not supported');
  if ('value' in Attributes) O[P] = Attributes.value;
  return O;
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-get-own-property-descriptor.js":
/*!******************************************************************************!*\
  !*** ./node_modules/core-js/internals/object-get-own-property-descriptor.js ***!
  \******************************************************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var call = __webpack_require__(/*! ../internals/function-call */ "./node_modules/core-js/internals/function-call.js");
var propertyIsEnumerableModule = __webpack_require__(/*! ../internals/object-property-is-enumerable */ "./node_modules/core-js/internals/object-property-is-enumerable.js");
var createPropertyDescriptor = __webpack_require__(/*! ../internals/create-property-descriptor */ "./node_modules/core-js/internals/create-property-descriptor.js");
var toIndexedObject = __webpack_require__(/*! ../internals/to-indexed-object */ "./node_modules/core-js/internals/to-indexed-object.js");
var toPropertyKey = __webpack_require__(/*! ../internals/to-property-key */ "./node_modules/core-js/internals/to-property-key.js");
var hasOwn = __webpack_require__(/*! ../internals/has-own-property */ "./node_modules/core-js/internals/has-own-property.js");
var IE8_DOM_DEFINE = __webpack_require__(/*! ../internals/ie8-dom-define */ "./node_modules/core-js/internals/ie8-dom-define.js");

// eslint-disable-next-line es/no-object-getownpropertydescriptor -- safe
var $getOwnPropertyDescriptor = Object.getOwnPropertyDescriptor;

// `Object.getOwnPropertyDescriptor` method
// https://tc39.es/ecma262/#sec-object.getownpropertydescriptor
exports.f = DESCRIPTORS ? $getOwnPropertyDescriptor : function getOwnPropertyDescriptor(O, P) {
  O = toIndexedObject(O);
  P = toPropertyKey(P);
  if (IE8_DOM_DEFINE) try {
    return $getOwnPropertyDescriptor(O, P);
  } catch (error) { /* empty */ }
  if (hasOwn(O, P)) return createPropertyDescriptor(!call(propertyIsEnumerableModule.f, O, P), O[P]);
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-get-own-property-names.js":
/*!*************************************************************************!*\
  !*** ./node_modules/core-js/internals/object-get-own-property-names.js ***!
  \*************************************************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {


var internalObjectKeys = __webpack_require__(/*! ../internals/object-keys-internal */ "./node_modules/core-js/internals/object-keys-internal.js");
var enumBugKeys = __webpack_require__(/*! ../internals/enum-bug-keys */ "./node_modules/core-js/internals/enum-bug-keys.js");

var hiddenKeys = enumBugKeys.concat('length', 'prototype');

// `Object.getOwnPropertyNames` method
// https://tc39.es/ecma262/#sec-object.getownpropertynames
// eslint-disable-next-line es/no-object-getownpropertynames -- safe
exports.f = Object.getOwnPropertyNames || function getOwnPropertyNames(O) {
  return internalObjectKeys(O, hiddenKeys);
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-get-own-property-symbols.js":
/*!***************************************************************************!*\
  !*** ./node_modules/core-js/internals/object-get-own-property-symbols.js ***!
  \***************************************************************************/
/***/ (function(__unused_webpack_module, exports) {


// eslint-disable-next-line es/no-object-getownpropertysymbols -- safe
exports.f = Object.getOwnPropertySymbols;


/***/ }),

/***/ "./node_modules/core-js/internals/object-is-prototype-of.js":
/*!******************************************************************!*\
  !*** ./node_modules/core-js/internals/object-is-prototype-of.js ***!
  \******************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");

module.exports = uncurryThis({}.isPrototypeOf);


/***/ }),

/***/ "./node_modules/core-js/internals/object-keys-internal.js":
/*!****************************************************************!*\
  !*** ./node_modules/core-js/internals/object-keys-internal.js ***!
  \****************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var hasOwn = __webpack_require__(/*! ../internals/has-own-property */ "./node_modules/core-js/internals/has-own-property.js");
var toIndexedObject = __webpack_require__(/*! ../internals/to-indexed-object */ "./node_modules/core-js/internals/to-indexed-object.js");
var indexOf = (__webpack_require__(/*! ../internals/array-includes */ "./node_modules/core-js/internals/array-includes.js").indexOf);
var hiddenKeys = __webpack_require__(/*! ../internals/hidden-keys */ "./node_modules/core-js/internals/hidden-keys.js");

var push = uncurryThis([].push);

module.exports = function (object, names) {
  var O = toIndexedObject(object);
  var i = 0;
  var result = [];
  var key;
  for (key in O) !hasOwn(hiddenKeys, key) && hasOwn(O, key) && push(result, key);
  // Don't enum bug & hidden keys
  while (names.length > i) if (hasOwn(O, key = names[i++])) {
    ~indexOf(result, key) || push(result, key);
  }
  return result;
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-keys.js":
/*!*******************************************************!*\
  !*** ./node_modules/core-js/internals/object-keys.js ***!
  \*******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var internalObjectKeys = __webpack_require__(/*! ../internals/object-keys-internal */ "./node_modules/core-js/internals/object-keys-internal.js");
var enumBugKeys = __webpack_require__(/*! ../internals/enum-bug-keys */ "./node_modules/core-js/internals/enum-bug-keys.js");

// `Object.keys` method
// https://tc39.es/ecma262/#sec-object.keys
// eslint-disable-next-line es/no-object-keys -- safe
module.exports = Object.keys || function keys(O) {
  return internalObjectKeys(O, enumBugKeys);
};


/***/ }),

/***/ "./node_modules/core-js/internals/object-property-is-enumerable.js":
/*!*************************************************************************!*\
  !*** ./node_modules/core-js/internals/object-property-is-enumerable.js ***!
  \*************************************************************************/
/***/ (function(__unused_webpack_module, exports) {


var $propertyIsEnumerable = {}.propertyIsEnumerable;
// eslint-disable-next-line es/no-object-getownpropertydescriptor -- safe
var getOwnPropertyDescriptor = Object.getOwnPropertyDescriptor;

// Nashorn ~ JDK8 bug
var NASHORN_BUG = getOwnPropertyDescriptor && !$propertyIsEnumerable.call({ 1: 2 }, 1);

// `Object.prototype.propertyIsEnumerable` method implementation
// https://tc39.es/ecma262/#sec-object.prototype.propertyisenumerable
exports.f = NASHORN_BUG ? function propertyIsEnumerable(V) {
  var descriptor = getOwnPropertyDescriptor(this, V);
  return !!descriptor && descriptor.enumerable;
} : $propertyIsEnumerable;


/***/ }),

/***/ "./node_modules/core-js/internals/ordinary-to-primitive.js":
/*!*****************************************************************!*\
  !*** ./node_modules/core-js/internals/ordinary-to-primitive.js ***!
  \*****************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var call = __webpack_require__(/*! ../internals/function-call */ "./node_modules/core-js/internals/function-call.js");
var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");
var isObject = __webpack_require__(/*! ../internals/is-object */ "./node_modules/core-js/internals/is-object.js");

var $TypeError = TypeError;

// `OrdinaryToPrimitive` abstract operation
// https://tc39.es/ecma262/#sec-ordinarytoprimitive
module.exports = function (input, pref) {
  var fn, val;
  if (pref === 'string' && isCallable(fn = input.toString) && !isObject(val = call(fn, input))) return val;
  if (isCallable(fn = input.valueOf) && !isObject(val = call(fn, input))) return val;
  if (pref !== 'string' && isCallable(fn = input.toString) && !isObject(val = call(fn, input))) return val;
  throw new $TypeError("Can't convert object to primitive value");
};


/***/ }),

/***/ "./node_modules/core-js/internals/own-keys.js":
/*!****************************************************!*\
  !*** ./node_modules/core-js/internals/own-keys.js ***!
  \****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var getBuiltIn = __webpack_require__(/*! ../internals/get-built-in */ "./node_modules/core-js/internals/get-built-in.js");
var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var getOwnPropertyNamesModule = __webpack_require__(/*! ../internals/object-get-own-property-names */ "./node_modules/core-js/internals/object-get-own-property-names.js");
var getOwnPropertySymbolsModule = __webpack_require__(/*! ../internals/object-get-own-property-symbols */ "./node_modules/core-js/internals/object-get-own-property-symbols.js");
var anObject = __webpack_require__(/*! ../internals/an-object */ "./node_modules/core-js/internals/an-object.js");

var concat = uncurryThis([].concat);

// all object keys, includes non-enumerable and symbols
module.exports = getBuiltIn('Reflect', 'ownKeys') || function ownKeys(it) {
  var keys = getOwnPropertyNamesModule.f(anObject(it));
  var getOwnPropertySymbols = getOwnPropertySymbolsModule.f;
  return getOwnPropertySymbols ? concat(keys, getOwnPropertySymbols(it)) : keys;
};


/***/ }),

/***/ "./node_modules/core-js/internals/path.js":
/*!************************************************!*\
  !*** ./node_modules/core-js/internals/path.js ***!
  \************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");

module.exports = global;


/***/ }),

/***/ "./node_modules/core-js/internals/require-object-coercible.js":
/*!********************************************************************!*\
  !*** ./node_modules/core-js/internals/require-object-coercible.js ***!
  \********************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var isNullOrUndefined = __webpack_require__(/*! ../internals/is-null-or-undefined */ "./node_modules/core-js/internals/is-null-or-undefined.js");

var $TypeError = TypeError;

// `RequireObjectCoercible` abstract operation
// https://tc39.es/ecma262/#sec-requireobjectcoercible
module.exports = function (it) {
  if (isNullOrUndefined(it)) throw new $TypeError("Can't call method on " + it);
  return it;
};


/***/ }),

/***/ "./node_modules/core-js/internals/shared-key.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/internals/shared-key.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var shared = __webpack_require__(/*! ../internals/shared */ "./node_modules/core-js/internals/shared.js");
var uid = __webpack_require__(/*! ../internals/uid */ "./node_modules/core-js/internals/uid.js");

var keys = shared('keys');

module.exports = function (key) {
  return keys[key] || (keys[key] = uid(key));
};


/***/ }),

/***/ "./node_modules/core-js/internals/shared-store.js":
/*!********************************************************!*\
  !*** ./node_modules/core-js/internals/shared-store.js ***!
  \********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var IS_PURE = __webpack_require__(/*! ../internals/is-pure */ "./node_modules/core-js/internals/is-pure.js");
var globalThis = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var defineGlobalProperty = __webpack_require__(/*! ../internals/define-global-property */ "./node_modules/core-js/internals/define-global-property.js");

var SHARED = '__core-js_shared__';
var store = module.exports = globalThis[SHARED] || defineGlobalProperty(SHARED, {});

(store.versions || (store.versions = [])).push({
  version: '3.37.1',
  mode: IS_PURE ? 'pure' : 'global',
  copyright: '© 2014-2024 Denis Pushkarev (zloirock.ru)',
  license: 'https://github.com/zloirock/core-js/blob/v3.37.1/LICENSE',
  source: 'https://github.com/zloirock/core-js'
});


/***/ }),

/***/ "./node_modules/core-js/internals/shared.js":
/*!**************************************************!*\
  !*** ./node_modules/core-js/internals/shared.js ***!
  \**************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var store = __webpack_require__(/*! ../internals/shared-store */ "./node_modules/core-js/internals/shared-store.js");

module.exports = function (key, value) {
  return store[key] || (store[key] = value || {});
};


/***/ }),

/***/ "./node_modules/core-js/internals/symbol-constructor-detection.js":
/*!************************************************************************!*\
  !*** ./node_modules/core-js/internals/symbol-constructor-detection.js ***!
  \************************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


/* eslint-disable es/no-symbol -- required for testing */
var V8_VERSION = __webpack_require__(/*! ../internals/engine-v8-version */ "./node_modules/core-js/internals/engine-v8-version.js");
var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");
var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");

var $String = global.String;

// eslint-disable-next-line es/no-object-getownpropertysymbols -- required for testing
module.exports = !!Object.getOwnPropertySymbols && !fails(function () {
  var symbol = Symbol('symbol detection');
  // Chrome 38 Symbol has incorrect toString conversion
  // `get-own-property-symbols` polyfill symbols converted to object are not Symbol instances
  // nb: Do not call `String` directly to avoid this being optimized out to `symbol+''` which will,
  // of course, fail.
  return !$String(symbol) || !(Object(symbol) instanceof Symbol) ||
    // Chrome 38-40 symbols are not inherited from DOM collections prototypes to instances
    !Symbol.sham && V8_VERSION && V8_VERSION < 41;
});


/***/ }),

/***/ "./node_modules/core-js/internals/to-absolute-index.js":
/*!*************************************************************!*\
  !*** ./node_modules/core-js/internals/to-absolute-index.js ***!
  \*************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var toIntegerOrInfinity = __webpack_require__(/*! ../internals/to-integer-or-infinity */ "./node_modules/core-js/internals/to-integer-or-infinity.js");

var max = Math.max;
var min = Math.min;

// Helper for a popular repeating case of the spec:
// Let integer be ? ToInteger(index).
// If integer < 0, let result be max((length + integer), 0); else let result be min(integer, length).
module.exports = function (index, length) {
  var integer = toIntegerOrInfinity(index);
  return integer < 0 ? max(integer + length, 0) : min(integer, length);
};


/***/ }),

/***/ "./node_modules/core-js/internals/to-indexed-object.js":
/*!*************************************************************!*\
  !*** ./node_modules/core-js/internals/to-indexed-object.js ***!
  \*************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


// toObject with fallback for non-array-like ES3 strings
var IndexedObject = __webpack_require__(/*! ../internals/indexed-object */ "./node_modules/core-js/internals/indexed-object.js");
var requireObjectCoercible = __webpack_require__(/*! ../internals/require-object-coercible */ "./node_modules/core-js/internals/require-object-coercible.js");

module.exports = function (it) {
  return IndexedObject(requireObjectCoercible(it));
};


/***/ }),

/***/ "./node_modules/core-js/internals/to-integer-or-infinity.js":
/*!******************************************************************!*\
  !*** ./node_modules/core-js/internals/to-integer-or-infinity.js ***!
  \******************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var trunc = __webpack_require__(/*! ../internals/math-trunc */ "./node_modules/core-js/internals/math-trunc.js");

// `ToIntegerOrInfinity` abstract operation
// https://tc39.es/ecma262/#sec-tointegerorinfinity
module.exports = function (argument) {
  var number = +argument;
  // eslint-disable-next-line no-self-compare -- NaN check
  return number !== number || number === 0 ? 0 : trunc(number);
};


/***/ }),

/***/ "./node_modules/core-js/internals/to-length.js":
/*!*****************************************************!*\
  !*** ./node_modules/core-js/internals/to-length.js ***!
  \*****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var toIntegerOrInfinity = __webpack_require__(/*! ../internals/to-integer-or-infinity */ "./node_modules/core-js/internals/to-integer-or-infinity.js");

var min = Math.min;

// `ToLength` abstract operation
// https://tc39.es/ecma262/#sec-tolength
module.exports = function (argument) {
  var len = toIntegerOrInfinity(argument);
  return len > 0 ? min(len, 0x1FFFFFFFFFFFFF) : 0; // 2 ** 53 - 1 == 9007199254740991
};


/***/ }),

/***/ "./node_modules/core-js/internals/to-object.js":
/*!*****************************************************!*\
  !*** ./node_modules/core-js/internals/to-object.js ***!
  \*****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var requireObjectCoercible = __webpack_require__(/*! ../internals/require-object-coercible */ "./node_modules/core-js/internals/require-object-coercible.js");

var $Object = Object;

// `ToObject` abstract operation
// https://tc39.es/ecma262/#sec-toobject
module.exports = function (argument) {
  return $Object(requireObjectCoercible(argument));
};


/***/ }),

/***/ "./node_modules/core-js/internals/to-primitive.js":
/*!********************************************************!*\
  !*** ./node_modules/core-js/internals/to-primitive.js ***!
  \********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var call = __webpack_require__(/*! ../internals/function-call */ "./node_modules/core-js/internals/function-call.js");
var isObject = __webpack_require__(/*! ../internals/is-object */ "./node_modules/core-js/internals/is-object.js");
var isSymbol = __webpack_require__(/*! ../internals/is-symbol */ "./node_modules/core-js/internals/is-symbol.js");
var getMethod = __webpack_require__(/*! ../internals/get-method */ "./node_modules/core-js/internals/get-method.js");
var ordinaryToPrimitive = __webpack_require__(/*! ../internals/ordinary-to-primitive */ "./node_modules/core-js/internals/ordinary-to-primitive.js");
var wellKnownSymbol = __webpack_require__(/*! ../internals/well-known-symbol */ "./node_modules/core-js/internals/well-known-symbol.js");

var $TypeError = TypeError;
var TO_PRIMITIVE = wellKnownSymbol('toPrimitive');

// `ToPrimitive` abstract operation
// https://tc39.es/ecma262/#sec-toprimitive
module.exports = function (input, pref) {
  if (!isObject(input) || isSymbol(input)) return input;
  var exoticToPrim = getMethod(input, TO_PRIMITIVE);
  var result;
  if (exoticToPrim) {
    if (pref === undefined) pref = 'default';
    result = call(exoticToPrim, input, pref);
    if (!isObject(result) || isSymbol(result)) return result;
    throw new $TypeError("Can't convert object to primitive value");
  }
  if (pref === undefined) pref = 'number';
  return ordinaryToPrimitive(input, pref);
};


/***/ }),

/***/ "./node_modules/core-js/internals/to-property-key.js":
/*!***********************************************************!*\
  !*** ./node_modules/core-js/internals/to-property-key.js ***!
  \***********************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var toPrimitive = __webpack_require__(/*! ../internals/to-primitive */ "./node_modules/core-js/internals/to-primitive.js");
var isSymbol = __webpack_require__(/*! ../internals/is-symbol */ "./node_modules/core-js/internals/is-symbol.js");

// `ToPropertyKey` abstract operation
// https://tc39.es/ecma262/#sec-topropertykey
module.exports = function (argument) {
  var key = toPrimitive(argument, 'string');
  return isSymbol(key) ? key : key + '';
};


/***/ }),

/***/ "./node_modules/core-js/internals/try-to-string.js":
/*!*********************************************************!*\
  !*** ./node_modules/core-js/internals/try-to-string.js ***!
  \*********************************************************/
/***/ (function(module) {


var $String = String;

module.exports = function (argument) {
  try {
    return $String(argument);
  } catch (error) {
    return 'Object';
  }
};


/***/ }),

/***/ "./node_modules/core-js/internals/uid.js":
/*!***********************************************!*\
  !*** ./node_modules/core-js/internals/uid.js ***!
  \***********************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");

var id = 0;
var postfix = Math.random();
var toString = uncurryThis(1.0.toString);

module.exports = function (key) {
  return 'Symbol(' + (key === undefined ? '' : key) + ')_' + toString(++id + postfix, 36);
};


/***/ }),

/***/ "./node_modules/core-js/internals/use-symbol-as-uid.js":
/*!*************************************************************!*\
  !*** ./node_modules/core-js/internals/use-symbol-as-uid.js ***!
  \*************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


/* eslint-disable es/no-symbol -- required for testing */
var NATIVE_SYMBOL = __webpack_require__(/*! ../internals/symbol-constructor-detection */ "./node_modules/core-js/internals/symbol-constructor-detection.js");

module.exports = NATIVE_SYMBOL
  && !Symbol.sham
  && typeof Symbol.iterator == 'symbol';


/***/ }),

/***/ "./node_modules/core-js/internals/v8-prototype-define-bug.js":
/*!*******************************************************************!*\
  !*** ./node_modules/core-js/internals/v8-prototype-define-bug.js ***!
  \*******************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");

// V8 ~ Chrome 36-
// https://bugs.chromium.org/p/v8/issues/detail?id=3334
module.exports = DESCRIPTORS && fails(function () {
  // eslint-disable-next-line es/no-object-defineproperty -- required for testing
  return Object.defineProperty(function () { /* empty */ }, 'prototype', {
    value: 42,
    writable: false
  }).prototype !== 42;
});


/***/ }),

/***/ "./node_modules/core-js/internals/weak-map-basic-detection.js":
/*!********************************************************************!*\
  !*** ./node_modules/core-js/internals/weak-map-basic-detection.js ***!
  \********************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var isCallable = __webpack_require__(/*! ../internals/is-callable */ "./node_modules/core-js/internals/is-callable.js");

var WeakMap = global.WeakMap;

module.exports = isCallable(WeakMap) && /native code/.test(String(WeakMap));


/***/ }),

/***/ "./node_modules/core-js/internals/well-known-symbol.js":
/*!*************************************************************!*\
  !*** ./node_modules/core-js/internals/well-known-symbol.js ***!
  \*************************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");
var shared = __webpack_require__(/*! ../internals/shared */ "./node_modules/core-js/internals/shared.js");
var hasOwn = __webpack_require__(/*! ../internals/has-own-property */ "./node_modules/core-js/internals/has-own-property.js");
var uid = __webpack_require__(/*! ../internals/uid */ "./node_modules/core-js/internals/uid.js");
var NATIVE_SYMBOL = __webpack_require__(/*! ../internals/symbol-constructor-detection */ "./node_modules/core-js/internals/symbol-constructor-detection.js");
var USE_SYMBOL_AS_UID = __webpack_require__(/*! ../internals/use-symbol-as-uid */ "./node_modules/core-js/internals/use-symbol-as-uid.js");

var Symbol = global.Symbol;
var WellKnownSymbolsStore = shared('wks');
var createWellKnownSymbol = USE_SYMBOL_AS_UID ? Symbol['for'] || Symbol : Symbol && Symbol.withoutSetter || uid;

module.exports = function (name) {
  if (!hasOwn(WellKnownSymbolsStore, name)) {
    WellKnownSymbolsStore[name] = NATIVE_SYMBOL && hasOwn(Symbol, name)
      ? Symbol[name]
      : createWellKnownSymbol('Symbol.' + name);
  } return WellKnownSymbolsStore[name];
};


/***/ }),

/***/ "./node_modules/core-js/modules/es.array.includes.js":
/*!***********************************************************!*\
  !*** ./node_modules/core-js/modules/es.array.includes.js ***!
  \***********************************************************/
/***/ (function(__unused_webpack_module, __unused_webpack_exports, __webpack_require__) {


var $ = __webpack_require__(/*! ../internals/export */ "./node_modules/core-js/internals/export.js");
var $includes = (__webpack_require__(/*! ../internals/array-includes */ "./node_modules/core-js/internals/array-includes.js").includes);
var fails = __webpack_require__(/*! ../internals/fails */ "./node_modules/core-js/internals/fails.js");
var addToUnscopables = __webpack_require__(/*! ../internals/add-to-unscopables */ "./node_modules/core-js/internals/add-to-unscopables.js");

// FF99+ bug
var BROKEN_ON_SPARSE = fails(function () {
  // eslint-disable-next-line es/no-array-prototype-includes -- detection
  return !Array(1).includes();
});

// `Array.prototype.includes` method
// https://tc39.es/ecma262/#sec-array.prototype.includes
$({ target: 'Array', proto: true, forced: BROKEN_ON_SPARSE }, {
  includes: function includes(el /* , fromIndex = 0 */) {
    return $includes(this, el, arguments.length > 1 ? arguments[1] : undefined);
  }
});

// https://tc39.es/ecma262/#sec-array.prototype-@@unscopables
addToUnscopables('includes');


/***/ }),

/***/ "./node_modules/core-js/modules/es.function.name.js":
/*!**********************************************************!*\
  !*** ./node_modules/core-js/modules/es.function.name.js ***!
  \**********************************************************/
/***/ (function(__unused_webpack_module, __unused_webpack_exports, __webpack_require__) {


var DESCRIPTORS = __webpack_require__(/*! ../internals/descriptors */ "./node_modules/core-js/internals/descriptors.js");
var FUNCTION_NAME_EXISTS = (__webpack_require__(/*! ../internals/function-name */ "./node_modules/core-js/internals/function-name.js").EXISTS);
var uncurryThis = __webpack_require__(/*! ../internals/function-uncurry-this */ "./node_modules/core-js/internals/function-uncurry-this.js");
var defineBuiltInAccessor = __webpack_require__(/*! ../internals/define-built-in-accessor */ "./node_modules/core-js/internals/define-built-in-accessor.js");

var FunctionPrototype = Function.prototype;
var functionToString = uncurryThis(FunctionPrototype.toString);
var nameRE = /function\b(?:\s|\/\*[\S\s]*?\*\/|\/\/[^\n\r]*[\n\r]+)*([^\s(/]*)/;
var regExpExec = uncurryThis(nameRE.exec);
var NAME = 'name';

// Function instances `.name` property
// https://tc39.es/ecma262/#sec-function-instances-name
if (DESCRIPTORS && !FUNCTION_NAME_EXISTS) {
  defineBuiltInAccessor(FunctionPrototype, NAME, {
    configurable: true,
    get: function () {
      try {
        return regExpExec(nameRE, functionToString(this))[1];
      } catch (error) {
        return '';
      }
    }
  });
}


/***/ }),

/***/ "./node_modules/core-js/modules/es.global-this.js":
/*!********************************************************!*\
  !*** ./node_modules/core-js/modules/es.global-this.js ***!
  \********************************************************/
/***/ (function(__unused_webpack_module, __unused_webpack_exports, __webpack_require__) {


var $ = __webpack_require__(/*! ../internals/export */ "./node_modules/core-js/internals/export.js");
var global = __webpack_require__(/*! ../internals/global */ "./node_modules/core-js/internals/global.js");

// `globalThis` object
// https://tc39.es/ecma262/#sec-globalthis
$({ global: true, forced: global.globalThis !== global }, {
  globalThis: global
});


/***/ }),

/***/ "./node_modules/core-js/modules/es.object.assign.js":
/*!**********************************************************!*\
  !*** ./node_modules/core-js/modules/es.object.assign.js ***!
  \**********************************************************/
/***/ (function(__unused_webpack_module, __unused_webpack_exports, __webpack_require__) {


var $ = __webpack_require__(/*! ../internals/export */ "./node_modules/core-js/internals/export.js");
var assign = __webpack_require__(/*! ../internals/object-assign */ "./node_modules/core-js/internals/object-assign.js");

// `Object.assign` method
// https://tc39.es/ecma262/#sec-object.assign
// eslint-disable-next-line es/no-object-assign -- required for testing
$({ target: 'Object', stat: true, arity: 2, forced: Object.assign !== assign }, {
  assign: assign
});


/***/ }),

/***/ "./node_modules/core-js/stable/array/includes.js":
/*!*******************************************************!*\
  !*** ./node_modules/core-js/stable/array/includes.js ***!
  \*******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../../es/array/includes */ "./node_modules/core-js/es/array/includes.js");

module.exports = parent;


/***/ }),

/***/ "./node_modules/core-js/stable/function/name.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/stable/function/name.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../../es/function/name */ "./node_modules/core-js/es/function/name.js");

module.exports = parent;


/***/ }),

/***/ "./node_modules/core-js/stable/global-this.js":
/*!****************************************************!*\
  !*** ./node_modules/core-js/stable/global-this.js ***!
  \****************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../es/global-this */ "./node_modules/core-js/es/global-this.js");

module.exports = parent;


/***/ }),

/***/ "./node_modules/core-js/stable/object/assign.js":
/*!******************************************************!*\
  !*** ./node_modules/core-js/stable/object/assign.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {


var parent = __webpack_require__(/*! ../../es/object/assign */ "./node_modules/core-js/es/object/assign.js");

module.exports = parent;


/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	!function() {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = function(module) {
/******/ 			var getter = module && module.__esModule ?
/******/ 				function() { return module['default']; } :
/******/ 				function() { return module; };
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
/*!************************!*\
  !*** ./src/xlwings.ts ***!
  \************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getAccessToken: function() { return /* reexport safe */ _auth__WEBPACK_IMPORTED_MODULE_5__.getAccessToken; },
/* harmony export */   getActiveBookName: function() { return /* reexport safe */ _utils__WEBPACK_IMPORTED_MODULE_6__.getActiveBookName; },
/* harmony export */   init: function() { return /* binding */ init; },
/* harmony export */   registerCallback: function() { return /* binding */ registerCallback; },
/* harmony export */   runPython: function() { return /* binding */ runPython; }
/* harmony export */ });
/* harmony import */ var core_js_actual_object_assign__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! core-js/actual/object/assign */ "./node_modules/core-js/actual/object/assign.js");
/* harmony import */ var core_js_actual_object_assign__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(core_js_actual_object_assign__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var core_js_actual_array_includes__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! core-js/actual/array/includes */ "./node_modules/core-js/actual/array/includes.js");
/* harmony import */ var core_js_actual_array_includes__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(core_js_actual_array_includes__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var core_js_actual_global_this__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! core-js/actual/global-this */ "./node_modules/core-js/actual/global-this.js");
/* harmony import */ var core_js_actual_global_this__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(core_js_actual_global_this__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var core_js_actual_function_name__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! core-js/actual/function/name */ "./node_modules/core-js/actual/function/name.js");
/* harmony import */ var core_js_actual_function_name__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(core_js_actual_function_name__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _alert__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./alert */ "./src/alert.ts");
/* harmony import */ var _auth__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./auth */ "./src/auth.ts");
/* harmony import */ var _utils__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./utils */ "./src/utils.ts");
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArray = (undefined && undefined.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
// core-js polyfills for ie11









// Hook up buttons with the click event upon loading xlwings.js
document.addEventListener("DOMContentLoaded", init);
function init() {
    var _this = this;
    var appPathElement = document.getElementById("app-path");
    var appPath = appPathElement
        ? JSON.parse(appPathElement.textContent)
        : null;
    var elements = document.querySelectorAll("[xw-click]");
    elements.forEach(function (element) {
        element.addEventListener("click", function (event) { return __awaiter(_this, void 0, void 0, function () {
            var token, _a, config;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!(typeof globalThis.getAuth === "function")) return [3 /*break*/, 2];
                        return [4 /*yield*/, globalThis.getAuth()];
                    case 1:
                        _a = _b.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        _a = "";
                        _b.label = 3;
                    case 3:
                        token = _a;
                        config = element.getAttribute("xw-config")
                            ? JSON.parse(element.getAttribute("xw-config"))
                            : {};
                        runPython(window.location.origin +
                            (appPath && appPath.appPath !== "" ? "/".concat(appPath.appPath) : "") +
                            "/xlwings/custom-scripts-call/" +
                            element.getAttribute("xw-click"), __assign(__assign({}, config), { auth: token }));
                        return [2 /*return*/];
                }
            });
        }); });
    });
}
var version = "0.31.7";
globalThis.callbacks = {};
function runPython() {
    return __awaiter(this, arguments, void 0, function (url, _a) {
        var error_1;
        var _this = this;
        if (url === void 0) { url = ""; }
        var _b = _a === void 0 ? {} : _a, _c = _b.auth, auth = _c === void 0 ? "" : _c, _d = _b.include, include = _d === void 0 ? "" : _d, _e = _b.exclude, exclude = _e === void 0 ? "" : _e, _f = _b.headers, headers = _f === void 0 ? {} : _f;
        return __generator(this, function (_g) {
            switch (_g.label) {
                case 0: return [4 /*yield*/, Office.onReady()];
                case 1:
                    _g.sent();
                    _g.label = 2;
                case 2:
                    _g.trys.push([2, 4, , 6]);
                    return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                            var workbook, worksheets, sheets, configSheet, config, configRange, configValues, includeArray, excludeArray, property, payload, activeSheet, selection, names, namedItems, names2, sheetsLoader, namesSheetScope, namesSheetsScope2, _loop_1, _i, sheetsLoader_1, item, response, rawData, forceSync, _loop_2, _a, _b, action;
                            return __generator(this, function (_c) {
                                switch (_c.label) {
                                    case 0:
                                        workbook = context.workbook;
                                        workbook.load("name");
                                        worksheets = workbook.worksheets;
                                        worksheets.load("items/name");
                                        return [4 /*yield*/, context.sync()];
                                    case 1:
                                        _c.sent();
                                        sheets = worksheets.items;
                                        configSheet = worksheets.getItemOrNullObject("xlwings.conf");
                                        return [4 /*yield*/, context.sync()];
                                    case 2:
                                        _c.sent();
                                        config = {};
                                        if (!!configSheet.isNullObject) return [3 /*break*/, 4];
                                        configRange = configSheet
                                            .getRange("A1")
                                            .getSurroundingRegion()
                                            .load("values");
                                        return [4 /*yield*/, context.sync()];
                                    case 3:
                                        _c.sent();
                                        configValues = configRange.values;
                                        configValues.forEach(function (el) { return (config[el[0].toString()] = el[1].toString()); });
                                        _c.label = 4;
                                    case 4:
                                        if (auth === "") {
                                            auth = config["AUTH"] || "";
                                        }
                                        if (include === "") {
                                            include = config["INCLUDE"] || "";
                                        }
                                        includeArray = [];
                                        if (include !== "") {
                                            includeArray = include.split(",").map(function (item) { return item.trim(); });
                                        }
                                        if (exclude === "") {
                                            exclude = config["EXCLUDE"] || "";
                                        }
                                        excludeArray = [];
                                        if (exclude !== "") {
                                            excludeArray = exclude.split(",").map(function (item) { return item.trim(); });
                                        }
                                        if (includeArray.length > 0 && excludeArray.length > 0) {
                                            throw "Either use 'include' or 'exclude', but not both!";
                                        }
                                        if (includeArray.length > 0) {
                                            sheets.forEach(function (sheet) {
                                                if (!includeArray.includes(sheet.name)) {
                                                    excludeArray.push(sheet.name);
                                                }
                                            });
                                        }
                                        if (Object.keys(headers).length === 0) {
                                            for (property in config) {
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
                                        payload = {};
                                        payload["client"] = "Office.js";
                                        payload["version"] = version;
                                        activeSheet = worksheets.getActiveWorksheet().load("position");
                                        selection = workbook.getSelectedRange().load("address");
                                        return [4 /*yield*/, context.sync()];
                                    case 5:
                                        _c.sent();
                                        payload["book"] = {
                                            name: workbook.name,
                                            active_sheet_index: activeSheet.position,
                                            selection: selection.address.split("!").pop(),
                                        };
                                        names = [];
                                        namedItems = context.workbook.names.load("name, type");
                                        return [4 /*yield*/, context.sync()];
                                    case 6:
                                        _c.sent();
                                        namedItems.items.forEach(function (namedItem, ix) {
                                            // Currently filtering to named ranges
                                            if (namedItem.type === "Range") {
                                                names.push({
                                                    name: namedItem.name,
                                                    sheet: namedItem.getRange().worksheet.load("position"),
                                                    range: namedItem.getRange().load("address"),
                                                    scope_sheet_name: null,
                                                    scope_sheet_index: null,
                                                    book_scope: true, // workbook.names contains only workbook scope!
                                                });
                                            }
                                        });
                                        return [4 /*yield*/, context.sync()];
                                    case 7:
                                        _c.sent();
                                        names2 = [];
                                        names.forEach(function (namedItem, ix) {
                                            names2.push({
                                                name: namedItem.name,
                                                sheet_index: namedItem.sheet.position,
                                                address: namedItem.range.address.split("!").pop(),
                                                scope_sheet_name: null,
                                                scope_sheet_index: null,
                                                book_scope: namedItem.book_scope,
                                            });
                                        });
                                        payload["names"] = names2;
                                        // Sheets
                                        payload["sheets"] = [];
                                        sheetsLoader = [];
                                        sheets.forEach(function (sheet) {
                                            sheet.load("name names");
                                            var lastCell;
                                            if (excludeArray.includes(sheet.name)) {
                                                lastCell = null;
                                            }
                                            else if (sheet.getUsedRange() !== undefined) {
                                                lastCell = sheet.getUsedRange().getLastCell().load("address");
                                            }
                                            else {
                                                lastCell = sheet.getRange("A1").load("address");
                                            }
                                            sheetsLoader.push({
                                                sheet: sheet,
                                                lastCell: lastCell,
                                            });
                                        });
                                        return [4 /*yield*/, context.sync()];
                                    case 8:
                                        _c.sent();
                                        sheetsLoader.forEach(function (item, ix) {
                                            if (!excludeArray.includes(item["sheet"].name)) {
                                                var range = void 0;
                                                range = item["sheet"]
                                                    .getRange("A1:".concat(item["lastCell"].address))
                                                    .load("values, numberFormatCategories");
                                                sheetsLoader[ix]["range"] = range;
                                                // Names (sheet scope)
                                                sheetsLoader[ix]["names"] = item["sheet"].names.load("name, type");
                                            }
                                        });
                                        return [4 /*yield*/, context.sync()];
                                    case 9:
                                        _c.sent();
                                        namesSheetScope = [];
                                        sheetsLoader.forEach(function (item) {
                                            if (!excludeArray.includes(item["sheet"].name)) {
                                                item["names"].items.forEach(function (namedItem) {
                                                    namesSheetScope.push({
                                                        name: namedItem.name,
                                                        sheet: namedItem.getRange().worksheet.load("position"),
                                                        range: namedItem.getRange().load("address"),
                                                        scope_sheet: namedItem.worksheet.load("name, position"),
                                                        book_scope: false,
                                                    });
                                                });
                                            }
                                        });
                                        return [4 /*yield*/, context.sync()];
                                    case 10:
                                        _c.sent();
                                        namesSheetsScope2 = [];
                                        namesSheetScope.forEach(function (namedItem) {
                                            namesSheetsScope2.push({
                                                name: namedItem.name,
                                                sheet_index: namedItem.sheet.position,
                                                address: namedItem.range.address.split("!").pop(),
                                                scope_sheet_name: namedItem.scope_sheet.name,
                                                scope_sheet_index: namedItem.scope_sheet.position,
                                                book_scope: namedItem.book_scope,
                                            });
                                        });
                                        // Add sheet scoped names to book scoped names
                                        payload["names"] = payload["names"].concat(namesSheetsScope2);
                                        _loop_1 = function (item) {
                                            var sheet, values, categories_1, tablesArray, tables, tablesLoader, _d, _e, table, _f, tablesLoader_1, table, picturesArray, shapes, _g, _h, shape;
                                            return __generator(this, function (_j) {
                                                switch (_j.label) {
                                                    case 0:
                                                        sheet = item["sheet"];
                                                        if (excludeArray.includes(item["sheet"].name)) {
                                                            values = [[]];
                                                        }
                                                        else {
                                                            values = item["range"].values;
                                                            if (Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
                                                                categories_1 = item["range"].numberFormatCategories;
                                                                // Handle dates
                                                                // https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/excel-samples#dates
                                                                values.forEach(function (valueRow, rowIndex) {
                                                                    var categoryRow = categories_1[rowIndex];
                                                                    valueRow.forEach(function (value, colIndex) {
                                                                        var category = categoryRow[colIndex];
                                                                        if ((category.toString() === "Date" ||
                                                                            category.toString() === "Time") &&
                                                                            typeof value === "number") {
                                                                            values[rowIndex][colIndex] = new Date(Math.round((value - 25569) * 86400 * 1000)).toISOString();
                                                                        }
                                                                    });
                                                                });
                                                            }
                                                        }
                                                        tablesArray = [];
                                                        if (!!excludeArray.includes(item["sheet"].name)) return [3 /*break*/, 3];
                                                        tables = sheet.tables.load([
                                                            "name",
                                                            "showHeaders",
                                                            "dataBodyRange",
                                                            "showTotals",
                                                            "style",
                                                            "showFilterButton",
                                                        ]);
                                                        return [4 /*yield*/, context.sync()];
                                                    case 1:
                                                        _j.sent();
                                                        tablesLoader = [];
                                                        for (_d = 0, _e = sheet.tables.items; _d < _e.length; _d++) {
                                                            table = _e[_d];
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
                                                        return [4 /*yield*/, context.sync()];
                                                    case 2:
                                                        _j.sent();
                                                        for (_f = 0, tablesLoader_1 = tablesLoader; _f < tablesLoader_1.length; _f++) {
                                                            table = tablesLoader_1[_f];
                                                            tablesArray.push({
                                                                name: table.name,
                                                                range_address: table.range.address.split("!").pop(),
                                                                header_row_range_address: table.showHeaders
                                                                    ? table.headerRowRange.address.split("!").pop()
                                                                    : null,
                                                                data_body_range_address: table.dataBodyRange.address
                                                                    .split("!")
                                                                    .pop(),
                                                                total_row_range_address: table.showTotals
                                                                    ? table.totalRowRange.address.split("!").pop()
                                                                    : null,
                                                                show_headers: table.showHeaders,
                                                                show_totals: table.showTotals,
                                                                table_style: table.style,
                                                                show_autofilter: table.showFilterButton,
                                                            });
                                                        }
                                                        _j.label = 3;
                                                    case 3:
                                                        picturesArray = [];
                                                        if (!!excludeArray.includes(item["sheet"].name)) return [3 /*break*/, 5];
                                                        shapes = sheet.shapes.load(["name", "width", "height", "type"]);
                                                        return [4 /*yield*/, context.sync()];
                                                    case 4:
                                                        _j.sent();
                                                        for (_g = 0, _h = sheet.shapes.items; _g < _h.length; _g++) {
                                                            shape = _h[_g];
                                                            if (shape.type == Excel.ShapeType.image) {
                                                                picturesArray.push({
                                                                    name: shape.name,
                                                                    height: shape.height,
                                                                    width: shape.width,
                                                                });
                                                            }
                                                        }
                                                        _j.label = 5;
                                                    case 5:
                                                        payload["sheets"].push({
                                                            name: item["sheet"].name,
                                                            values: values,
                                                            pictures: picturesArray,
                                                            tables: tablesArray,
                                                        });
                                                        return [2 /*return*/];
                                                }
                                            });
                                        };
                                        _i = 0, sheetsLoader_1 = sheetsLoader;
                                        _c.label = 11;
                                    case 11:
                                        if (!(_i < sheetsLoader_1.length)) return [3 /*break*/, 14];
                                        item = sheetsLoader_1[_i];
                                        return [5 /*yield**/, _loop_1(item)];
                                    case 12:
                                        _c.sent();
                                        _c.label = 13;
                                    case 13:
                                        _i++;
                                        return [3 /*break*/, 11];
                                    case 14: return [4 /*yield*/, fetch(url, {
                                            method: "POST",
                                            headers: headers,
                                            body: JSON.stringify(payload),
                                        })];
                                    case 15:
                                        response = _c.sent();
                                        if (!(response.status !== 200)) return [3 /*break*/, 17];
                                        return [4 /*yield*/, response.text()];
                                    case 16: throw _c.sent();
                                    case 17: return [4 /*yield*/, response.json()];
                                    case 18:
                                        rawData = _c.sent();
                                        _c.label = 19;
                                    case 19:
                                        if (!(rawData !== null)) return [3 /*break*/, 23];
                                        forceSync = ["sheet"];
                                        _loop_2 = function (action) {
                                            return __generator(this, function (_k) {
                                                switch (_k.label) {
                                                    case 0: return [4 /*yield*/, globalThis.callbacks[action.func](context, action)];
                                                    case 1:
                                                        _k.sent();
                                                        if (!forceSync.some(function (el) { return action.func.toLowerCase().includes(el); })) return [3 /*break*/, 3];
                                                        return [4 /*yield*/, context.sync()];
                                                    case 2:
                                                        _k.sent();
                                                        _k.label = 3;
                                                    case 3: return [2 /*return*/];
                                                }
                                            });
                                        };
                                        _a = 0, _b = rawData["actions"];
                                        _c.label = 20;
                                    case 20:
                                        if (!(_a < _b.length)) return [3 /*break*/, 23];
                                        action = _b[_a];
                                        return [5 /*yield**/, _loop_2(action)];
                                    case 21:
                                        _c.sent();
                                        _c.label = 22;
                                    case 22:
                                        _a++;
                                        return [3 /*break*/, 20];
                                    case 23: return [2 /*return*/];
                                }
                            });
                        }); })];
                case 3:
                    _g.sent();
                    return [3 /*break*/, 6];
                case 4:
                    error_1 = _g.sent();
                    console.error(error_1);
                    return [4 /*yield*/, (0,_alert__WEBPACK_IMPORTED_MODULE_4__.xlAlert)(error_1, "Error", "ok", "critical", "")];
                case 5:
                    _g.sent();
                    return [3 /*break*/, 6];
                case 6: return [2 /*return*/];
            }
        });
    });
}
function getRange(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var sheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    sheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    return [2 /*return*/, sheets.items[action["sheet_position"]].getRangeByIndexes(action.start_row, action.start_column, action.row_count, action.column_count)];
            }
        });
    });
}
function getSheet(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var sheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    sheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    return [2 /*return*/, sheets.items[action.sheet_position]];
            }
        });
    });
}
function getTable(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var sheets, tables;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    sheets = context.workbook.worksheets.load("items");
                    tables = sheets.items[action.sheet_position].tables.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    return [2 /*return*/, tables.items[parseInt(action.args[0].toString())]];
            }
        });
    });
}
function getShapeByType(context, sheetPosition, shapeIndex, shapeType) {
    return __awaiter(this, void 0, void 0, function () {
        var sheets, shapes, myshapes;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    sheets = context.workbook.worksheets.load("items");
                    shapes = sheets.items[sheetPosition].shapes.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    myshapes = shapes.items.filter(function (shape) { return shape.type === shapeType; });
                    return [2 /*return*/, myshapes[shapeIndex]];
            }
        });
    });
}
function registerCallback(callback) {
    globalThis.callbacks[callback.name] = callback;
}
// Functions map
// Didn't find a way to use registerCallback so that webpack won't strip out these
// functions when optimizing
var funcs = {
    setValues: setValues,
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
    setRangeName: setRangeName,
    namesAdd: namesAdd,
    nameDelete: nameDelete,
    runMacro: runMacro,
    rangeDelete: rangeDelete,
    rangeInsert: rangeInsert,
    rangeSelect: rangeSelect,
    rangeClearContents: rangeClearContents,
    rangeClearFormats: rangeClearFormats,
    rangeClear: rangeClear,
    addTable: addTable,
    setTableName: setTableName,
    resizeTable: resizeTable,
    showAutofilterTable: showAutofilterTable,
    showHeadersTable: showHeadersTable,
    showTotalsTable: showTotalsTable,
    setTableStyle: setTableStyle,
    copyRange: copyRange,
    sheetDelete: sheetDelete,
    sheetClear: sheetClear,
    sheetClearFormats: sheetClearFormats,
    sheetClearContents: sheetClearContents,
};
Object.assign(globalThis.callbacks, funcs);
// Callbacks
function setValues(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var dt, dtString, range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    action.values.forEach(function (valueRow, rowIndex) {
                        valueRow.forEach(function (value, colIndex) {
                            if (typeof value === "string" &&
                                value.length > 18 &&
                                value.includes("T")) {
                                dt = new Date(Date.parse(value));
                                // Excel on macOS does use the wrong locale if you set a custom one via
                                // macOS Settings > Date & Time > Open Language & Region > Apps
                                // as the date format seems to stick to the Region selected under General
                                // while toLocaleDateString then respects the specific selected language.
                                // Providing Office.context.contentLanguage fixes this but isn't available for
                                // Office Scripts
                                // https://learn.microsoft.com/en-us/office/dev/add-ins/develop/localization#match-datetime-format-with-client-locale
                                dtString = dt.toLocaleDateString(Office.context.contentLanguage);
                                // Note that adding the time will format the cell as Custom instead of Date/Time
                                // which xlwings currently doesn't translate to datetime when reading
                                if (dtString !== "Invalid Date") {
                                    if (dt.getHours() +
                                        dt.getMinutes() +
                                        dt.getSeconds() +
                                        dt.getMilliseconds() !==
                                        0) {
                                        dtString += " " + dt.toLocaleTimeString();
                                    }
                                    action.values[rowIndex][colIndex] = dtString;
                                }
                            }
                        });
                    });
                    return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.values = action.values;
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
function rangeClearContents(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.clear(Excel.ClearApplyTo.contents);
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
function rangeClearFormats(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.clear(Excel.ClearApplyTo.formats);
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
function rangeClear(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.clear(Excel.ClearApplyTo.all);
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
function addSheet(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var sheet;
        return __generator(this, function (_a) {
            if (action.args[1] !== null) {
                sheet = context.workbook.worksheets.add(action.args[1].toString());
            }
            else {
                sheet = context.workbook.worksheets.add();
            }
            sheet.position = parseInt(action.args[0].toString());
            return [2 /*return*/];
        });
    });
}
function setSheetName(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var sheets;
        return __generator(this, function (_a) {
            sheets = context.workbook.worksheets.load("items");
            sheets.items[action.sheet_position].name = action.args[0].toString();
            return [2 /*return*/];
        });
    });
}
function setAutofit(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range, range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!(action.args[0] === "columns")) return [3 /*break*/, 2];
                    return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.format.autofitColumns();
                    return [3 /*break*/, 4];
                case 2: return [4 /*yield*/, getRange(context, action)];
                case 3:
                    range = _a.sent();
                    range.format.autofitRows();
                    _a.label = 4;
                case 4: return [2 /*return*/];
            }
        });
    });
}
function setRangeColor(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.format.fill.color = action.args[0].toString();
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
function activateSheet(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var worksheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    worksheets = context.workbook.worksheets;
                    worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    worksheets.items[parseInt(action.args[0].toString())].activate();
                    return [2 /*return*/];
            }
        });
    });
}
function addHyperlink(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range, hyperlink;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    hyperlink = {
                        textToDisplay: action.args[1].toString(),
                        screenTip: action.args[2].toString(),
                        address: action.args[0].toString(),
                    };
                    range.hyperlink = hyperlink;
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
function setNumberFormat(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.numberFormat = [[action.args[0].toString()]];
                    return [2 /*return*/];
            }
        });
    });
}
function setPictureName(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var myshape;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getShapeByType(context, action.sheet_position, Number(action.args[0]), Excel.ShapeType.image)];
                case 1:
                    myshape = _a.sent();
                    myshape.name = action.args[1].toString();
                    return [2 /*return*/];
            }
        });
    });
}
function setPictureHeight(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var myshape;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getShapeByType(context, action.sheet_position, Number(action.args[0]), Excel.ShapeType.image)];
                case 1:
                    myshape = _a.sent();
                    myshape.height = Number(action.args[1]);
                    return [2 /*return*/];
            }
        });
    });
}
function setPictureWidth(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var myshape;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getShapeByType(context, action.sheet_position, Number(action.args[0]), Excel.ShapeType.image)];
                case 1:
                    myshape = _a.sent();
                    myshape.width = Number(action.args[1]);
                    return [2 /*return*/];
            }
        });
    });
}
function deletePicture(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var myshape;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getShapeByType(context, action.sheet_position, Number(action.args[0]), Excel.ShapeType.image)];
                case 1:
                    myshape = _a.sent();
                    myshape.delete();
                    return [2 /*return*/];
            }
        });
    });
}
function addPicture(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var imageBase64, colIndex, rowIndex, left, top, sheet, anchorCell, image;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    imageBase64 = action["args"][0].toString();
                    colIndex = Number(action["args"][1]);
                    rowIndex = Number(action["args"][2]);
                    left = Number(action["args"][3]);
                    top = Number(action["args"][4]);
                    return [4 /*yield*/, getSheet(context, action)];
                case 1:
                    sheet = _a.sent();
                    anchorCell = sheet
                        .getRangeByIndexes(rowIndex, colIndex, 1, 1)
                        .load("left, top");
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    left = Math.max(left, anchorCell.left);
                    top = Math.max(top, anchorCell.top);
                    image = sheet.shapes.addImage(imageBase64);
                    image.left = left;
                    image.top = top;
                    return [2 /*return*/];
            }
        });
    });
}
function updatePicture(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var imageBase64, sheet, image, imgName, imgLeft, imgTop, imgHeight, imgWidth, newImage;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    imageBase64 = action["args"][0].toString();
                    return [4 /*yield*/, getSheet(context, action)];
                case 1:
                    sheet = _a.sent();
                    return [4 /*yield*/, getShapeByType(context, action.sheet_position, Number(action.args[1]), Excel.ShapeType.image)];
                case 2:
                    image = _a.sent();
                    image = image.load("name, left, top, height, width");
                    return [4 /*yield*/, context.sync()];
                case 3:
                    _a.sent();
                    imgName = image.name;
                    imgLeft = image.left;
                    imgTop = image.top;
                    imgHeight = image.height;
                    imgWidth = image.width;
                    image.delete();
                    newImage = sheet.shapes.addImage(imageBase64);
                    newImage.name = imgName;
                    newImage.left = imgLeft;
                    newImage.top = imgTop;
                    newImage.height = imgHeight;
                    newImage.width = imgWidth;
                    return [2 /*return*/];
            }
        });
    });
}
function alert(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var myPrompt, myTitle, myButtons, myMode, myCallback;
        return __generator(this, function (_a) {
            myPrompt = action.args[0].toString();
            myTitle = action.args[1].toString();
            myButtons = action.args[2].toString();
            myMode = action.args[3].toString();
            myCallback = action.args[4].toString();
            (0,_alert__WEBPACK_IMPORTED_MODULE_4__.xlAlert)(myPrompt, myTitle, myButtons, myMode, myCallback);
            return [2 /*return*/];
        });
    });
}
function setRangeName(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    context.workbook.names.add(action.args[0].toString(), range);
                    return [2 /*return*/];
            }
        });
    });
}
function namesAdd(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var name, refersTo, sheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    name = action.args[0].toString();
                    refersTo = action.args[1].toString();
                    if (!(action.sheet_position === null)) return [3 /*break*/, 1];
                    context.workbook.names.add(name, refersTo);
                    return [3 /*break*/, 3];
                case 1:
                    sheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    sheets.items[action.sheet_position].names.add(name, refersTo);
                    _a.label = 3;
                case 3: return [2 /*return*/];
            }
        });
    });
}
function nameDelete(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var name, book_scope, scope_sheet_index, sheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    name = action.args[2].toString();
                    book_scope = Boolean(action.args[4]);
                    scope_sheet_index = Number(action.args[5]);
                    if (!(book_scope === true)) return [3 /*break*/, 1];
                    context.workbook.names.getItem(name).delete();
                    return [3 /*break*/, 3];
                case 1:
                    sheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 2:
                    _a.sent();
                    sheets.items[scope_sheet_index].names.getItem(name).delete();
                    _a.label = 3;
                case 3: return [2 /*return*/];
            }
        });
    });
}
function runMacro(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0: return [4 /*yield*/, (_a = globalThis.callbacks)[action.args[0].toString()].apply(_a, __spreadArray([context], action.args.slice(1), false))];
                case 1:
                    _b.sent();
                    return [2 /*return*/];
            }
        });
    });
}
function rangeDelete(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range, shift;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    shift = action.args[0].toString();
                    if (shift === "up") {
                        range.delete(Excel.DeleteShiftDirection.up);
                    }
                    else if (shift === "left") {
                        range.delete(Excel.DeleteShiftDirection.left);
                    }
                    return [2 /*return*/];
            }
        });
    });
}
function rangeInsert(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range, shift;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    shift = action.args[0].toString();
                    if (shift === "down") {
                        range.insert(Excel.InsertShiftDirection.down);
                    }
                    else if (shift === "right") {
                        range.insert(Excel.InsertShiftDirection.right);
                    }
                    return [2 /*return*/];
            }
        });
    });
}
function rangeSelect(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRange(context, action)];
                case 1:
                    range = _a.sent();
                    range.select();
                    return [2 /*return*/];
            }
        });
    });
}
function addTable(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var worksheets, mytable;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    worksheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    mytable = worksheets.items[action.sheet_position].tables.add(action.args[0].toString(), Boolean(action.args[1]));
                    if (action.args[2] !== null) {
                        mytable.style = action.args[2].toString();
                    }
                    if (action.args[3] !== null) {
                        mytable.name = action.args[3].toString();
                    }
                    return [2 /*return*/];
            }
        });
    });
}
function setTableName(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var mytable;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getTable(context, action)];
                case 1:
                    mytable = _a.sent();
                    mytable.name = action.args[1].toString();
                    return [2 /*return*/];
            }
        });
    });
}
function resizeTable(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var mytable;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getTable(context, action)];
                case 1:
                    mytable = _a.sent();
                    mytable.resize(action.args[1].toString());
                    return [2 /*return*/];
            }
        });
    });
}
function showAutofilterTable(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var mytable;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getTable(context, action)];
                case 1:
                    mytable = _a.sent();
                    mytable.showFilterButton = Boolean(action.args[1]);
                    return [2 /*return*/];
            }
        });
    });
}
function showHeadersTable(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var mytable;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getTable(context, action)];
                case 1:
                    mytable = _a.sent();
                    mytable.showHeaders = Boolean(action.args[1]);
                    return [2 /*return*/];
            }
        });
    });
}
function showTotalsTable(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var mytable;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getTable(context, action)];
                case 1:
                    mytable = _a.sent();
                    mytable.showTotals = Boolean(action.args[1]);
                    return [2 /*return*/];
            }
        });
    });
}
function setTableStyle(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var mytable;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getTable(context, action)];
                case 1:
                    mytable = _a.sent();
                    mytable.style = action.args[1].toString();
                    return [2 /*return*/];
            }
        });
    });
}
function copyRange(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var destination, _a, _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    destination = context.workbook.worksheets.items[parseInt(action.args[0].toString())].getRange(action.args[1].toString());
                    _b = (_a = destination).copyFrom;
                    return [4 /*yield*/, getRange(context, action)];
                case 1:
                    _b.apply(_a, [_c.sent()]);
                    return [2 /*return*/];
            }
        });
    });
}
function sheetDelete(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var worksheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    worksheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    worksheets.items[action.sheet_position].delete();
                    return [2 /*return*/];
            }
        });
    });
}
function sheetClear(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var worksheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    worksheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    worksheets.items[action.sheet_position]
                        .getRanges()
                        .clear(Excel.ClearApplyTo.all);
                    return [2 /*return*/];
            }
        });
    });
}
function sheetClearFormats(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var worksheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    worksheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    worksheets.items[action.sheet_position]
                        .getRanges()
                        .clear(Excel.ClearApplyTo.formats);
                    return [2 /*return*/];
            }
        });
    });
}
function sheetClearContents(context, action) {
    return __awaiter(this, void 0, void 0, function () {
        var worksheets;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    worksheets = context.workbook.worksheets.load("items");
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    worksheets.items[action.sheet_position]
                        .getRanges()
                        .clear(Excel.ClearApplyTo.contents);
                    return [2 /*return*/];
            }
        });
    });
}

xlwings = __webpack_exports__;
/******/ })()
;
//# sourceMappingURL=xlwings.js.map