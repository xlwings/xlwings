/**
 * Required Notice: Copyright (C) Zoomer Analytics GmbH.
 *
 * xlwings PRO is dual-licensed under one of the following licenses:
 *
 * * PolyForm Noncommercial License 1.0.0 (for noncommercial use):
 *   https://polyformproject.org/licenses/noncommercial/1.0.0
 * * xlwings PRO License (for commercial use):
 *   https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt
 *
 * Commercial licenses can be purchased at https://www.xlwings.org
 */

const debug = false;
let invocations = new Set();
let bodies = new Set();
let runtime;
let contentLanguage;
let socket = null;

Office.onReady(function (info) {
  // Socket.io
  socket = globalThis.socket ? globalThis.socket : null;

  if (socket !== null) {
    socket.on("disconnect", () => {
      if (debug) {
        console.log("disconnect");
      }
      for (let invocation of invocations) {
        invocation.setResult([["Stream disconnected"]]);
      }
      invocations.clear();
    });

    socket.on("connect", () => {
      // Without this, you'd have to hit Ctrl+Alt+F9, which isn't available on the web
      if (debug) {
        console.log("connect");
      }
      for (let body of bodies) {
        socket.emit("xlwings:function-call", body);
      }
    });
  }

  // Runtime version
  if (
    Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.4")
  ) {
    runtime = "1.4";
  } else if (
    Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.3")
  ) {
    runtime = "1.3";
  } else if (
    Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.2")
  ) {
    runtime = "1.2";
  } else {
    runtime = "1.1";
  }

  // Content Language
  contentLanguage = Office.context.contentLanguage;
});

async function base() {
  // Arguments
  let argsArr = Array.prototype.slice.call(arguments);
  let funcName = argsArr[0];
  let isStreaming = argsArr[1];
  let args = argsArr.slice(2, -1);
  let invocation = argsArr[argsArr.length - 1];

  // Body
  let body = {
    func_name: funcName,
    args: args,
    caller_address: invocation.address, // not available for streaming functions
    content_language: contentLanguage,
    version: "placeholder_xlwings_version",
    runtime: runtime,
  };

  // Streaming functions communicate via socket.io
  if (isStreaming) {
    if (socket === null) {
      console.error(
        "To enable streaming functions, you need to load the socket.io js client before xlwings.min.js and custom-functions-code"
      );
      return;
    }
    let taskKey = `${funcName}_${args}`;
    body.task_key = taskKey;
    socket.emit("xlwings:function-call", body);
    if (debug) {
      console.log(`emit xlwings:function-call ${funcName}`);
    }
    invocation.setResult([["Waiting for stream..."]]);

    socket.off(`xlwings:set-result-${taskKey}`);
    socket.on(`xlwings:set-result-${taskKey}`, (data) => {
      invocation.setResult(data.result);
      if (debug) {
        console.log(`Set Result`);
      }
    });

    invocations.add(invocation);
    bodies.add(body);

    return;
  }

  // Normal functions communicate via REST API
  let headers = {};
  headers["Content-Type"] = "application/json";
  headers["Authorization"] =
    typeof globalThis.getAuth === "function" ? await globalThis.getAuth() : "";

  // Socket.io
  const sid = socket && socket.id ? socket.id.toString() : null;
  headers["sid"] = sid;

  try {
    let response = await fetch(
      window.location.origin + "placeholder_custom_functions_call_path",
      {
        method: "POST",
        headers: headers,
        body: JSON.stringify(body),
      }
    );
    if (!response.ok) {
      let errMsg = await response.text();
      return showError(errMsg);
    } else {
      let responseData = await response.json();
      return responseData.result;
    }
  } catch (error) {
    return showError(error.toString());
  }
}

function showError(errorMessage) {
  if (
    Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.2")
  ) {
    // Error message is only visible by hovering over the error flag!
    let excelError = new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      errorMessage
    );
    throw excelError;
  } else {
    return [[errorMessage]];
  }
}
