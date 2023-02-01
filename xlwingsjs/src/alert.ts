﻿// https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins

let dialog: Office.Dialog;

function dialogCallback(asyncResult: Office.AsyncResult<Office.Dialog>) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log(`${asyncResult.error.message} [${asyncResult.error.code}]`);
  } else {
    dialog = asyncResult.value;
    // Handle messages and events
    dialog.addEventHandler(
      Office.EventType.DialogMessageReceived,
      processMessage
    );
    dialog.addEventHandler(
      Office.EventType.DialogEventReceived,
      processDialogEvent
    );
  }
}

function processMessage(arg: Office.DialogParentMessageReceivedEventArgs) {
  dialog.close();
  let [selection, callback] = arg.message.split("|");
  if (callback !== "" && callback in globalThis.funcs) {
    globalThis.funcs[callback](selection);
  } else {
    if (callback !== "" && !(callback in globalThis.funcs)) {
      throw new Error(
        `Didn't find callback '${callback}'! Make sure to run xlwings.registerCallback(${callback}) before calling runPython.`
      );
    }
  }
}

function processDialogEvent(arg: { error: number }) {
  switch (arg.error) {
    case 12002:
      console.log(
        "The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid."
      );
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

export function xlAlert(
  prompt: string,
  title: string,
  buttons: string,
  mode: string,
  callback: string,
  width?: number,
  height?: number,
  url?: string,
) {
  if (typeof width === 'undefined') {
    console.log('width undef')
    if (Office.context.platform.toString() === "OfficeOnline") {
      width = 28;
    } else if (Office.context.platform.toString() === "PC") {
      width = 28; // seems to have a wider min width
    } else {
      width = 32;
    }
  }
  if (typeof height === 'undefined') {
    if (Office.context.platform.toString() === "OfficeOnline") {
      height = 36;
    } else if (Office.context.platform.toString() === "PC") {
      height = 40;
    } else {
      height = 30;
    }
  }
  if (typeof url === 'undefined') {
    url = window.location.origin +
    `/xlwings/alert?prompt=` +
    encodeURIComponent(`${prompt}`) +
    `&title=` +
    encodeURIComponent(`${title}`) +
    `&buttons=${buttons}&mode=${mode}&callback=${callback}`
  }
  Office.context.ui.displayDialogAsync(
    url,
    { height: height, width: width, displayInIframe: true },
    dialogCallback
  );
}