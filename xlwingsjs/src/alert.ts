// https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins

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
  if (callback !== "" && callback in globalThis.callbacks) {
    globalThis.callbacks[callback](selection);
  } else {
    if (callback !== "" && !(callback in globalThis.callbacks)) {
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

export async function xlAlert(
  prompt: string,
  title: string,
  buttons: string,
  mode: string,
  callback: string
) {
  await Office.onReady();
  let width: number;
  let height: number;
  if (Office.context.platform.toString() === "OfficeOnline") {
    width = 28;
    height = 36;
  } else if (Office.context.platform.toString() === "PC") {
    width = 28; // seems to have a wider min width
    height = 40;
  } else {
    width = 32;
    height = 30;
  }

  // See xlwings-server repo for how app-path is provided
  const appPathElement = document.getElementById("app-path");
  const appPath = appPathElement ? JSON.parse(appPathElement.textContent) : null;
  Office.context.ui.displayDialogAsync(
    window.location.origin +
      (appPath && appPath.appPath !== "" ? `/${appPath.appPath}` : "") +
      `/xlwings/alert?prompt=` +
      encodeURIComponent(`${prompt}`) +
      `&title=` +
      encodeURIComponent(`${title}`) +
      `&buttons=${buttons}&mode=${mode}&callback=${callback}`,
    { height: height, width: width, displayInIframe: true },
    dialogCallback
  );
}
