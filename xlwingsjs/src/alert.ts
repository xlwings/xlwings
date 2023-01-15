let dialog;

function dialogCallback(asyncResult) {
  if (asyncResult.status == "failed") {
    switch (asyncResult.error.code) {
      case 12004:
        console.log("Domain is not trusted");
        break;
      case 12005:
        console.log("HTTPS is required");
        break;
      case 12007:
        console.log("A dialog is already opened.");
        break;
      case 12009:
        console.log("The user chose to ignore the dialog box.");
        break;
      case 12011:
        console.log("The user's browser configuration is blocking popups.");
        break;
      default:
        console.log(asyncResult.error.message);
        break;
    }
  } else {
    dialog = asyncResult.value;
    dialog.addEventHandler(
      Office.EventType.DialogMessageReceived,
      messageHandler
    );
    dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
  }
}

function messageHandler(arg) {
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

function eventHandler(arg) {
  switch (arg.error) {
    case 12002:
      console.log("Cannot load URL, no such page or bad URL syntax.");
      break;
    case 12003:
      console.log("HTTPS is required.");
      break;
    case 12006:
      console.log("Dialog closed by user");
      break;
    default:
      console.log("Undefined error in dialog window");
      break;
  }
}

export function xlAlert(
  prompt: string,
  title: string,
  buttons: string,
  mode: string,
  callback: string
) {
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
  Office.context.ui.displayDialogAsync(
    window.location.origin +
      `/alert?prompt=` +
      encodeURIComponent(`${prompt}`) +
      `&title=` +
      encodeURIComponent(`${title}`) +
      `&buttons=${buttons}&mode=${mode}&callback=${callback}`,
    { height: height, width: width, displayInIframe: true },
    dialogCallback
  );
}
