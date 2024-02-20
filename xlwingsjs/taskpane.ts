import * as xlwings from "./src/xlwings";

Office.onReady(function (info) {});

document.getElementById("run").addEventListener("click", hello);
document.getElementById("show-alert").addEventListener("click", showAlert);
document
  .getElementById("integration-test-read")
  .addEventListener("click", integrationTestRead);
document
  .getElementById("integration-test-write")
  .addEventListener("click", integrationTestWrite);

globalThis.getAuth = async function () {
  // Required by custom functions, but should also be used with runPython.
  // Uses SSO to provide an Azure AD access token as Authorization header
  // if the manifest has been set up accordingly. The access token must
  // then be verified by the backend, see:
  // https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins#validate-the-access-token
  // Replace this function with your own logic if you don't want to use SSO.
  // NOTE: the access token is also an identity token

  return await xlwings.getAccessToken();
  // return "";
};

function myCallback(arg: string) {
  console.log(`You selected ${arg} from taskpane.ts!`);
}
xlwings.registerCallback(myCallback);

async function hello() {
  console.log("Called 'run' from Task pane");
  await xlwings.runPython(window.location.origin + "/hello", {
    auth: await globalThis.getAuth(),
  });
}

// Note: when bound to ribbon, this needs event.completed() and Office.actions.associate
async function helloRibbon(event: Office.AddinCommands.Event) {
  console.log("Called 'helloRibbon' from ribbon");
  await xlwings.runPython(window.location.origin + "/hello");
  event.completed();
}
Office.actions.associate("run-hello", helloRibbon);

async function showAlert() {
  console.log("Called 'showAlert' from taskpane.ts");
  await xlwings.runPython(window.location.origin + "/show-alert");
}

async function integrationTestRead() {
  console.log("Called 'integrationTestRead' from taskpane.ts");
  await xlwings.runPython(window.location.origin + "/integration-test-read");
}

async function integrationTestWrite() {
  console.log("Called 'integrationTestWrite' from taskpane.ts");
  await xlwings.runPython(window.location.origin + "/integration-test-write");
}
