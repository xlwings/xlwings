import * as xlwings from "./src/xlwings";

Office.onReady(function (info) {});

document.getElementById("run").addEventListener("click", run);
document.getElementById("show-alert").addEventListener("click", showAlert);
document.getElementById("integration-test-read").addEventListener("click", integrationTestRead);
document.getElementById("integration-test-write").addEventListener("click", integrationTestWrite);

function myCallback(arg: string) {
  console.log(`You selected ${arg} from taskpane.ts!`);
}
xlwings.registerCallback(myCallback);

async function run() {
  console.log("Called 'run' from taskpane.ts")
  await xlwings.runPython(window.location.origin + "/hello");
}

async function showAlert() {
  console.log("Called 'showAlert' from taskpane.ts")
  await xlwings.runPython(window.location.origin + "/show-alert");
}

async function integrationTestRead() {
  console.log("Called 'integrationTestRead' from taskpane.ts")
  await xlwings.runPython(window.location.origin + "/integration-test-read");
}

async function integrationTestWrite() {
  console.log("Called 'integrationTestWrite' from taskpane.ts")
  await xlwings.runPython(window.location.origin + "/integration-test-write");
}
