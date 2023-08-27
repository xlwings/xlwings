// This file is unused and would only be used in connection with a separate Runtime (FunctionFile in manifest)
// Activating it would require npm start to also build commands.html
import * as xlwings from "./src/xlwings";

Office.onReady(function (info) {});

function myCallback(arg: string) {
  console.log(`You selected ${arg} from commands.ts!`);
}
xlwings.registerCallback(myCallback);

async function hello(event: Office.AddinCommands.Event) {
  console.log("Called 'run' from commands.ts")
  await xlwings.runPython(window.location.origin + "/hello");
  event.completed();
}
Office.actions.associate("run", hello);
