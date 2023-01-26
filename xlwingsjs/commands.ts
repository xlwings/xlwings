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
