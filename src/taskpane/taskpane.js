/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, setTimeout, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const iframe = document.createElement(`iframe`);
  const html = `<html><body>Foo</body></html>`;

  document.getElementById("iframe-content").appendChild(iframe);

  iframe.contentWindow.document.open();
  iframe.contentWindow.document.write(html);
  iframe.contentWindow.document.close();

  // eslint-disable-next-line no-undef
  setTimeout(() => {
    console.log("Displaying dialog");

    Office.context.ui.displayDialogAsync("https://www.contoso.com", (result) => {
      if (result.error) {
        console.error(result.error);
      }
    });
  }, 3000);
}
