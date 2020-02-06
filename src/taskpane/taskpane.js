/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await OneNote.run(async context => {
      // Get the current page.
      var activePage = context.application.getActivePage();

      var newPage = activePage.insertPageAsSibling("Before", "Daily Tasks");

      // Queue a command to add an outline to the page.

      context.application.navigateToPage(newPage);
      activePage = context.application.getActivePage();

      var mondayHtml = "<h2>Monday</h2>";
      var monOutline = activePage.addOutline(40, 90, mondayHtml);
      var monParagraphs = monOutline.paragraphs;
      monParagraphs.getItemAt(1).addNoteTag("ToDo", "Normal");
      // monParagraphs
      //   .getItemAt(1)
      //   .paragraphs.getItemAt(0)
      //   .addNoteTag("ToDo", "Normal");

      var tuesdayHtml = "<h2>Tuesday</h2>";
      activePage.addOutline(40, 180, tuesdayHtml);

      var wednesdayHtml = "<h2>Wednesday</h2>";
      activePage.addOutline(40, 270, wednesdayHtml);

      var thursdayHtml = "<h2>Thursday</h2>";
      activePage.addOutline(40, 360, thursdayHtml);

      var fridayHtml = "<h2>Friday</h2>";
      activePage.addOutline(40, 450, fridayHtml);

      // Run the queued commands, and return a promise to indicate task completion.
      return context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
