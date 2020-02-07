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

function getMonday(date) {
  while (date.getDay() != 1) {
    date.setDate(date.getDate() + 1);
  }
  return date;
}

function buildOutline(today, activePage) {
  var weekDays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  var title = "";
  for (var day = 0; day < weekDays.length; day++) {
    var dd = String(today.getDate()).padStart(2, "0");
    var mm = String(today.getMonth() + 1).padStart(2, "0");

    if (today.getDay() == 1) {
      title += mm + "/" + dd + " - ";
    }
    if (today.getDay() == 5) {
      title += mm + "/" + dd;
    }
    var html = "<h2>" + weekDays[day] + " " + mm + "/" + dd + "</h2><p>: </p><p>: </p>";
    var activeOutline = activePage.addOutline(40, 90 + 90 * day, html);
    activeOutline.paragraphs.getItemAt(1).addNoteTag("ToDo", "Normal");
    activeOutline.paragraphs.getItemAt(2).addNoteTag("ToDo", "Normal");
    today.setDate(today.getDate() + 1);
  }
  return title;
}

export async function run() {
  try {
    await OneNote.run(async context => {
      // Get the current page.
      var activePage = context.application.getActivePage();

      var newPage = activePage.insertPageAsSibling("Before", "Daily Tasks");

      // Queue a command to add an outline to the page.

      context.application.navigateToPage(newPage);
      activePage = context.application.getActivePage();
      var today = new Date();
      today = getMonday(today);
      var title = buildOutline(today, activePage);

      activePage.title = title;
      // Run the queued commands, and return a promise to indicate task completion.
      return context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
