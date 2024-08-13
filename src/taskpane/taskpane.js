/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      let currentSheet = context.workbook.worksheets.getActiveWorksheet();
      let traceTable = currentSheet.tables.getItem("Trace");
      // Get data from the header row.
      let headerRange = traceTable.getHeaderRowRange().load("values");
    await context.sync();

    let headerValues = headerRange.values;
    let headerValues2 = headerRange.values[0];
    console.log(headerValues);
    console.log(headerValues2);
    console.log(headerValues2.length);
    //You can iterate through headers and perform operations on specific columns based on their names:
    headerValues2.forEach(async (header, index) => {
      if (header === "ID") {
          // Perform operations on the "ID" column
          let idColumnRange = traceTable.getDataBodyRange().getColumn(index).load("values,rowIndex,rowCount");
          await context.sync();
          let id64RowValues = idColumnRange.values.flat().filter(el => el === 64);
          let startRowIndex=idColumnRange.rowIndex;
          let endRowIndex=idColumnRange.rowCount + startRowIndex -1;
          let count = id64RowValues.length;
          console.log(`Count of rows with ID=64: ${count}`);
          console.log(`Value Array of rows with ID=64: ${id64RowValues}`);
          console.log(`All row index: ${startRowIndex} to ${endRowIndex}`);
      }
  });
    });
  } catch (error) {
    console.error(error);
  }
}
