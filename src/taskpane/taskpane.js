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
          let timestampColumnRange = traceTable.getDataBodyRange().getColumn(index-2).load("values,rowIndex,rowCount");
          let byte0ColumnRange = traceTable.getDataBodyRange().getColumn(index+4).load("values,rowIndex,rowCount");
          await context.sync();
          let idValues =idColumnRange.values.flat();
          let id64RowIndices = idValues
          .map((el, idx) => el === 64 ? idx : -1) // +2 to adjust for Excel row numbers (assuming table starts at row 1)
          .filter(index => index !== -1); // Filter out the -1 values
          let id64Timestamps= id64RowIndices.map(idx => timestampColumnRange.values[idx]);
          let id64Bytes0=id64RowIndices.map(idx=>byte0ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16));
      let count = id64RowIndices.length;
      console.log(`Count of rows with ID=64: ${count}`);
      console.log(`Row indices with ID=64: ${id64RowIndices}`);
      console.log(`Row indices with ID=64: ${id64Timestamps}`);
      console.log(`Row indices with ID=64: ${id64Bytes0}`);
      }
  });
    });
  } catch (error) {
    console.error(error);
  }
}
