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
          let byte1ColumnRange = traceTable.getDataBodyRange().getColumn(index+5).load("values,rowIndex,rowCount");
          let byte2ColumnRange = traceTable.getDataBodyRange().getColumn(index+6).load("values,rowIndex,rowCount");
          let byte3ColumnRange = traceTable.getDataBodyRange().getColumn(index+7).load("values,rowIndex,rowCount");
          let byte4ColumnRange = traceTable.getDataBodyRange().getColumn(index+8).load("values,rowIndex,rowCount");
          let byte5ColumnRange = traceTable.getDataBodyRange().getColumn(index+9).load("values,rowIndex,rowCount");
          let byte6ColumnRange = traceTable.getDataBodyRange().getColumn(index+10).load("values,rowIndex,rowCount");
          let byte7ColumnRange = traceTable.getDataBodyRange().getColumn(index+10).load("values,rowIndex,rowCount");
          await context.sync();
          let idValues =idColumnRange.values.flat();
          let id64RowIndices = idValues
          .map((el, idx) => el === 64 ? idx : -1) // +2 to adjust for Excel row numbers (assuming table starts at row 1)
          .filter(index => index !== -1); // Filter out the -1 values
          let id64Timestamps= id64RowIndices.map(idx => timestampColumnRange.values[idx]);
          let id64Bytes0=id64RowIndices.map(idx=>byte0ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let id64Bytes1=id64RowIndices.map(idx=>byte1ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let id64Bytes2=id64RowIndices.map(idx=>byte2ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let id64Bytes3=id64RowIndices.map(idx=>byte3ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let id64Bytes4=id64RowIndices.map(idx=>byte4ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let id64Bytes5=id64RowIndices.map(idx=>byte5ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let id64Bytes6=id64RowIndices.map(idx=>byte6ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let id64Bytes7=id64RowIndices.map(idx=>byte7ColumnRange.values[idx]).map(hexStr=>parseInt(hexStr,16).toString(2).padStart(8,'0'));
          let combinedArray=id64Bytes0.map((element,index)=>element+id64Bytes1[index]+id64Bytes2[index]+id64Bytes3[index]+id64Bytes4[index]+id64Bytes5[index]+id64Bytes6[index]+id64Bytes7[index]);
      //let count = id64RowIndices.length;
      //console.log(`Count of rows with ID=64: ${count}`);
      //console.log(`Row indices with ID=64: ${id64RowIndices}`);
      //console.log(`Row indices with ID=64: ${id64Timestamps}`);
      console.log(`Row indices with ID=64: ${combinedArray}`);
      }
  });
    });
  } catch (error) {
    console.error(error);
  }
}
