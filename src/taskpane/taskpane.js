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
      const sheet0x64 = context.workbook.worksheets.getItem("0x64Trace");
          // Calculate the range based on the length of the array
       const numberOfRows = combinedArray.length;
       console.log(`Row indices with ID=64: ${numberOfRows}`);
       const startCell = "A2"; // Starting from cell A1
       const endCell = `A${numberOfRows+1}`; // Calculate the ending cell based on the array length
       const rangeAddress = `${startCell}:${endCell}`; // Define the range address

    // Select the range where you want to fill the array
        const range0x64 = sheet0x64.getRange(rangeAddress);
     let stringArray= combinedArray.map(num => [num.toString()]);
    // Set the values of the range with the stringArray
    range0x64.numberFormat = '@'; // '@' sets the format to Text
        range0x64.values = stringArray;
    
    // Load the range properties and sync
    range0x64.load("values");
       await context.sync();

      console.log("Array filled successfully!");

      }
  });
    });
  } catch (error) {
    console.error(error);
  }
}
// API key and endpoint URL
let apiKey = "94b84be7424d4fbd9XXXXXXX";//Enter your API key here
let apiUrl ="https://api.openweathermap.org/data/2.5/weather?units=metric&lang=en";
// DOM elements
let searchBox = document.querySelector(".search input");
let searchButton = document.querySelector(".search button");
let weather_icon = document.querySelector(".weather-icon");
// Variable to store Celsius value
let cel;
// Function to check the weather for a city
async function checkWeather(city) {
  try {
    // Make API call to fetch weather data
    const response = await fetch(`${apiUrl}&q=${city}&appid=${apiKey}`);

    if (!response.ok) {
      throw new Error("Unable to fetch weather data.");
    }

    // Parse the response JSON
    const data = await response.json();

    // Update the DOM with weather information
    document.querySelector(".city").innerHTML = data.name;
    const tempCelcius = Math.round(data.main.temp);
    document.querySelector(".temp").innerHTML = tempCelcius + "°C";
    document.querySelector(".humidity").innerHTML = data.main.humidity + "%";
    document.querySelector(".pressure").innerHTML = data.main.pressure;

    // Store the Celsius value
    cel = tempCelcius;
  } catch (error) {
    // Display error message and hide weather section
    document.querySelector(".err").style.display = "block";
    document.querySelector(".weather").style.display = "none";
    console.error(error);
  }
}
// Event listener for search button click
searchButton.addEventListener("click", () => {
  const city = searchBox.value.trim();
  if (city !== "") {
    // Call checkWeather function with the entered city
    checkWeather(city);
  }
});

// Event listener for Fahrenheit button click
document.getElementById("farenheit").addEventListener("click", () => {
  // Convert Celsius to Fahrenheit and update the HTML
  if (cel !== undefined) {
    let fer = Math.floor(cel * 1.8 + 32);
    document.querySelector(".temp").innerHTML = fer + "°F";
  }
});

// Event listener for Celsius button click
document.getElementById("celcius").addEventListener("click", () => {
  // Restore the Celsius value and update the HTML
  if (cel !== undefined) {
    document.querySelector(".temp").innerHTML = cel + "°C";
  }
});