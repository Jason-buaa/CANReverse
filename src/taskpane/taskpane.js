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
    //You can iterate through headers and perform operations on specific columns based on their names:
    headerValues2.forEach(async (header, index) => {
      if (header === "ID") {
          // Perform operations on the "ID" column
          let idColumnRange = traceTable.getDataBodyRange().getColumn(index).load("values");
          await context.sync();
          
          let count = idColumnRange.values.flat().filter(value => value === 64).length;
          console.log(`Count of rows with ID=64: ${count}`);
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