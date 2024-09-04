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
          let combinedArray=id64Bytes7.map((element,index)=>element+id64Bytes6[index]+id64Bytes5[index]+id64Bytes4[index]+id64Bytes3[index]+id64Bytes2[index]+id64Bytes1[index]+id64Bytes0[index]);
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
     let stringArray= combinedArray.map(num => [num]);
    // Set the values of the range with the stringArray
    range0x64.numberFormat = '@'; // '@' sets the format to Text
        range0x64.values = stringArray;
     
        // 计算XOR
        let resultArray = xorAdjacentElementsDirect(combinedArray);
        console.log(`Result XOR Array:/${resultArray}`);
        let valTang=sumBinaryColumns(resultArray);
        //console.log(valTang);
    // Load the range properties and sync
      range0x64.load("values");
      const sheetTang = context.workbook.worksheets.getItem("Tang");
      const rangeBitheader = sheetTang.getRange("A1:BL1");
      const tangValuerange = sheetTang.getRange("A2:BL2");
      // 填充单元格
      let bitLabels = [];
      for (let i = 63; i >= 0; i--) {
        bitLabels.push(`bit${i}`);
      }
      
      // 设置单元格值
      rangeBitheader.values = [bitLabels];
      tangValuerange.values=[splitStringIntoSubarrays(valTang).map(Number)];
      console.log(splitStringIntoSubarrays(valTang));


      let dataTangRange = sheetTang.getRange("A1:BL2");
      let chart = sheetTang.charts.add(
      Excel.ChartType.line, 
      dataTangRange, 
      Excel.ChartSeriesBy.auto);

      chart.title.text = "TANG";
      chart.legend.position = Excel.ChartLegendPosition.right;
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
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

function xorAdjacentElementsDirect(binaryStringArray) {
  // 获取字符串的长度
  const length = binaryStringArray[0].length;
  // 创建一个新数组，长度为原数组长度减1
  let xorArray = [];

  // 遍历原数组，除了最后一个元素
  for (let i = 0; i < binaryStringArray.length - 1; i++) {
    // 遍历每个字符串的每个字符
    let xorString=xorBinaryStrings(binaryStringArray[i],binaryStringArray[i+1]);
    // 将完整的XOR结果字符串添加到新数组中
    xorArray.push(xorString);
  }

  return xorArray;
}

function xorBinaryStrings(str1, str2) {
  // 确保两个字符串长度相同，较短的字符串前面补0
  let maxLength = Math.max(str1.length, str2.length);
  let xorResult = '';
  for (let i = 0; i < maxLength; i++) {
    // 对应位进行XOR操作
    xorResult += (parseInt(str1[i], 10) ^ parseInt(str2[i], 10)).toString();
  }
  return xorResult;
}
function sumBinaryColumns(binaryArray) {
  // 确定数组中最长的字符串长度
  let maxLength = Math.max(...binaryArray.map(str => str.length));
  
  // 初始化每列的和为0，长度为最长字符串的长度
  let sum = new Array(maxLength).fill(0);

  // 遍历数组的每一行
  for (let row of binaryArray) {
    // 遍历每一列
    for (let col = 0; col < maxLength; col++) {
      // 如果当前位置有值（即不是超出当前行字符串长度的填充0），则累加到对应的列
      if (col < row.length) {
        sum[col] += parseInt(row[col], 2);
      }
    }
  }

  // 将每列的和转换为十进制字符串并返回
  return sum.join(',');
}
function splitStringIntoSubarrays(str) {
  // 使用split方法按逗号分割字符串
  let numbersArray = str.split(',');
  // 使用map方法将每个分割后的字符串放入一个单独的数组中
  let subarrays = numbersArray.map(number => [number]);
  return subarrays;
}