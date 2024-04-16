const fs = require("fs");
const XLSX = require("xlsx");

// Membaca file xlsx
const workbook = XLSX.readFile("Book1.xlsx");

// Mengambil nama sheet yang akan digunakan
const sheetName = workbook.SheetNames[0];

// Mengambil data dari sheet
const worksheet = workbook.Sheets[sheetName];

let temp = [];
const getAHS = Object.keys(worksheet).filter((key) => key.includes("A"));
const getTitle = getAHS.map((el) => el.replace("A", "B"));
const getPrice = getAHS
  .map((el) => el.replace("A", "F"))
  .map((el) => parseInt(el.slice(1) - 2))
  .map((el) => "F" + el);

const getLastPrice = Object.keys(worksheet).filter((key) => key.includes("F"));

// console.log("getAHS : ", getAHS);
// console.log("getTitle : ", getTitle);
// console.log("getPrice : ", getPrice);
console.log("\n<<<<<<<==========================>>>>>>>>>>>>\n");

for (let i = 0; i < getAHS.length; i++) {
  let price = 0;
  if (i + 1 === getAHS.length) {
    price = worksheet[getLastPrice[getLastPrice.length - 1]].v;
    console.log(`price ${getLastPrice[getLastPrice.length - 1]}: ${price}`);
    temp.push({
      AHS_code: worksheet[getAHS[i]].v,
      task_name: worksheet[getTitle[i]].v,
      price: worksheet[getLastPrice[getLastPrice.length - 1]].v,
    });
  } else {
    price = worksheet[getPrice[i + 1]].v;
    console.log(`price ${getPrice[i + 1]}: ${price}`);
    temp.push({
      AHS_code: worksheet[getAHS[i]].v,
      task_name: worksheet[getTitle[i]].v,
      price: worksheet[getPrice[i + 1]].v,
    });
  }
}

// console.log("temp :", temp);
console.log("\n<<<<<<<==========================>>>>>>>>>>>>\n");

// Menyimpan data ke file JSON
fs.writeFileSync("data.json", JSON.stringify(temp, null, 2), "utf-8");

console.log("Data berhasil dienkstrak dan disimpan ke data.json");