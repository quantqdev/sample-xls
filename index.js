import XLSX from "xlsx-js-style";

import fs from "fs";

const dir = "./out";
if (!fs.existsSync(dir)) {
  fs.mkdirSync(dir, { recursive: true });
}

// STEP 1: Create a new workbook
const wb = XLSX.utils.book_new();

// STEP 2: Create data rows and styles
const firstRow = [
  { v: "Courier: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
  { v: "bold & color", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } } } },
  { v: "fill: color", t: "s", s: { fill: { fgColor: { rgb: "E9E9E9" } } } },
  { v: "line\nbreak", t: "s", s: { alignment: { wrapText: true } } },
];

const rows = [firstRow];
rows.push(["Name", "Birth year", "Sex"]);
for (let i = 0; i < 100; i++) {
  const sex = Math.random() > 0.5 ? "Male" : "Female";
  rows.push([
    { v: `User ${i}` },
    { v: 1990 + i },
    { v: sex, s: { fill: { fgColor: { rgb: sex === "Male" ? "0000FF" : "00FF00" } } } },
  ]);
}

// STEP 3: Create worksheet with rows; Add worksheet to workbook
const ws = XLSX.utils.aoa_to_sheet(rows);
XLSX.utils.book_append_sheet(wb, ws, "readme demo");

// STEP 4: Write Excel file
XLSX.writeFile(wb, `${dir}/xlsx-js-style-demo.xlsx`);
