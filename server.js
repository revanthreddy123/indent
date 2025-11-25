// server.js
const express = require("express");
const cors = require("cors");
const path = require("path");
const Excel = require("exceljs");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(express.json());

// Excel File Path
const EXCEL_FILE_PATH = path.join(__dirname, "./Indent.xlsx");

// Load workbook helper
async function loadWorkbook() {
  if (!fs.existsSync(EXCEL_FILE_PATH)) {
    throw new Error(`Excel file not found at: ${EXCEL_FILE_PATH}`);
  }

  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE_PATH);
  return workbook;
}

/* ---------------------------------------------------
   1️⃣  CREATE OR OPEN DAILY SHEET
--------------------------------------------------- */
app.post("/api/sheet", async (req, res) => {
  try {
    const { date } = req.body;
    if (!date) return res.status(400).json({ error: "Missing date" });

    const sheetName = String(date).substring(0, 31);

    const workbook = await loadWorkbook();

    // If already exists → OK
    if (workbook.getWorksheet(sheetName)) {
      return res.json({ status: "exists", sheetName });
    }

    // MASTER lookup
    const template =
      workbook.getWorksheet("MASTER") ||
      workbook.getWorksheet("Master") ||
      workbook.getWorksheet("Template") ||
      workbook.getWorksheet(1);

    if (!template) {
      return res.status(500).json({ error: "MASTER sheet not found!" });
    }

    // Create new sheet
    const newSheet = workbook.addWorksheet(sheetName);

    // Copy rows including formulas
    template.eachRow({ includeEmpty: true }, (row, rowNum) => {
      const newRow = newSheet.getRow(rowNum);
      newRow.values = row.values;
      if (row.height) newRow.height = row.height;
    });

    // Copy merged cells (if any)
    if (template._merges) {
      template._merges.forEach(m => newSheet.mergeCells(m));
    }

    workbook.calcProperties.fullCalcOnLoad = true;

    await workbook.xlsx.writeFile(EXCEL_FILE_PATH);

    return res.json({ status: "created", sheetName });
  } catch (err) {
    console.error("/api/sheet ERR:", err);
    res.status(500).json({ error: err.message });
  }
});

/* ---------------------------------------------------
   2️⃣  FETCH EXISTING COMPANY VALUES 
   POST: { date, company }
--------------------------------------------------- */
app.post("/api/get-company", async (req, res) => {
  try {
    const { date, company } = req.body;
    if (!date || !company)
      return res.status(400).json({ error: "Missing fields" });

    const sheetName = String(date).substring(0, 31);
    const workbook = await loadWorkbook();
    const sheet = workbook.getWorksheet(sheetName);

    if (!sheet) return res.status(404).json({ error: "Sheet not found" });

    // Find company column in row 1
    let companyCol = null;
    sheet.getRow(1).eachCell((cell, colNum) => {
      if (
        cell.value &&
        String(cell.value).trim().toLowerCase() ===
          String(company).trim().toLowerCase()
      ) {
        companyCol = colNum;
      }
    });

    if (!companyCol) return res.json({ quantities: {} });

    // Read vegetable list from column B
    const quantities = {};
    let row = 2;

    while (true) {
      const veg = sheet.getRow(row).getCell(2).value;
      if (!veg || veg === "") break;

      const cellVal = sheet.getRow(row).getCell(companyCol).value;

      let finalVal = "";
      if (cellVal && typeof cellVal === "object" && "result" in cellVal) {
        finalVal = cellVal.result;
      } else {
        finalVal = cellVal;
      }

      quantities[String(veg).trim()] = finalVal || "";
      row++;
    }

    return res.json({ quantities });
  } catch (err) {
    console.error("/api/get-company ERR:", err);
    res.status(500).json({ error: err.message });
  }
});

/* ---------------------------------------------------
   3️⃣ UPDATE COMPANY VALUES (WITHOUT BREAKING FORMULAS)
   POST: { date, company, quantities }
--------------------------------------------------- */
app.post("/api/update-company", async (req, res) => {
  try {
    const { date, company, quantities } = req.body;

    if (!date || !company || !quantities)
      return res.status(400).json({ error: "Missing fields" });

    const sheetName = String(date).substring(0, 31);
    const workbook = await loadWorkbook();
    const sheet = workbook.getWorksheet(sheetName);

    if (!sheet) return res.status(404).json({ error: "Sheet not found" });

    // Find company column
    let companyCol = null;
    sheet.getRow(1).eachCell((cell, colNum) => {
      if (
        cell.value &&
        String(cell.value).trim().toLowerCase() ===
          String(company).trim().toLowerCase()
      ) {
        companyCol = colNum;
      }
    });

    if (!companyCol)
      return res
        .status(400)
        .json({ error: `Company "${company}" not found in header row` });

    // Build map vegName → row number
    const vegRowMap = {};
    let r = 2;
    while (true) {
      const veg = sheet.getRow(r).getCell(2).value;
      if (!veg || veg === "") break;
      vegRowMap[String(veg).toLowerCase()] = r;
      r++;
    }

    const updates = [];

    for (const [vegName, qty] of Object.entries(quantities)) {
      const key = String(vegName).trim().toLowerCase();
      const rowIndex = vegRowMap[key];

      if (!rowIndex) {
        updates.push({ veg: vegName, status: "not-found" });
        continue;
      }

      const cell = sheet.getRow(rowIndex).getCell(companyCol);

      const numericValue =
        qty === "" || qty === null || qty === undefined
          ? null
          : Number(qty);

      // If formula exists → store new number as cached result
      if (cell.value && typeof cell.value === "object" && cell.value.formula) {
        cell.value = { formula: cell.value.formula, result: numericValue };
      } else {
        cell.value = numericValue;
      }

      sheet.getRow(rowIndex).commit();
      updates.push({ veg: vegName, status: "updated" });
    }

    workbook.calcProperties.fullCalcOnLoad = true;
    await workbook.xlsx.writeFile(EXCEL_FILE_PATH);

    return res.json({ updated: updates });
  } catch (err) {
    console.error("/api/update-company ERR:", err);
    res.status(500).json({ error: err.message });
  }
});

/* ---------------------------------------------------
   4️⃣ DOWNLOAD FILE
--------------------------------------------------- */
app.get("/download", (req, res) => {
  if (!fs.existsSync(EXCEL_FILE_PATH)) {
    return res.status(404).send("Excel file not found");
  }

  res.download(EXCEL_FILE_PATH);
});

/* ---------------------------------------------------
   START SERVER
--------------------------------------------------- */
const PORT = 5000;
app.listen(PORT, () => {
  console.log(`SERVER RUNNING ON PORT ${PORT}`);
  console.log(`Using file: ${EXCEL_FILE_PATH}`);
});
