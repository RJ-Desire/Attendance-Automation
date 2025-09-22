const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");

const app = express();
const upload = multer();

function getShiftCode(timeStr) {
  if (!timeStr) return "";
  let clean = timeStr.replace(/[:\s]/g, "").slice(0, 2);
  let hour = parseInt(clean, 10);
  if (isNaN(hour)) return "";
  if (hour >= 4 && hour < 6) return "FS";
  if (hour >= 6 && hour < 10) return "GS";
  if (hour >= 11 && hour < 16) return "SS";
  if (hour >= 16) return "NS";
  return "";
}

function normalize(s) {
  return (s || "").toString().toLowerCase().replace(/\s/g, "");
}

function findColIndex(headerRow, variants) {
  for (let i = 1; i <= headerRow.cellCount; i++) {
    let val = normalize(headerRow.getCell(i).value);
    for (let v of variants) {
      if (val.includes(v)) return i;
    }
  }
  return -1;
}

app.post(
  "/process",
  upload.fields([
    { name: "logReport", maxCount: 1 },
    { name: "deptSheets", maxCount: 5 },
  ]),
  async (req, res) => {
    try {
      const logFile = req.files["logReport"][0];
      const deptFiles = req.files["deptSheets"];

      const logWorkbook = new ExcelJS.Workbook();
      await logWorkbook.xlsx.load(logFile.buffer);
      const logSheet = logWorkbook.worksheets[0];

      // Find log report columns
      const empIdCol = findColIndex(logSheet.getRow(1), [
        "employeeid",
        "employeesid",
        "empid",
        "emp",
      ]);
      const dateCol = findColIndex(logSheet.getRow(1), ["logdate", "date"]);
      const timeCol = findColIndex(logSheet.getRow(1), [
        "time",
        "in",
        "punchin",
      ]);

      if (empIdCol < 1 || dateCol < 1 || timeCol < 1) {
        return res.status(400).send("Missing required columns in log report");
      }

      // Build punch map
      let punchMap = {};
      logSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const empId = row.getCell(empIdCol).text;
        const dateVal = row.getCell(dateCol).text || row.getCell(dateCol).value;
        const timeVal = row.getCell(timeCol).text;
        if (!empId || !dateVal || !timeVal) return;

        let dateKey =
          dateVal instanceof Date
            ? dateVal.toLocaleDateString()
            : dateVal.toString();
        let key = empId + "|" + dateKey;

        if (!punchMap[key] || punchMap[key] > timeVal) {
          punchMap[key] = timeVal;
        }
      });

      let updatedFiles = [];
      for (let df of deptFiles) {
        let workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(df.buffer);

        let sheet = workbook.worksheets[0];
        const headerRow = sheet.getRow(1);

        const deptEmpIdCol = findColIndex(headerRow, [
          "employeeid",
          "employeesid",
          "empid",
          "emp",
        ]);

        let dateCols = [];
        headerRow.eachCell((cell, colNumber) => {
          if (
            cell.type === ExcelJS.ValueType.String &&
            /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(cell.value)
          ) {
            dateCols.push(colNumber);
          }
        });

        sheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;
          let empId = row.getCell(deptEmpIdCol).text;
          if (!empId) return;
          dateCols.forEach((dc) => {
            let dateHdr = headerRow.getCell(dc).value;
            let dateKey =
              dateHdr instanceof Date
                ? dateHdr.toLocaleDateString()
                : dateHdr.toString();
            let key = empId + "|" + dateKey;
            if (punchMap[key]) {
              let code = getShiftCode(punchMap[key]);
              if (code) {
                row.getCell(dc).value = code;
              }
            }
          });
        });

        let buffer = await workbook.xlsx.writeBuffer();
        updatedFiles.push({
          filename: "updated_" + df.originalname,
          buffer,
        });
      }

      // Send first updated file for demo
      res.set({
        "Content-Disposition": `attachment; filename="${updatedFiles[0].filename}"`,
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      return res.send(updatedFiles[0].buffer);
    } catch (error) {
      console.error(error);
      res.status(500).send("Error processing files: " + error.message);
    }
  },
);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
