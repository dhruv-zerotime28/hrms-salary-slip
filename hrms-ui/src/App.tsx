import {
  useState,
  useRef,
  useCallback,
  type DragEvent,
  type ChangeEvent,
} from "react";
import * as XLSX from "xlsx";
import { pdf } from "@react-pdf/renderer";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { SalarySlipDocument, type EmployeeData } from "./SalarySlipPdf";
import { numberToWords } from "./numberToWords";
import "./App.css";

const REQUIRED_COLUMNS = [
  "No.",
  "Sr. No.",
  "NAME",
  "FATHER'S NAME",
  "DESIGNATION",
  "STAFF/WORKER",
  "DOB",
  "D.O.J.",
  "BASIC+DA",
  "HRA",
  "LTA",
  "ALLOWANCE",
  "GROSS TOTAL",
  "Payable Days",
  "T.Days",
  "Basic+DA(calc)",
  "HRA(calc)",
  "LTA(calc)",
  "Allowance(calc)",
  "Reb. (OT/ Pending/PL)",
  "Gross",
  "Retention Bonus",
  "PT",
  "TDS",
  "Total  Payble",
  "Branch",
  "Bank A/c No.",
  "IFSC CODE",
];

const SALARY_SHEET_KEYWORDS = ["salary", "sal", "payroll", "wages"];
const MAX_FILE_SIZE_MB = 5;
const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

type SheetRow = Record<string, string | number | undefined>;

function formatBytes(bytes: number): string {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(2) + " MB";
}

function normalizeCol(s: string): string {
  return s.replace(/\s+/g, "").toLowerCase();
}

function App() {
  const [file, setFile] = useState<File | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [sheetData, setSheetData] = useState<SheetRow[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [errors, setErrors] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [success, setSuccess] = useState("");
  const [month, setMonth] = useState("");
  const [year, setYear] = useState(new Date().getFullYear().toString());
  const fileInputRef = useRef<HTMLInputElement>(null);

  const validateSheet = useCallback((wb: XLSX.WorkBook, sheetName: string) => {
    const validationErrors: string[] = [];
    const worksheet = wb.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json<SheetRow>(worksheet);

    if (jsonData.length === 0) {
      validationErrors.push(`Sheet "${sheetName}" has no data rows.`);
      setSheetData([]);
      setColumns([]);
      setErrors(validationErrors);
      return;
    }

    const fileColumns = Object.keys(jsonData[0]);
    setColumns(fileColumns);

    const missingCols = REQUIRED_COLUMNS.filter(
      (col) =>
        !fileColumns.some((fc) => normalizeCol(fc) === normalizeCol(col)),
    );
    if (missingCols.length > 0) {
      validationErrors.push(
        `Missing required columns: ${missingCols.join(", ")}`,
      );
    }

    const emptyFieldRows: string[] = [];
    jsonData.forEach((row, idx) => {
      REQUIRED_COLUMNS.forEach((col) => {
        const matchedKey = fileColumns.find(
          (fc) => normalizeCol(fc) === normalizeCol(col),
        );
        if (
          matchedKey &&
          (row[matchedKey] === undefined || row[matchedKey] === "")
        ) {
          emptyFieldRows.push(`Row ${idx + 2}: "${col}" is empty`);
        }
      });
    });
    if (emptyFieldRows.length > 0) {
      validationErrors.push(...emptyFieldRows.slice(0, 10));
      if (emptyFieldRows.length > 10) {
        validationErrors.push(
          `...and ${emptyFieldRows.length - 10} more empty field issues`,
        );
      }
    }

    setSheetData(jsonData.slice(0, 5));
    setErrors(validationErrors);
  }, []);

  const handleSheetChange = useCallback(
    (sheetName: string) => {
      setSelectedSheet(sheetName);
      setErrors([]);
      setSheetData([]);
      setColumns([]);
      setSuccess("");
      if (workbook && sheetName) {
        validateSheet(workbook, sheetName);
      }
    },
    [workbook, validateSheet],
  );

  const validateAndParseFile = useCallback(
    (selectedFile: File) => {
      setErrors([]);
      setSheetData([]);
      setColumns([]);
      setSuccess("");
      setWorkbook(null);
      setSheetNames([]);
      setSelectedSheet("");

      const validationErrors: string[] = [];

      const ext = selectedFile.name.split(".").pop()?.toLowerCase();
      if (ext !== "xlsx" && ext !== "xls") {
        validationErrors.push(
          "Invalid file type. Only .xlsx and .xls files are accepted.",
        );
        setErrors(validationErrors);
        return;
      }

      if (selectedFile.size > MAX_FILE_SIZE_MB * 1048576) {
        validationErrors.push(
          `File size exceeds ${MAX_FILE_SIZE_MB} MB limit.`,
        );
        setErrors(validationErrors);
        return;
      }

      setFile(selectedFile);

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: "array" });
          setWorkbook(wb);
          setSheetNames(wb.SheetNames);

          const salarySheet = wb.SheetNames.find((name) =>
            SALARY_SHEET_KEYWORDS.some((kw) => name.toLowerCase().includes(kw)),
          );
          const autoSheet = salarySheet || wb.SheetNames[0];
          setSelectedSheet(autoSheet);
          validateSheet(wb, autoSheet);
        } catch {
          setErrors([
            "Failed to parse the Excel file. Please ensure it is a valid .xlsx/.xls file.",
          ]);
          setFile(null);
        }
      };
      reader.readAsArrayBuffer(selectedFile);
    },
    [validateSheet],
  );

  const handleDrop = useCallback(
    (e: DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setDragging(false);
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile) validateAndParseFile(droppedFile);
    },
    [validateAndParseFile],
  );

  const handleFileSelect = useCallback(
    (e: ChangeEvent<HTMLInputElement>) => {
      const selected = e.target.files?.[0];
      if (selected) validateAndParseFile(selected);
      e.target.value = "";
    },
    [validateAndParseFile],
  );

  const removeFile = useCallback(() => {
    setFile(null);
    setWorkbook(null);
    setSheetNames([]);
    setSelectedSheet("");
    setSheetData([]);
    setColumns([]);
    setErrors([]);
    setSuccess("");
    setProgress("");
  }, []);

  const handleSubmit = useCallback(async () => {
    if (!workbook || !selectedSheet || !month || !year) return;

    setLoading(true);
    setErrors([]);
    setSuccess("");
    setProgress("Reading Excel data...");

    try {
      const worksheet = workbook.Sheets[selectedSheet];
      const allData = XLSX.utils.sheet_to_json<SheetRow>(worksheet);
      const cols = Object.keys(allData[0] || {});

      const getField = (row: SheetRow, fieldName: string): string => {
        const norm = normalizeCol(fieldName);
        const key = cols.find((c) => normalizeCol(c) === norm);
        return key && row[key] !== undefined && row[key] !== ""
          ? String(row[key])
          : "-";
      };

      function formatExcelDate(value: string | number | undefined): string {
        if (!value) return "-";
        // If it's a number, treat as Excel serial date
        if (typeof value === "number" && !isNaN(value)) {
          let serial = value;
          if (serial >= 60) serial -= 1;
          const excelEpoch = new Date(Date.UTC(1899, 11, 30));
          const d = new Date(excelEpoch.getTime() + serial * 86400000);
          return (
            d.getUTCDate().toString().padStart(2, "0") +
            "-" +
            d.toLocaleString("en-IN", { month: "short" }) +
            "-" +
            d.getUTCFullYear()
          );
        }
        // If it's a string that is a number, treat as Excel serial date
        if (typeof value === "string" && /^\d+$/.test(value)) {
          let serial = parseInt(value, 10);
          if (serial >= 60) serial -= 1;
          const excelEpoch = new Date(Date.UTC(1899, 11, 30));
          const d = new Date(excelEpoch.getTime() + serial * 86400000);
          return (
            d.getUTCDate().toString().padStart(2, "0") +
            "-" +
            d.toLocaleString("en-IN", { month: "short" }) +
            "-" +
            d.getUTCFullYear()
          );
        }
        // If it's already a string date, try to parse and format
        if (typeof value === "string") {
          // Try parse as ISO or dd/mm/yyyy
          let d = new Date(value);
          if (
            isNaN(d.getTime()) &&
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/.test(value)
          ) {
            const [day, month, year] = value.split(/[\/\-]/);
            d = new Date(
              Number(year.length === 2 ? "20" + year : year),
              Number(month) - 1,
              Number(day),
            );
          }
          if (!isNaN(d.getTime())) {
            return (
              d.getDate().toString().padStart(2, "0") +
              "-" +
              d.toLocaleString("en-IN", { month: "short" }) +
              "-" +
              d.getFullYear()
            );
          }
          return value;
        }
        return String(value);
      }

      const employees: EmployeeData[] = allData.map((row) => {
        const totalDays = parseFloat(getField(row, "T.Days")) || 0;
        const payableDays = parseFloat(getField(row, "Payable Days")) || 0;
        const pt = parseFloat(getField(row, "PT")) || 0;
        const tds = parseFloat(getField(row, "TDS")) || 0;
        const pf = parseFloat(getField(row, "PF")) || 0;
        const esic = parseFloat(getField(row, "ESIC")) || 0;
        const bonus = parseFloat(getField(row, "Retention Bonus")) || 0;
        const netPay = parseFloat(getField(row, "Total  Payble")) || 0;

        return {
          name: getField(row, "NAME"),
          panNumber: getField(row, "PAN"),
          designation: getField(row, "DESIGNATION"),
          doj: formatExcelDate(getField(row, "D.O.J.")),
          employeeId: getField(row, "Sr. No."),
          totalDays: String(totalDays),
          payableDays: String(payableDays),
          absentDays: String(totalDays - payableDays),
          basicDA: getField(row, "Basic+DA(calc)"),
          hra: getField(row, "HRA(calc)"),
          lta: getField(row, "LTA(calc)"),
          allowance: getField(row, "Allowance(calc)"),
          grossSalary: getField(row, "GROSS TOTAL"),
          other: getField(row, "Reb. (OT/ Pending/PL)"),
          RetentionBonus: getField(row, "Retention Bonus"),
          pt: getField(row, "PT"),
          tds: getField(row, "TDS"),
          pf: getField(row, "PF"),
          esic: getField(row, "ESIC"),
          earningTotal: getField(row, "Gross"),
          deductionTotal: String(pt + tds + bonus + pf + esic),
          netPay: getField(row, "Total  Payble"),
          netPayInWords: numberToWords(netPay),
          bankAc: getField(row, "Bank A/c No."),
          ifsc: getField(row, "IFSC CODE"),
          branch: getField(row, "Branch"),
        };
      });

      const zip = new JSZip();

      for (let i = 0; i < employees.length; i++) {
        const emp = employees[i];
        setProgress(
          `Generating PDF ${i + 1} of ${employees.length} — ${emp.name}`,
        );
        // Yield to UI so progress renders
        await new Promise((r) => setTimeout(r, 0));

        const logoSrc = `${window.location.origin}/zerotime_logo_full.png`;
        const blob = await pdf(
          <SalarySlipDocument
            data={emp}
            month={month}
            year={year}
            logoSrc={logoSrc}
          />,
        ).toBlob();

        const safeName =
          emp.name
            .replace(/[^a-zA-Z0-9 ]/g, "")
            .trim()
            .replace(/\s+/g, "_") || `employee_${i + 1}`;
        zip.file(`${safeName}_salary_slip.pdf`, blob);
      }

      setProgress("Creating ZIP file...");
      await new Promise((r) => setTimeout(r, 0));

      const zipBlob = await zip.generateAsync({ type: "blob" });
      saveAs(zipBlob, `salary-slips-${month}-${year}.zip`);

      setSuccess(
        `Generated ${employees.length} salary slips and downloaded as ZIP!`,
      );

      // Reset form after successful download
      setFile(null);
      setWorkbook(null);
      setSheetNames([]);
      setSelectedSheet("");
      setSheetData([]);
      setColumns([]);
    } catch (err) {
      console.error(err);
      setErrors(["Failed to generate salary slips. Please try again."]);
    } finally {
      setLoading(false);
      setProgress("");
    }
  }, [workbook, selectedSheet, month, year]);

  const canSubmit =
    file && selectedSheet && month && year && errors.length === 0 && !loading;

  return (
    <div className="app-container">
      <header className="app-header">
        <h1 className="com-title">ZEROTIME SOLUTIONS</h1>
        <h1>Salary Slip Generator</h1>
        <p>Upload employee salary data to generate PDF salary slips</p>
      </header>

      {/* Drop Zone */}
      <div
        className={`drop-zone ${dragging ? "dragging" : ""} ${file ? "has-file" : ""}`}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          style={{ display: "none" }}
          onChange={handleFileSelect}
        />
        {!file ? (
          <>
            <span className="drop-zone-icon">📄</span>
            <div className="drop-zone-text">
              Drop your Excel file here or <strong>click to browse</strong>
            </div>
            <div className="drop-zone-hint">
              Accepts .xlsx and .xls files up to {MAX_FILE_SIZE_MB} MB
            </div>
          </>
        ) : (
          <>
            <span className="drop-zone-icon">✅</span>
            <div className="drop-zone-text">File loaded — click to replace</div>
          </>
        )}
      </div>

      {/* File Info */}
      {file && (
        <div className="file-info">
          <div className="file-info-details">
            <span className="file-info-icon">📊</span>
            <div>
              <div className="file-info-name">{file.name}</div>
              <div className="file-info-meta">
                {formatBytes(file.size)} • {sheetNames.length} sheet
                {sheetNames.length !== 1 ? "s" : ""}
              </div>
            </div>
          </div>
          <button
            className="file-remove-btn"
            onClick={removeFile}
            title="Remove file"
          >
            ✕
          </button>
        </div>
      )}

      {/* Sheet Selector */}
      {sheetNames.length > 1 && (
        <div className="sheet-selector">
          <label htmlFor="sheet-select">Select Salary Sheet</label>
          <select
            id="sheet-select"
            value={selectedSheet}
            onChange={(e) => handleSheetChange(e.target.value)}
          >
            {sheetNames.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
          {selectedSheet && (
            <span className="sheet-badge">
              Selected: <strong>{selectedSheet}</strong>
            </span>
          )}
        </div>
      )}

      {sheetNames.length === 1 && selectedSheet && (
        <div className="sheet-selector">
          <span className="sheet-badge">
            Sheet: <strong>{selectedSheet}</strong>
          </span>
        </div>
      )}

      {/* Month & Year */}
      {file && (
        <div className="month-year-section">
          <div className="month-year-field">
            <label htmlFor="month-select">Month</label>
            <select
              id="month-select"
              value={month}
              onChange={(e) => setMonth(e.target.value)}
            >
              <option value="">Select month</option>
              {MONTHS.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          </div>
          <div className="month-year-field">
            <label htmlFor="year-input">Year</label>
            <input
              id="year-input"
              type="number"
              min="2020"
              max="2030"
              value={year}
              onChange={(e) => setYear(e.target.value)}
            />
          </div>
        </div>
      )}

      {/* Validation Errors */}
      {errors.length > 0 && (
        <div className="error-box">
          <strong>Validation Issues</strong>
          <ul>
            {errors.map((err, i) => (
              <li key={i}>{err}</li>
            ))}
          </ul>
        </div>
      )}

      {/* Data Preview */}
      {sheetData.length > 0 && columns.length > 0 && (
        <div className="validation-section">
          <h3>Data Preview (first {sheetData.length} rows)</h3>
          <div className="validation-table-wrapper">
            <table className="validation-table">
              <thead>
                <tr>
                  {columns.map((col) => (
                    <th key={col}>{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sheetData.map((row, i) => (
                  <tr key={i}>
                    {columns.map((col) => (
                      <td key={col}>
                        {row[col] !== undefined ? String(row[col]) : ""}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Success */}
      {success && <div className="success-box">{success}</div>}

      {/* Submit */}
      <div className="submit-section">
        <button
          className="submit-btn"
          disabled={!canSubmit}
          onClick={handleSubmit}
        >
          {loading && <span className="spinner" />}
          {loading
            ? progress || "Generating..."
            : "Generate & Download Salary Slips"}
        </button>
      </div>
    </div>
  );
}

export default App;
