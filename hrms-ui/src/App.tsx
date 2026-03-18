import {
  useState,
  useRef,
  useCallback,
  type DragEvent,
  type ChangeEvent,
} from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
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
  "Bonus/Security",
  "PT",
  "TDS",
  "Total  Payble",
  "Branch",
  "Bank A/c No.",
  "IFSC CODE",
];

const SALARY_SHEET_KEYWORDS = ["salary", "sal", "payroll", "wages"];
const MAX_FILE_SIZE_MB = 5;

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
  const [, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [sheetData, setSheetData] = useState<SheetRow[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [errors, setErrors] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [success, setSuccess] = useState("");
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

  const validateAndParseFile = useCallback(
    (selectedFile: File) => {
      setErrors([]);
      setSheetData([]);
      setColumns([]);
      setSuccess("");
      setWorkbook(null);
      setSheetNames([]);

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
    setSheetData([]);
    setColumns([]);
    setErrors([]);
    setSuccess("");
    setProgress("");
  }, []);

  const handleSubmit = useCallback(async () => {
    if (!file) return;

    setLoading(true);
    setErrors([]);
    setSuccess("");
    setProgress("Sending to server...");

    try {
      const formData = new FormData();
      formData.append("file", file);
      // formData.append("sheet", selectedSheet);
      // formData.append("month", month);
      // formData.append("year", year);

      const response = await fetch(
  "http://localhost:5678/webhook-test/generate-salary-slip",
  {
    method: "POST",
    // Send the file object directly as the body
    body: file, 
    headers: {
      "SalarySlipGenerate": "ztshrssg",
      // Manually set the content type to match your Excel file
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  }
);

      if (!response.ok) {
        throw new Error(`Server error: ${response.status}`);
      }

      setProgress("Downloading ZIP...");
      const zipBlob = await response.blob();
      saveAs(zipBlob, `salary-slips.zip`);

      setSuccess("Salary slips generated and downloaded!");

      // Reset form after successful download
      setFile(null);
      setWorkbook(null);
      setSheetNames([]);  
      setSheetData([]);
      setColumns([]);
    } catch (err) {
      console.error(err);
      setErrors(["Failed to generate salary slips. Please try again."]);
    } finally {
      setLoading(false);
      setProgress("");
    }
  }, [file]);

  const canSubmit = file && errors.length === 0 && !loading;

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
