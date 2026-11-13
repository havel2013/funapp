import React, { useRef, useState } from "react";
import { Workbook } from "@fortune-sheet/react";
import "@fortune-sheet/react/dist/index.css";
import * as XLSX from "xlsx";

export default function ExcelEditor() {
  const workbookRef = useRef(null);

  // 1. Výchozí stav tabulky při otevření stránky
  const [sheets, setSheets] = useState([
    {
      name: "List 1",
      celldata: [
        { r: 0, c: 0, v: { v: "Položka", bl: 1 } },
        { r: 0, c: 1, v: { v: "Cena", bl: 1 } },
        { r: 1, c: 0, v: { v: "Káva" } },
        { r: 1, c: 1, v: { v: 65 } },
        { r: 2, c: 0, v: { v: "Dort" } },
        { r: 2, c: 1, v: { v: 95 } },
        { r: 3, c: 0, v: { v: "Celkem", bl: 1 } },
        { r: 3, c: 1, v: { f: "=SUM(B2:B3)" } }, // Ukázka vzorce
      ],
    },
  ]);

  // 2. NAHRÁNÍ souboru z počítače (.xlsx) do tabulky
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Převedeme všechny listy z Excelu do formátu pro FortuneSheet
      const importedSheets = workbook.SheetNames.map((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const celldata = [];
        rows.forEach((row, r) => {
          row.forEach((val, c) => {
            if (val !== undefined && val !== null) {
              celldata.push({ r, c, v: { v: val } });
            }
          });
        });

        return { name: sheetName, celldata };
      });

      setSheets(importedSheets);
    };
    reader.readAsArrayBuffer(file);
  };

  // 3. STAŽENÍ upravené tabulky zpět jako Excel (.xlsx)
  const handleDownload = () => {
    if (!workbookRef.current) return;

    // Získáme aktuální data ze všech listů ve FortuneSheet
    const currentSheets = workbookRef.current.getAllSheets();
    const newWorkbook = XLSX.utils.book_new();

    currentSheets.forEach((sheet) => {
      // Převedeme mřížku buněk zpět na klasické řádky a sloupce
      const matrix = (sheet.data || []).map((row) =>
        row.map((cell) => (cell ? cell.v : ""))
      );
      const worksheet = XLSX.utils.aoa_to_sheet(matrix);
      XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheet.name);
    });

    // Vyvolá stažení souboru v prohlížeči
    XLSX.writeFile(newWorkbook, "upravena_tabulka.xlsx");
  };

  return (
    <div style={{ padding: "20px" }}>
      <div style={{ marginBottom: "15px", display: "flex", gap: "10px" }}>
        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
        <button onClick={handleDownload}>Stáhnout jako Excel (.xlsx)</button>
      </div>

      {/* Důležité: Obalující div musí mít nastavenou pevnou výšku */}
      <div style={{ width: "100%", height: "500px", border: "1px solid #ccc" }}>
        <Workbook
          ref={workbookRef}
          data={sheets}
          key={JSON.stringify(sheets)} // Zajistí překreslení při nahrání nového souboru
        />
      </div>
    </div>
  );
}
