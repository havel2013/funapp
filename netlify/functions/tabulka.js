import React, { useRef, useState, useEffect } from "react";
import { Workbook } from "@fortune-sheet/react";
import "@fortune-sheet/react/dist/index.css";
import * as XLSX from "xlsx";

// Výchozí tabulka (použije se, pokud na serveru ještě nic není uloženo)
const VYCHOZI_DATA = [
  {
    name: "List 1",
    celldata: [
      { r: 0, c: 0, v: { v: "Položka", bl: 1 } },
      { r: 0, c: 1, v: { v: "Cena", bl: 1 } },
      { r: 1, c: 0, v: { v: "Káva" } },
      { r: 1, c: 1, v: { v: 65 } },
      { r: 2, c: 0, v: { v: "Celkem", bl: 1 } },
      { r: 2, c: 1, v: { f: "=SUM(B2:B2)" } },
    ],
  },
];

export default function ExcelEditor() {
  const workbookRef = useRef(null);
  const [sheets, setSheets] = useState(VYCHOZI_DATA);
  const [status, setStatus] = useState("Načítám data...");

  // 1. NAČTENÍ sdílené tabulky ze serveru při otevření stránky
  useEffect(() => {
    fetch("/.netlify/functions/tabulka")
      .then((res) => {
        if (!res.ok) throw new Error("Server neodpovídá");
        return res.json();
      })
      .then((savedData) => {
        if (savedData && savedData.length > 0) {
          setSheets(savedData);
          setStatus("Načtena nejnovější verze ze serveru");
        } else {
          setStatus("Zobrazena výchozí tabulka");
        }
      })
      .catch(() => {
        // Pokud běžíte na GitHub Pages nebo lokálně bez Netlify, použije se výchozí tabulka
        setStatus("Lokální režim (ukládání na server vyžaduje Netlify)");
      });
  }, []);

  // 2. ULOŽENÍ aktuální tabulky na Netlify pro ostatní uživatele
  const handleSaveToServer = async () => {
    if (!workbookRef.current) return;
    setStatus("Ukládám...");

    // Získáme aktuální stav všech listů a převedeme mřížku zpět na celldata
    const currentSheets = workbookRef.current.getAllSheets();
    const formattedSheets = currentSheets.map((sheet) => {
      const celldata = [];
      (sheet.data || []).forEach((row, r) => {
        row.forEach((cell, c) => {
          if (cell !== null && cell !== undefined) {
            celldata.push({ r, c, v: cell });
          }
        });
      });
      return { name: sheet.name, celldata };
    });

    try {
      const res = await fetch("/.netlify/functions/tabulka", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(formattedSheets),
      });

      if (res.ok) {
        setStatus("Uloženo! Ostatní uživatelé nyní uvidí vaše změny.");
      } else {
        setStatus("Chyba při ukládání na server.");
      }
    } catch {
      setStatus("Nelze se připojit k serveru Netlify.");
    }
  };

  // 3. NAHRÁNÍ souboru z počítače (.xlsx)
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });

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
      setStatus("Soubor nahrán (pro sdílení klikněte na Uložit pro všechny)");
    };
    reader.readAsArrayBuffer(file);
  };

  // 4. STAŽENÍ jako Excel (.xlsx)
  const handleDownload = () => {
    if (!workbookRef.current) return;
    const currentSheets = workbookRef.current.getAllSheets();
    const newWorkbook = XLSX.utils.book_new();

    currentSheets.forEach((sheet) => {
      const matrix = (sheet.data || []).map((row) =>
        row.map((cell) => (cell ? cell.v : ""))
      );
      const worksheet = XLSX.utils.aoa_to_sheet(matrix);
      XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheet.name);
    });

    XLSX.writeFile(newWorkbook, "sdilena_tabulka.xlsx");
  };

  return (
    <div style={{ padding: "20px" }}>
      <div style={{ marginBottom: "10px", display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "center" }}>
        <button
          onClick={handleSaveToServer}
          style={{ backgroundColor: "#10b981", color: "white", padding: "8px 16px", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "bold" }}
        >
          Uložit změny pro všechny
        </button>

        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
        <button onClick={handleDownload}>Stáhnout jako Excel (.xlsx)</button>
      </div>

      <div style={{ marginBottom: "10px", fontSize: "14px", color: "#555" }}>
        Stav: <strong>{status}</strong>
      </div>

      <div style={{ width: "100%", height: "500px", border: "1px solid #ccc" }}>
        <Workbook
          ref={workbookRef}
          data={sheets}
          key={JSON.stringify(sheets)}
        />
      </div>
    </div>
  );
}
