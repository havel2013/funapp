import React, { useRef, useState, useEffect } from "react";
import { Workbook } from "@fortune-sheet/react";
import "@fortune-sheet/react/dist/index.css";
import * as XLSX from "xlsx";

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
  const [version, setVersion] = useState(0); // Slouží k vynucení překreslení tabulky
  const [status, setStatus] = useState("Načítám data...");

  // 1. NAČTENÍ sdílené tabulky ze serveru (bez cache)
  const loadFromServer = () => {
    setStatus("Načítám nejnovější data ze serveru...");
    fetch(`/.netlify/functions/tabulka?t=${Date.now()}`, {
      cache: "no-store",
    })
      .then((res) => {
        if (!res.ok) throw new Error("Server neodpovídá");
        return res.json();
      })
      .then((savedData) => {
        if (savedData && Array.isArray(savedData) && savedData.length > 0) {
          setSheets(savedData);
          setVersion((v) => v + 1); // Vynutí překreslení tabulky s novými daty
          setStatus("Načtena nejnovější verze ze serveru");
        } else {
          setStatus("Zobrazena výchozí tabulka (na serveru zatím nic není)");
        }
      })
      .catch(() => {
        setStatus("Lokální režim (ukládání vyžaduje Netlify)");
      });
  };

  useEffect(() => {
    loadFromServer();
  }, []);

  // 2. ULOŽENÍ aktuální tabulky na Netlify
  const handleSaveToServer = async () => {
    if (!workbookRef.current) return;
    setStatus("Ukládám...");

    const currentSheets = workbookRef.current.getAllSheets();

    const formattedSheets = currentSheets.map((sheet) => {
      const celldata = [];
      (sheet.data || []).forEach((row, r) => {
        if (!row) return;
        row.forEach((cell, c) => {
          // Uložíme buňku pouze tehdy, pokud má hodnotu nebo vzorec
          if (cell && (cell.v !== undefined && cell.v !== null && cell.v !== "" || cell.f)) {
            celldata.push({ r, c, v: cell });
          }
        });
      });

      return {
        name: sheet.name || "List 1",
        celldata,
      };
    });

    try {
      const res = await fetch("/.netlify/functions/tabulka", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(formattedSheets),
      });

      if (res.ok) {
        setStatus("Uloženo! Změny jsou nyní na serveru pro všechny.");
      } else {
        const errData = await res.json().catch(() => ({}));
        setStatus(`Chyba serveru (${res.status}): ${errData.error || "Neznámá chyba"}`);
      }
    } catch {
      setStatus("Nelze se připojit k serveru Netlify.");
    }
  };

  // 3. NAHRÁNÍ souboru z počítače (.xlsx)
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
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
            if (val !== undefined && val !== null && val !== "") {
              celldata.push({ r, c, v: { v: val } });
            }
          });
        });

        return { name: sheetName, celldata };
      });

      setSheets(importedSheets);
      setVersion((v) => v + 1);
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
        (row || []).map((cell) => (cell ? cell.v : ""))
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

        <button
          onClick={loadFromServer}
          style={{ backgroundColor: "#3b82f6", color: "white", padding: "8px 16px", border: "none", borderRadius: "4px", cursor: "pointer" }}
        >
          Načíst nejnovější ze serveru
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
          key={version}
        />
      </div>
    </div>
  );
}
