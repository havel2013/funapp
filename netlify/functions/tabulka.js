import { getStore } from "@netlify/blobs";

export default async (req) => {
  // Vytvoří nebo otevře úložiště s názvem "moje-excel-data"
  const store = getStore("moje-excel-data");

  // 1. NAČTENÍ dat (když uživatel otevře stránku)
  if (req.method === "GET") {
    const data = await store.get("hlavni-tabulka", { type: "json" });
    return new Response(JSON.stringify(data || null), {
      headers: { "Content-Type": "application/json" },
    });
  }

  // 2. ULOŽENÍ dat (když uživatel klikne na Uložit)
  if (req.method === "POST") {
    const body = await req.json();
    await store.setJSON("hlavni-tabulka", body);
    return new Response(JSON.stringify({ ok: true }), {
      headers: { "Content-Type": "application/json" },
    });
  }

  return new Response("Metoda není povolena", { status: 405 });
};
