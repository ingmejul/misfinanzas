import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

const CATEGORIES = {
  income: [
    { id: "salary", label: "Salario", emoji: "💼" },
    { id: "freelance", label: "Freelance", emoji: "💻" },
    { id: "investment", label: "Inversión", emoji: "📈" },
    { id: "gift", label: "Regalo", emoji: "🎁" },
    { id: "other_in", label: "Otro", emoji: "➕" },
  ],
  expense: [
    { id: "food", label: "Comida", emoji: "🍽️" },
    { id: "transport", label: "Transporte", emoji: "🚗" },
    { id: "health", label: "Salud", emoji: "❤️" },
    { id: "entertainment", label: "Ocio", emoji: "🎬" },
    { id: "bills", label: "Servicios", emoji: "🏠" },
    { id: "shopping", label: "Compras", emoji: "🛍️" },
    { id: "other_ex", label: "Otro", emoji: "➖" },
  ],
};

const formatMXN = (n) =>
  new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN" }).format(n);

const today = () => new Date().toISOString().split("T")[0];

export default function App() {
  const [records, setRecords] = useState([
    { id: 1, type: "income", category: "salary", amount: 18000, note: "Quincena mayo", date: "2026-05-01" },
    { id: 2, type: "expense", category: "food", amount: 850, note: "Super semana", date: "2026-05-05" },
    { id: 3, type: "expense", category: "transport", amount: 420, note: "Gasolina", date: "2026-05-07" },
    { id: 4, type: "income", category: "freelance", amount: 4500, note: "Proyecto web", date: "2026-05-10" },
    { id: 5, type: "expense", category: "entertainment", amount: 350, note: "Cine + cena", date: "2026-05-11" },
  ]);

  const [view, setView] = useState("home"); // home | add | history | export
  const [form, setForm] = useState({ type: "expense", category: "", amount: "", note: "", date: today() });
  const [email, setEmail] = useState("");
  const [toast, setToast] = useState(null);
  const [filterType, setFilterType] = useState("all");

  const totals = useMemo(() => {
    const income = records.filter((r) => r.type === "income").reduce((a, r) => a + r.amount, 0);
    const expense = records.filter((r) => r.type === "expense").reduce((a, r) => a + r.amount, 0);
    return { income, expense, balance: income - expense };
  }, [records]);

  const filtered = useMemo(() =>
    filterType === "all" ? records : records.filter((r) => r.type === filterType),
    [records, filterType]
  );

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  };

  const addRecord = () => {
    if (!form.category || !form.amount || isNaN(parseFloat(form.amount))) {
      showToast("Completa todos los campos", "error");
      return;
    }
    const newRec = { ...form, id: Date.now(), amount: parseFloat(form.amount) };
    setRecords((prev) => [newRec, ...prev]);
    setForm({ type: "expense", category: "", amount: "", note: "", date: today() });
    showToast("Registro agregado ✓");
    setView("home");
  };

  const deleteRecord = (id) => {
    setRecords((prev) => prev.filter((r) => r.id !== id));
    showToast("Registro eliminado");
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();

    // Main sheet
    const rows = records.map((r) => {
      const cats = [...CATEGORIES.income, ...CATEGORIES.expense];
      const cat = cats.find((c) => c.id === r.category);
      return {
        Fecha: r.date,
        Tipo: r.type === "income" ? "Ingreso" : "Gasto",
        Categoría: cat ? `${cat.emoji} ${cat.label}` : r.category,
        Concepto: r.note || "—",
        Monto: r.type === "expense" ? -r.amount : r.amount,
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 10 }, { wch: 18 }, { wch: 26 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws, "Movimientos");

    // Summary sheet
    const summary = [
      { Concepto: "Total Ingresos", Monto: totals.income },
      { Concepto: "Total Gastos", Monto: totals.expense },
      { Concepto: "Balance", Monto: totals.balance },
    ];
    const ws2 = XLSX.utils.json_to_sheet(summary);
    ws2["!cols"] = [{ wch: 22 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, ws2, "Resumen");

    XLSX.writeFile(wb, "MisFinanzas.xlsx");
    return wb;
  };

  const handleExportAndEmail = () => {
    if (!email || !email.includes("@")) {
      showToast("Ingresa un correo válido", "error");
      return;
    }
    exportExcel();
    // Build mailto with CSV body as fallback info
    const subject = encodeURIComponent("📊 Reporte de Finanzas Personales - MisFinanzas");
    const body = encodeURIComponent(
      `Hola,\n\nAdjunto encontrarás el reporte de finanzas personales generado el ${new Date().toLocaleDateString("es-MX")}.\n\nResumen:\n• Ingresos: ${formatMXN(totals.income)}\n• Gastos: ${formatMXN(totals.expense)}\n• Balance: ${formatMXN(totals.balance)}\n\nEl archivo Excel ha sido descargado en tu dispositivo. Adjúntalo manualmente a este correo.\n\nGenerado con MisFinanzas App`
    );
    window.location.href = `mailto:${email}?subject=${subject}&body=${body}`;
    showToast("Archivo descargado · Abriendo correo…");
  };

  const getCatInfo = (type, catId) => {
    const list = type === "income" ? CATEGORIES.income : CATEGORIES.expense;
    return list.find((c) => c.id === catId) || { emoji: "•", label: catId };
  };

  /* ─── SCREENS ─── */

  const HomeScreen = () => (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {/* Balance card */}
      <div style={{
        background: "linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)",
        borderRadius: 28,
        padding: "32px 28px",
        color: "#fff",
        position: "relative",
        overflow: "hidden",
        boxShadow: "0 20px 60px rgba(15,52,96,0.5)",
      }}>
        <div style={{
          position: "absolute", top: -40, right: -40, width: 180, height: 180,
          borderRadius: "50%", background: "rgba(255,255,255,0.04)"
        }} />
        <div style={{
          position: "absolute", bottom: -60, left: -20, width: 220, height: 220,
          borderRadius: "50%", background: "rgba(83,144,217,0.08)"
        }} />
        <p style={{ margin: 0, fontSize: 13, opacity: 0.65, letterSpacing: 2, textTransform: "uppercase", fontFamily: "monospace" }}>
          Balance Actual
        </p>
        <h1 style={{
          margin: "10px 0 24px", fontSize: 42, fontFamily: "'Georgia', serif",
          fontWeight: 700, letterSpacing: -1,
          color: totals.balance >= 0 ? "#7fffd4" : "#ff7b7b"
        }}>
          {formatMXN(totals.balance)}
        </h1>
        <div style={{ display: "flex", gap: 24 }}>
          <div>
            <p style={{ margin: 0, fontSize: 11, opacity: 0.55, letterSpacing: 1.5, textTransform: "uppercase" }}>Ingresos</p>
            <p style={{ margin: "4px 0 0", fontSize: 19, fontWeight: 600, color: "#7fffd4" }}>
              {formatMXN(totals.income)}
            </p>
          </div>
          <div style={{ width: 1, background: "rgba(255,255,255,0.12)" }} />
          <div>
            <p style={{ margin: 0, fontSize: 11, opacity: 0.55, letterSpacing: 1.5, textTransform: "uppercase" }}>Gastos</p>
            <p style={{ margin: "4px 0 0", fontSize: 19, fontWeight: 600, color: "#ff9f9f" }}>
              {formatMXN(totals.expense)}
            </p>
          </div>
        </div>
      </div>

      {/* Quick actions */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        {[
          { label: "Agregar", sub: "Registro", emoji: "＋", action: () => setView("add"), bg: "#0f3460", accent: "#7fffd4" },
          { label: "Historial", sub: "Movimientos", emoji: "📋", action: () => setView("history"), bg: "#1a1a2e", accent: "#a78bfa" },
          { label: "Exportar", sub: "Excel", emoji: "📊", action: () => setView("export"), bg: "#16213e", accent: "#60a5fa" },
          { label: "Resumen", sub: `${records.length} registros`, emoji: "📈", action: () => setView("history"), bg: "#1a1a2e", accent: "#fbbf24" },
        ].map((btn) => (
          <button key={btn.label} onClick={btn.action} style={{
            background: btn.bg,
            border: `1px solid rgba(255,255,255,0.08)`,
            borderRadius: 20,
            padding: "18px 16px",
            cursor: "pointer",
            textAlign: "left",
            transition: "transform 0.15s, box-shadow 0.15s",
            boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
          }}
            onMouseEnter={e => e.currentTarget.style.transform = "translateY(-2px)"}
            onMouseLeave={e => e.currentTarget.style.transform = "translateY(0)"}
          >
            <div style={{ fontSize: 26, marginBottom: 8 }}>{btn.emoji}</div>
            <div style={{ color: btn.accent, fontSize: 16, fontWeight: 700 }}>{btn.label}</div>
            <div style={{ color: "rgba(255,255,255,0.4)", fontSize: 12, marginTop: 2 }}>{btn.sub}</div>
          </button>
        ))}
      </div>

      {/* Recent */}
      <div>
        <p style={{ margin: "0 0 12px", color: "rgba(255,255,255,0.5)", fontSize: 12, letterSpacing: 2, textTransform: "uppercase" }}>
          Recientes
        </p>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {records.slice(0, 4).map((r) => {
            const cat = getCatInfo(r.type, r.category);
            return (
              <div key={r.id} style={{
                background: "rgba(255,255,255,0.04)",
                borderRadius: 16,
                padding: "14px 16px",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                border: "1px solid rgba(255,255,255,0.06)",
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ fontSize: 22 }}>{cat.emoji}</span>
                  <div>
                    <div style={{ color: "#fff", fontSize: 14, fontWeight: 500 }}>{r.note || cat.label}</div>
                    <div style={{ color: "rgba(255,255,255,0.35)", fontSize: 11, marginTop: 2 }}>{r.date}</div>
                  </div>
                </div>
                <span style={{
                  fontSize: 15, fontWeight: 700,
                  color: r.type === "income" ? "#7fffd4" : "#ff9f9f"
                }}>
                  {r.type === "income" ? "+" : "-"}{formatMXN(r.amount)}
                </span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );

  const AddScreen = () => {
    const cats = form.type === "income" ? CATEGORIES.income : CATEGORIES.expense;
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
        <h2 style={{ margin: 0, color: "#fff", fontSize: 26, fontFamily: "'Georgia', serif" }}>
          Nuevo Registro
        </h2>

        {/* Type toggle */}
        <div style={{ display: "flex", background: "rgba(255,255,255,0.06)", borderRadius: 16, padding: 4 }}>
          {["expense", "income"].map((t) => (
            <button key={t} onClick={() => setForm({ ...form, type: t, category: "" })} style={{
              flex: 1, padding: "12px", borderRadius: 13,
              border: "none", cursor: "pointer", fontSize: 14, fontWeight: 600,
              transition: "all 0.2s",
              background: form.type === t
                ? (t === "income" ? "#7fffd4" : "#ff9f9f")
                : "transparent",
              color: form.type === t ? "#0f0f1a" : "rgba(255,255,255,0.5)",
            }}>
              {t === "income" ? "💚 Ingreso" : "❤️ Gasto"}
            </button>
          ))}
        </div>

        {/* Categories */}
        <div>
          <p style={{ margin: "0 0 10px", color: "rgba(255,255,255,0.5)", fontSize: 12, letterSpacing: 1.5, textTransform: "uppercase" }}>
            Categoría
          </p>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {cats.map((c) => (
              <button key={c.id} onClick={() => setForm({ ...form, category: c.id })} style={{
                padding: "8px 14px", borderRadius: 100, border: "1px solid",
                cursor: "pointer", fontSize: 13, transition: "all 0.15s",
                borderColor: form.category === c.id ? (form.type === "income" ? "#7fffd4" : "#ff9f9f") : "rgba(255,255,255,0.12)",
                background: form.category === c.id ? (form.type === "income" ? "rgba(127,255,212,0.15)" : "rgba(255,159,159,0.15)") : "transparent",
                color: form.category === c.id ? "#fff" : "rgba(255,255,255,0.5)",
              }}>
                {c.emoji} {c.label}
              </button>
            ))}
          </div>
        </div>

        {/* Amount */}
        <div>
          <p style={{ margin: "0 0 8px", color: "rgba(255,255,255,0.5)", fontSize: 12, letterSpacing: 1.5, textTransform: "uppercase" }}>
            Monto (MXN)
          </p>
          <div style={{ position: "relative" }}>
            <span style={{ position: "absolute", left: 16, top: "50%", transform: "translateY(-50%)", color: "rgba(255,255,255,0.4)", fontSize: 18 }}>$</span>
            <input
              type="number" placeholder="0.00" value={form.amount}
              onChange={e => setForm({ ...form, amount: e.target.value })}
              style={{
                width: "100%", padding: "16px 16px 16px 36px", borderRadius: 14,
                background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)",
                color: "#fff", fontSize: 22, fontWeight: 700, boxSizing: "border-box",
                outline: "none",
              }}
            />
          </div>
        </div>

        {/* Note */}
        <div>
          <p style={{ margin: "0 0 8px", color: "rgba(255,255,255,0.5)", fontSize: 12, letterSpacing: 1.5, textTransform: "uppercase" }}>
            Concepto
          </p>
          <input
            placeholder="Descripción (opcional)" value={form.note}
            onChange={e => setForm({ ...form, note: e.target.value })}
            style={{
              width: "100%", padding: "14px 16px", borderRadius: 14,
              background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)",
              color: "#fff", fontSize: 15, boxSizing: "border-box", outline: "none",
            }}
          />
        </div>

        {/* Date */}
        <div>
          <p style={{ margin: "0 0 8px", color: "rgba(255,255,255,0.5)", fontSize: 12, letterSpacing: 1.5, textTransform: "uppercase" }}>
            Fecha
          </p>
          <input
            type="date" value={form.date}
            onChange={e => setForm({ ...form, date: e.target.value })}
            style={{
              width: "100%", padding: "14px 16px", borderRadius: 14,
              background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)",
              color: "#fff", fontSize: 15, boxSizing: "border-box", outline: "none",
            }}
          />
        </div>

        <button onClick={addRecord} style={{
          padding: "18px", borderRadius: 16, border: "none", cursor: "pointer",
          background: form.type === "income"
            ? "linear-gradient(135deg, #2dd4bf, #7fffd4)"
            : "linear-gradient(135deg, #f87171, #fca5a5)",
          color: "#0f0f1a", fontSize: 17, fontWeight: 800, letterSpacing: 0.5,
          boxShadow: "0 8px 30px rgba(0,0,0,0.4)",
          transition: "transform 0.15s",
        }}
          onMouseEnter={e => e.currentTarget.style.transform = "scale(1.02)"}
          onMouseLeave={e => e.currentTarget.style.transform = "scale(1)"}
        >
          Guardar Registro
        </button>
      </div>
    );
  };

  const HistoryScreen = () => (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <h2 style={{ margin: 0, color: "#fff", fontSize: 26, fontFamily: "'Georgia', serif" }}>Historial</h2>

      <div style={{ display: "flex", background: "rgba(255,255,255,0.06)", borderRadius: 14, padding: 4 }}>
        {[["all", "Todos"], ["income", "Ingresos"], ["expense", "Gastos"]].map(([v, l]) => (
          <button key={v} onClick={() => setFilterType(v)} style={{
            flex: 1, padding: "10px 4px", borderRadius: 11, border: "none", cursor: "pointer",
            fontSize: 12, fontWeight: 600, transition: "all 0.2s",
            background: filterType === v ? "rgba(255,255,255,0.12)" : "transparent",
            color: filterType === v ? "#fff" : "rgba(255,255,255,0.4)",
          }}>{l}</button>
        ))}
      </div>

      {filtered.length === 0 && (
        <div style={{ textAlign: "center", color: "rgba(255,255,255,0.3)", padding: "40px 0" }}>
          Sin registros
        </div>
      )}

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {filtered.map((r) => {
          const cat = getCatInfo(r.type, r.category);
          return (
            <div key={r.id} style={{
              background: "rgba(255,255,255,0.04)",
              borderRadius: 18, padding: "16px",
              display: "flex", justifyContent: "space-between", alignItems: "center",
              border: "1px solid rgba(255,255,255,0.06)",
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{
                  width: 44, height: 44, borderRadius: 13,
                  background: r.type === "income" ? "rgba(127,255,212,0.1)" : "rgba(255,159,159,0.1)",
                  display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22
                }}>
                  {cat.emoji}
                </div>
                <div>
                  <div style={{ color: "#fff", fontSize: 14, fontWeight: 600 }}>{r.note || cat.label}</div>
                  <div style={{ color: "rgba(255,255,255,0.35)", fontSize: 11, marginTop: 2 }}>
                    {cat.label} · {r.date}
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <span style={{
                  fontSize: 15, fontWeight: 700,
                  color: r.type === "income" ? "#7fffd4" : "#ff9f9f"
                }}>
                  {r.type === "income" ? "+" : "-"}{formatMXN(r.amount)}
                </span>
                <button onClick={() => deleteRecord(r.id)} style={{
                  background: "rgba(255,100,100,0.1)", border: "none",
                  borderRadius: 10, width: 32, height: 32, cursor: "pointer",
                  color: "#ff9f9f", fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center"
                }}>
                  ✕
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );

  const ExportScreen = () => (
    <div style={{ display: "flex", flexDirection: "column", gap: 22 }}>
      <h2 style={{ margin: 0, color: "#fff", fontSize: 26, fontFamily: "'Georgia', serif" }}>Exportar</h2>

      {/* Summary preview */}
      <div style={{
        background: "rgba(255,255,255,0.04)", borderRadius: 20,
        padding: 20, border: "1px solid rgba(255,255,255,0.08)"
      }}>
        <p style={{ margin: "0 0 14px", color: "rgba(255,255,255,0.5)", fontSize: 12, letterSpacing: 1.5, textTransform: "uppercase" }}>
          Vista previa del reporte
        </p>
        {[
          { label: "Total Ingresos", value: totals.income, color: "#7fffd4" },
          { label: "Total Gastos", value: totals.expense, color: "#ff9f9f" },
          { label: "Balance Final", value: totals.balance, color: totals.balance >= 0 ? "#7fffd4" : "#ff9f9f" },
        ].map((row) => (
          <div key={row.label} style={{
            display: "flex", justifyContent: "space-between",
            padding: "10px 0", borderBottom: "1px solid rgba(255,255,255,0.06)"
          }}>
            <span style={{ color: "rgba(255,255,255,0.6)", fontSize: 14 }}>{row.label}</span>
            <span style={{ color: row.color, fontWeight: 700, fontSize: 15 }}>{formatMXN(row.value)}</span>
          </div>
        ))}
        <div style={{ marginTop: 12, color: "rgba(255,255,255,0.4)", fontSize: 12 }}>
          📋 {records.length} movimientos · 2 hojas de Excel
        </div>
      </div>

      {/* Download only */}
      <button onClick={exportExcel} style={{
        padding: "16px", borderRadius: 16, border: "1px solid rgba(96,165,250,0.3)",
        background: "rgba(96,165,250,0.1)", color: "#60a5fa", fontSize: 15, fontWeight: 700,
        cursor: "pointer", transition: "all 0.15s",
      }}
        onMouseEnter={e => { e.currentTarget.style.background = "rgba(96,165,250,0.2)"; }}
        onMouseLeave={e => { e.currentTarget.style.background = "rgba(96,165,250,0.1)"; }}
      >
        📥 Descargar Excel (MisFinanzas.xlsx)
      </button>

      {/* Divider */}
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ flex: 1, height: 1, background: "rgba(255,255,255,0.08)" }} />
        <span style={{ color: "rgba(255,255,255,0.3)", fontSize: 12 }}>o también por correo</span>
        <div style={{ flex: 1, height: 1, background: "rgba(255,255,255,0.08)" }} />
      </div>

      {/* Email */}
      <div>
        <p style={{ margin: "0 0 8px", color: "rgba(255,255,255,0.5)", fontSize: 12, letterSpacing: 1.5, textTransform: "uppercase" }}>
          Correo electrónico
        </p>
        <input
          type="email" placeholder="tucorreo@ejemplo.com" value={email}
          onChange={e => setEmail(e.target.value)}
          style={{
            width: "100%", padding: "14px 16px", borderRadius: 14,
            background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)",
            color: "#fff", fontSize: 15, boxSizing: "border-box", outline: "none",
          }}
        />
      </div>

      <button onClick={handleExportAndEmail} style={{
        padding: "18px", borderRadius: 16, border: "none", cursor: "pointer",
        background: "linear-gradient(135deg, #3b82f6, #60a5fa)",
        color: "#fff", fontSize: 17, fontWeight: 800,
        boxShadow: "0 8px 30px rgba(59,130,246,0.4)",
        transition: "transform 0.15s",
      }}
        onMouseEnter={e => e.currentTarget.style.transform = "scale(1.02)"}
        onMouseLeave={e => e.currentTarget.style.transform = "scale(1)"}
      >
        📧 Descargar y Enviar por Correo
      </button>

      <p style={{ margin: 0, color: "rgba(255,255,255,0.25)", fontSize: 12, textAlign: "center", lineHeight: 1.6 }}>
        El archivo se descargará automáticamente.<br />Se abrirá tu cliente de correo para adjuntarlo.
      </p>
    </div>
  );

  const screens = { home: <HomeScreen />, add: <AddScreen />, history: <HistoryScreen />, export: <ExportScreen /> };

  const navItems = [
    { id: "home", emoji: "🏠", label: "Inicio" },
    { id: "add", emoji: "＋", label: "Agregar" },
    { id: "history", emoji: "📋", label: "Historial" },
    { id: "export", emoji: "📊", label: "Exportar" },
  ];

  return (
    <div style={{
      minHeight: "100vh", background: "#0a0a14",
      display: "flex", justifyContent: "center", alignItems: "center",
      padding: 20, fontFamily: "-apple-system, 'SF Pro Display', 'Helvetica Neue', sans-serif",
    }}>
      {/* iOS phone frame */}
      <div style={{
        width: 390, minHeight: 844,
        background: "#0f0f1a",
        borderRadius: 48,
        boxShadow: "0 40px 120px rgba(0,0,0,0.8), 0 0 0 1px rgba(255,255,255,0.1), inset 0 0 0 1px rgba(255,255,255,0.05)",
        display: "flex", flexDirection: "column",
        overflow: "hidden", position: "relative",
      }}>
        {/* Status bar */}
        <div style={{
          padding: "14px 28px 8px", display: "flex",
          justifyContent: "space-between", alignItems: "center"
        }}>
          <span style={{ color: "#fff", fontSize: 15, fontWeight: 600 }}>9:41</span>
          <div style={{
            width: 120, height: 32, background: "#000", borderRadius: 20,
            position: "absolute", left: "50%", transform: "translateX(-50%)"
          }} />
          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
            <span style={{ color: "#fff", fontSize: 12 }}>●●●</span>
            <span style={{ color: "#fff", fontSize: 12 }}>📶</span>
            <span style={{ color: "#fff", fontSize: 12 }}>🔋</span>
          </div>
        </div>

        {/* Content */}
        <div style={{ flex: 1, padding: "12px 20px 100px", overflowY: "auto" }}>
          {screens[view]}
        </div>

        {/* Tab bar */}
        <div style={{
          position: "absolute", bottom: 0, left: 0, right: 0,
          background: "rgba(15,15,26,0.92)",
          backdropFilter: "blur(20px)",
          borderTop: "1px solid rgba(255,255,255,0.07)",
          padding: "12px 8px 28px",
          display: "flex", justifyContent: "space-around",
        }}>
          {navItems.map((item) => (
            <button key={item.id} onClick={() => setView(item.id)} style={{
              background: "none", border: "none", cursor: "pointer",
              display: "flex", flexDirection: "column", alignItems: "center", gap: 4,
              padding: "6px 12px", borderRadius: 14,
              transition: "all 0.2s",
              background: view === item.id ? "rgba(255,255,255,0.08)" : "transparent",
            }}>
              <span style={{
                fontSize: item.id === "add" ? 22 : 20,
                filter: view === item.id ? "none" : "grayscale(0.7) opacity(0.5)"
              }}>
                {item.emoji}
              </span>
              <span style={{
                fontSize: 10, fontWeight: 600,
                color: view === item.id ? "#fff" : "rgba(255,255,255,0.35)",
                letterSpacing: 0.3
              }}>
                {item.label}
              </span>
            </button>
          ))}
        </div>
      </div>

      {/* Toast */}
      {toast && (
        <div style={{
          position: "fixed", bottom: 40, left: "50%", transform: "translateX(-50%)",
          background: toast.type === "error" ? "#ff4444" : "#1a1a2e",
          border: `1px solid ${toast.type === "error" ? "#ff6666" : "rgba(127,255,212,0.3)"}`,
          color: "#fff", padding: "14px 24px", borderRadius: 16, fontSize: 14, fontWeight: 600,
          boxShadow: "0 10px 40px rgba(0,0,0,0.5)", zIndex: 999,
          animation: "slideUp 0.3s ease",
        }}>
          {toast.msg}
        </div>
      )}

      <style>{`
        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; }
        input::placeholder { color: rgba(255,255,255,0.25); }
        input[type="date"]::-webkit-calendar-picker-indicator { filter: invert(0.5); }
        * { -webkit-tap-highlight-color: transparent; }
        @keyframes slideUp {
          from { opacity: 0; transform: translateX(-50%) translateY(20px); }
          to { opacity: 1; transform: translateX(-50%) translateY(0); }
        }
        ::-webkit-scrollbar { display: none; }
      `}</style>
    </div>
  );
}