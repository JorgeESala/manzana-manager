# AGENTS.md — Manzana 80 Condo Management

## What This Is

Two Google Apps Script projects (Ingresos + Egresos) deployed to separate Google Spreadsheets via `clasp`. No npm, no tests, no build system. All code runs server-side in Google's V8 runtime or client-side in HTML dialogs.

## Project Structure

```
Ingresos/          ← Income module (main logic, ~1140 lines in Code.js)
  Code.js          ← All backend functions (menu, payments, charges, reports, portal)
  Instalador.js    ← Sheet structure installer (run once to set up headers)
  FormPago.html    ← Payment form (3-step: search → select charge → pay)
  FormCargos.html  ← Monthly charge generation form
  FormPropietario.html ← Add new owner form
  Portal.html      ← Owner self-service web app (React + Tailwind, served via doGet)
Egresos/           ← Expense module (separate spreadsheet)
  Código.js        ← Expense registration
  ReportesFinancieros.js ← Cash flow dashboard (reads Ingresos spreadsheet)
  ReportesAvanzados.js   ← Financial report by date range
  FormEgreso.html  ← Expense form
```

## Critical Concepts

### Edificio ≠ Manzana
- **Edificio** (column B in Propietarios): Building number like "P1", "P2"
- **Manzana** (column I in Propietarios): Block name "MZ 17" or "MZ 19"
- Both blocks can have the same building+dept (e.g., P2 101 exists in both MZ 17 and MZ 19)
- idVivienda format is "P2 101" — block is tracked separately

### Sheet Column Layouts (Instalador.js is the source of truth)

**Propietarios:** A=ID, B=Edificio, C=Depto, D=Nombre, E=Tel, F=Correo, G=Estatus, H=ID Vivienda (computed), I=Manzana

**Cargos:** A=Fecha, B=ID Vivienda, C=Concepto, D=Monto (original, never modify), E=Estatus, F=Mes (format "Mes Año"), G=Manzana, H=Saldo (remaining balance)

**Historial:** A=Fecha, B=Edificio, C=Depto, D=Nombre, E=Concepto, F=Detalle, G=Monto, H=Forma Pago, I=Referencia, J=Manzana

**Egresos:** A=Fecha, B=Folio, C=Categoría, D=Manzana, E=Descripción, F=Monto, G=Método, H=Referencia

### Monto vs Saldo
- **Monto (column D):** Original charge amount. NEVER modify this.
- **Saldo (column H):** Remaining balance. Decreases with each payment. Set to 0 when paid.
- On partial payment: update Saldo only, leave Monto unchanged.

### Month Format
Charges store month as "Mes Año" (e.g., "Junio 2026", not just "Junio"). When comparing with current month, always use the full format:
```javascript
const mesAnioActual = mesActualStr + " " + fechaHoy.getFullYear();
const esMesAnterior = mes !== mesAnioActual;
```
Never compare just the month name — it will break for same-month-different-year charges.

### Late Fee (Recargo)
- Flat $50 (not 10%) for maintenance charges of $500 only
- Applies when: charge month+year ≠ current month+year (any previous month/year) OR day > 10
  - June 2026 charge paid on June 3, 2026 → no late fee (same month+year, day ≤ 10)
  - June 2026 charge paid on June 11, 2026 → late fee (same month+year, but day > 10)
  - June 2025 charge paid on June 3, 2026 → late fee (different year)
  - May 2026 charge paid on June 3, 2026 → late fee (different month)
- Applied to Saldo, not Monto
- Displayed in Portal and payment forms with "$500 → $550 (+Recargo)" indicator

### Block Selector
Panel sheet B5 has a dropdown (MZ 17 / MZ 19). Some functions filter by this, others show both blocks. Dashboard and blocklist show both blocks; charge generation uses both blocks automatically.

## Deployment

```bash
# Push Ingresos to Google
cd Ingresos && clasp push

# Push Egresos to Google
cd Egresos && clasp push
```

Each module has its own `.clasp.json` with a different `scriptId` pointing to different spreadsheets.

## Cross-Spreadsheet Reference

Egresos reads from the Ingresos spreadsheet via hardcoded ID in ReportesFinancieros.js and ReportesAvanzados.js:
```javascript
const ID_ARCHIVO_INGRESOS = "1MROUyVaih7gAcS1MhCAVt4167Em7ewiugqLOnNvILbY";
```

## Common Pitfalls

1. **Don't use `generarCargosMes()`** — legacy function that only stores depto number, not full idVivienda. Use `abrirFormularioCargos()` instead.
2. **`registrarPagoInteligente()` is also legacy** — uses prompts instead of forms and doesn't look up owner names. Use `abrirFormularioPago()` instead.
3. **When adding functions to the menu**, update `onOpen()` in Code.js.
4. **HTML forms** use Bootstrap 5 for styling. Portal uses React + Tailwind.
5. **Data validation** for Manzana dropdown is set in Instalador.js — re-run installer to apply.
6. **The `generarRecibo()` function** reads from the Recibos sheet template cells (B3, G2, F18, etc.) — don't change those cell references.
7. **Propietarios column H** (ID Vivienda) is assumed but not created by the installer — it must exist manually or be added.
