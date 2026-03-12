# Project_TTM — Trade Ticket Master v6

> Automated Bond Trade Processing System (Excel VBA)
> Version: v6.5 | Last Updated: 2026-03-12

## Overview

Trade Ticket Master v6 is a production-grade Excel VBA system that automates the end-to-end front-to-back office workflow for bond trade processing — from Bloomberg real-time data ingestion and daily trade report parsing, through trade ticket generation, to regulatory compliance checking and one-click export.

Built for institutional fixed income portfolio management, the system manages an arbitrary number of portfolios via a configuration-driven architecture — no VBA code changes required to add, remove, or modify portfolio definitions.

## Key Technical Highlights

### In-Memory Array I/O with Dictionary Indexing

The core sync engine (`SyncDataByPrimaryKey`) reads entire worksheet ranges into VBA arrays in a single I/O operation, performs all matching and transformation in memory using `Scripting.Dictionary` for O(1) primary-key lookups, then writes results back in one batch. This replaces the naive cell-by-cell approach (~35,000 individual read/write calls for 1,000 bonds × 35 columns) with exactly 2 I/O operations regardless of data size.

### Dynamic Header Mapping

All data synchronization and compliance checking use runtime header discovery — column positions are resolved by scanning Row 2 headers into a Dictionary at execution time, not by hardcoded column indices. This means Bloomberg fields can be added, removed, or reordered in the spreadsheet without any VBA modification. The compliance engine (`modCompliance`) validates all 28 required fields exist before execution, listing any missing fields and aborting gracefully.

### Configuration-Driven Portfolio Management

Portfolio definitions (account number, name, company, accounting classification) live in a dedicated `Config_Portfolio` worksheet. `InitPortfolios` dynamically loads them into a runtime array with three-layer defense: missing worksheet → `Err.Raise`, blank/non-numeric rows → auto-skip, zero valid rows → `Err.Raise`. All downstream modules — trade ticket layout, compliance checking, data sync — automatically adapt to any portfolio count.

### Defensive Programming & State Preservation

Every long-running subroutine follows the State Preservation Pattern: save `Application.Calculation`, `EnableEvents`, and `ScreenUpdating` before entry, then restore them unconditionally via `On Error GoTo ErrHandler` → `CleanUp` block. This eliminates the classic VBA failure mode where an unhandled error permanently locks Excel into manual calculation with events disabled.

Additional defensive measures include:

- **Error Bubbling**: Low-level errors propagate via `Err.Raise` to the `RunDaily` checkpoint system (no silent `MsgBox` swallowing)
- **SafeCDbl wrapper**: All numeric conversions guard against Bloomberg returning non-numeric strings ("TBD", "N/A", "—")
- **EnableCancelKey**: User interrupts (Esc / Ctrl+Break) route through ErrHandler for clean teardown
- **Template Row Approach**: New bond BDP formulas copy from a validated Row 3 template, preventing dirty-data propagation from the last row
- **Pre-execution Save Point**: `ThisWorkbook.Save` before `InitPortfolios` ensures recovery from hard crashes

### Smart Deduplication & Data Integrity

`CleanupDuplicates` performs same-day deduplication with built-in protection for the historical ledger worksheet (`Bond交易明細`). `AppendToBondDetail` uses a delete-then-write strategy for idempotent daily appends. `NukeGhostData` performs precision physical row deletion (actual data end + 100-row buffer) to reset Excel's `UsedRange` without the multi-second penalty of deleting 1M empty rows.

### Dynamic Trade Ticket Layout

`FillTradeTicketFromDetail` uses a 5-phase dynamic layout engine:

1. **Count** — Tally buy/sell trades per portfolio
2. **Calculate** — Compute row positions for headers, subtotals, and grand totals
3. **Clear & Write** — Reset content area and write section structure
4. **Fill** — Populate trade data into calculated positions
5. **Rebuild Signatures** — Dynamically place approval signature fields below the last data row

This replaced the original fixed-slot design (hardcoded 50 rows per section) with a fully elastic layout that adapts to any number of trades and portfolios.

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│  modMain — Entry Points & Orchestration                 │
│  RunDaily (9-step checkpoint) │ RunSyncBBG │ Export     │
└──────────┬──────────────────────────────────┬───────────┘
           │                                  │
     ┌─────▼──────┐                    ┌──────▼──────┐
     │ modConfig  │                    │modCompliance│
     │ Portfolio  │                    │ 6-Rule      │
     │ Loader     │                    │ Compliance  │
     │ (Config-   │                    │ Engine +    │
     │  Driven)   │                    │ Dynamic     │
     └─────┬──────┘                    │ Header      │
           │                           │ Validation  │
     ┌─────▼───────────────────────┐   └─────────────┘
     │  modProcess — Core Engine                      │
     │  SyncDataByPrimaryKey (shared DRY engine)      │
     │  ├─ In-Memory Array I/O                        │
     │  ├─ Dictionary-based header mapping             │
     │  ├─ State Preservation Pattern                 │
     │  └─ Error Bubbling to Checkpoint               │
     │                                                │
     │  ReadPAM → Detect → Enrich → Sync → Append    │
     │  → FillTradeTicket → Dedup → RefreshDATAFORFIN │
     └────────────────────────────────────────────────┘
```

### Module Summary

| Module | Lines | Responsibility |
|--------|------:|----------------|
| **modConfig** | 130 | Configuration-driven portfolio loading, type declarations, utility functions |
| **modMain** | 383 | Entry points (`RunDaily`, `RunSyncBBG`, `ExportTradeTicket`), 9-step checkpoint tracking, persistent logging |
| **modProcess** | 759 | Core data pipeline: trade report parsing, bond detection, Bloomberg sync, trade ticket generation, deduplication |
| **modCompliance** | 308 | 6-rule compliance engine with dynamic header validation and composite credit rating scoring |
| **Total** | **1,580** | — |

## Worksheets

### Input (3 sheets)

| Sheet | Purpose |
|-------|---------|
| **PAM_Input** | Raw daily trade data (paste from in-house or sub-advisory portfolio manager's daily report) |
| **Restricted_List** | Group-level restricted investment list (ticker, industry, country) |
| **matrix** | Country/credit rating lookup (DM/EM classification, rating-to-score conversion) |

### Output (4 sheets)

| Sheet | Purpose |
|-------|---------|
| **Trade_Ticket** | Standardized bond execution record (N portfolios × Buy/Sell, dynamic layout) |
| **Compliance_Report** | Buy-side compliance results (30 columns, PASS/FAIL/SKIP per trade) |
| **Bond交易明細** | Historical trade ledger (23 columns) |
| **DATAFORFIN** | Bond master data report for middle/back-office operations including ESG compliance checks |

### System (3 sheets)

| Sheet | Purpose |
|-------|---------|
| **BBG_DATABASE** | Bloomberg BDP formula layer (~2,600 bonds × 60 fields, live formulas) |
| **BBG_Value** | Pure-value snapshot of BBG_DATABASE for fast VBA reads (header-matched sync) |
| **Config_Portfolio** | Portfolio definitions — editable by end users, no VBA changes needed |

## Compliance Engine

Six sequential rules (first FAIL terminates):

| # | Check | Condition | FAIL Reason |
|---|-------|-----------|-------------|
| 1 | Restricted List | Ticker or Industry on group list | Not allowed by group policy |
| 2 | Coal Energy | `COAL_ENERGY_CAPACITY_PCT > 30%` | Not allowed by group policy |
| 3 | Floating Rate | `RESET_IDX = SOFRRATE` | Floating rate reset daily |
| 4 | Issuer Equity | Equity < 0, or Equity = 0 with poor rating (excl. Sovereign) | Issuer equity < 0 |
| 5 | Credit Rating | Composite score > 10 | Rating constraints |
| 6 | IMA Constraints | Convertible / AT1 Bail-In / Preferred | IMA constraints |

Credit rating uses a three-agency composite: per-agency best-of (bond → issuer → guarantor), then composite worst-of-three (or best-of-available if fewer than three agencies rate the bond).

## Daily Workflow

```
1. Open workbook → wait for Bloomberg data load (no #N/A)
2. Paste daily trade report into PAM_Input (Ctrl+A → Ctrl+C → Ctrl+V at A1)
3. Click "RunDaily" button on Trade_Ticket sheet
   └─ InitPortfolios → ReadPAM → DetectNewBonds → EnrichFromBBG
      → SyncDBToValue → AppendToBondDetail → FillTradeTicket
      → CleanupDuplicates (×2) → RefreshDATAFORFIN
4. Click "RunComplianceCheck" on Compliance_Report sheet
5. Run ExportTradeTicket → saves pure-value .xlsx (no formulas, no macros)
```

## Technical Notes

### File Encoding

All `.bas` files are encoded in **Big5 (CP950)**, the standard encoding for VBA Editor on Traditional Chinese Windows. Import via VBA Editor: `File → Import File` — the editor handles Big5 natively.

### Adding Bloomberg Fields

The system uses dynamic header mapping — no VBA changes needed:

1. Add the new field header to `BBG_DATABASE` Row 2
2. Add the corresponding BDP formula in Row 1, drag down
3. Add the identical header to `BBG_Value` Row 2 (case-sensitive match)
4. Run `RunDaily` or `RunSyncBBG` — sync engine auto-discovers new columns

### Adding Portfolios

Edit the `Config_Portfolio` worksheet directly (no VBA changes):

1. Append a new row: Account Number (A), Name (B), Company (C), Accounting Class (D)
2. Save — `InitPortfolios` auto-loads on next `RunDaily`

## License

This project is released under the [MIT License](LICENSE).
