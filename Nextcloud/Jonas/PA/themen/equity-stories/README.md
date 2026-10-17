# Equity Story Coding

Research coding of Equity Stories using a fixed category system (Excel) and coding guide.

## Status

- **Guide:** [`leitfaden/Coding-Guidelines.md`](leitfaden/Coding-Guidelines.md)
- **Swissquote:** coded as `AI-v4` (2026-10-17); calibrations applied (trust awards → 2.3+3.2; 111 Exclude)
- **Next:** human review of `protokoll/swissquote-review-pack.md` + master Excel

## Layout

| Path | Content |
| --- | --- |
| `leitfaden/` | NEW Coding Guidelines (`Coding-Guidelines.md`) |
| `vorlage/` | Excel template + **Instructions** sheet (NEW guide rules) |
| `quellen/` | Source PDF / extracts / OCR |
| `coding/` | Master workbook + `swissquote-units.json` |
| `protokoll/` | Exclusion log + review pack |

## Conventions

- **Unit ID:** `{Company}-{NNN}` (e.g. `Swissquote-001`)
- **Coder (current):** `AI-v4`
- **One row = one coding unit**; subcategory columns binary `0/1`
- Do **not** edit derived formula columns (`* Present`, Value Driver, Investment Outcome)
- **Uncertainty / Review Flag = 1** for borderline split/category choice or >2 candidate subcats
- **Original Statement** = verbatim quote only (no `Summary:` / pillar wrappers)
- **9.1** includes external ESG/sustainability ratings (MSCI, Sustainalytics, …)

## Swissquote snapshot (AI-v4)

- Units: **134** | Residual 99: **3** | Uncertainty: **0**
- Master: `coding/equity-story-coding-master.xlsx`
- Review: `protokoll/swissquote-review-pack.md`
