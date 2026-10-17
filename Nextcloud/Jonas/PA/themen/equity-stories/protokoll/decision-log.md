# Decision Log — Equity Story Coding

Binding precedents after calibration.

## Format

- **YYYY-MM-DD — Topic** — Decision. Rationale. Example unit IDs.

## Accepted precedents

- **2026-10-17 — Trust awards → multiple coding 2.3 + 3.2** — External trust awards / rankings (e.g. Newsweek/Statista “World’s Most Trustworthy Companies”) simultaneously satisfy Brand Strength (reputation recognition) and Customer Relationships & Retention (customer-trust signal). Code **both** 2.3 and 3.2. Example: `Swissquote-027`.

- **2026-10-17 — 10.4 unchanged** — Proposed broadening of 10.4 (risk/investment-approach as strength without named instruments) **rejected**; definition restored to original. `Swissquote-112` (*Prudent investment approach*) remains **99**.


- **2026-10-16 — Split before framing/cause reading** — If a sentence contains two substantively independent investment arguments, **split into two coding units** (Leitfaden Coding Unit + example „strong balance sheet and experienced management team“). Do **not** collapse a customer/trust (or similar) claim into a capital/financial claim by reading the second clause only as “cause/framing”. Multiple coding only when the formulation is not cleanly separable; rule of thumb ≤2 categories, else split. Examples: `Swissquote-014` / `Swissquote-015` (trust vs capital); also applied in AI-v1.1 resplit pass to 001, 002, 042, target-segment bullets, ESG ratings/topics, domestic-market cluster.

## Open (from AI-v1 Swissquote pilot)

Pending human comparison / Jonas review:

1. **Summary slides:** Exclude thematic restatements to avoid frequency double-counting? (AI excluded p.34)
2. **„sustainable growth“** without ESG content → 13.x only, not 9.x? (Swissquote-007, U=1)
3. **Non-transaction income „increased in importance“** → 1.2 vs 1.1 vs 10.2
4. **Innovation awards / „most innovative companies“** → 5.3 vs 2.3
5. **Chart-only figures:** support prose only, not separate units (AI practice) — confirmed in gate review
