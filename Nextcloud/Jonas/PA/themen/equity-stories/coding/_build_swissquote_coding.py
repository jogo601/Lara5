#!/usr/bin/env python3
"""Build Swissquote equity-story coding deliverables from guide + source only."""
from __future__ import annotations

import json
import shutil
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

BASE = Path("/Users/jyg/Nextcloud/Jonas/PA/themen/equity-stories")
TEMPLATE = BASE / "vorlage/Coding Table second draft with updated categories.xlsx"
MASTER = BASE / "coding/equity-story-coding-master.xlsx"
JSON_OUT = BASE / "coding/swissquote-units.json"
EXCL_OUT = BASE / "protokoll/swissquote-exclusion-log.md"
REVIEW_OUT = BASE / "protokoll/swissquote-review-pack.md"

META = {
    "company": "Swissquote",
    "source_type": "Investor presentation / Equity Story",
    "access_date": "2026-10-17",
    "language": "en",
    "document_format": "PDF",
    "source_document_type": "presentation",
    "source_url": "themen/equity-stories/quellen/21-swissquote.pdf",
    "coder": "AI-v3",
    "coding_date": "2026-10-17",
}

# Column index (1-based) for subcategory codes
CODE_COL = {
    "1.1": 14, "1.2": 15, "1.3": 16,
    "2.1": 18, "2.2": 19, "2.3": 20, "2.4": 21,
    "3.1": 23, "3.2": 24, "3.3": 25, "3.4": 26,
    "4.1": 28, "4.2": 29, "4.3": 30, "4.4": 31, "4.5": 32,
    "5.1": 34, "5.2": 35, "5.3": 36, "5.4": 37,
    "6.1": 39, "6.2": 40,
    "7.1": 42, "7.2": 43, "7.3": 44, "7.4": 45,
    "8.1": 47, "8.2": 48, "8.3": 49, "8.4": 50,
    "9.1": 52, "9.2": 53, "9.3": 54, "9.4": 55,
    "10.1": 57, "10.2": 58, "10.3": 59, "10.4": 60,
    "11.1": 62, "11.2": 63, "11.3": 64, "11.4": 65,
    "12.1": 67, "12.2": 68, "12.3": 69, "12.4": 70, "12.5": 71,
    "13.1": 73, "13.2": 74, "13.3": 75, "13.4": 76, "13.5": 77,
    "99": 78,
}

DERIVED_COLS = {11, 12, 13, 17, 22, 27, 33, 38, 41, 46, 51, 56, 61, 66, 72}

# units: statement, location, codes[], note, uncertainty
UNITS: list[dict] = []


def u(statement: str, location: str, codes: list[str], note: str = "", uncertainty: int = 0):
    UNITS.append(
        {
            "statement": statement,
            "location": location,
            "codes": codes,
            "note": note,
            "uncertainty": uncertainty,
        }
    )


# ── Page 4 ──
u(
    "Swissquote combines a unique multi-asset class trading platform with the innovative power of a fintech company based on a solid Swiss bank.",
    "p.4",
    ["2.2", "5.1", "12.1"],
    "Inseparable CEO USP quote combining platform differentiation, fintech innovation, and solid Swiss bank; coded best two (2.2, 5.1) plus 12.1 noted as third candidate — uncertainty flagged per >2 rule (two best-fitting kept as 2.2+5.1; 12.1 also coded as core to quote meaning — review).",
    1,
)
# Re-do page 4 properly per guide: if >2 apply, first split. Split into three units:
UNITS.pop()
u(
    "unique multi-asset class trading platform",
    "p.4",
    ["2.2"],
    "From CEO quote: Swissquote combines a unique multi-asset class trading platform with the innovative power of a fintech company based on a solid Swiss bank.",
)
u(
    "innovative power of a fintech company",
    "p.4",
    ["5.1"],
    "From CEO quote (split): innovative power of a fintech company.",
)
u(
    "based on a solid Swiss bank",
    "p.4",
    ["12.1"],
    "From CEO quote (split): solid Swiss bank as capitalization/solidity argument.",
)

# ── Page 5 ──
u(
    "A profitably growing tech bank",
    "p.5",
    ["12.3"],
    "Equity-story headline presenting achieved profitable growth as tech bank positioning.",
)
u(
    "Swissquote combines the pedigree of being a fully regulated Swiss bank with the innovative mindset of a tech company which is engrained in our DNA.",
    "p.5",
    ["2.2", "4.4"],
    "Split candidates considered; kept as dual coding: regulated Swiss-bank pedigree as differentiation (2.2) and tech mindset engrained in DNA as culture (4.4).",
)
u(
    "TRUST & SECURITY — A fully regulated Swiss bank",
    "p.5",
    ["2.2"],
    "OCR figure label pairing regulated Swiss bank with trust & security as differentiation.",
)
u(
    "TECH LEADERSHIP & INNOVATION — The innovative mindset of a tech company",
    "p.5",
    ["5.3"],
    "OCR figure label explicitly claiming tech leadership (distinct from general innovation wording in body text).",
)
u(
    "Over the years, we have built a leading multi-asset platform with over 3 million of online tradable financial products",
    "p.5",
    ["2.1"],
)
u(
    "that provides a top-notch customer experience",
    "p.5",
    ["3.1"],
    "Continuation of platform sentence; customer experience as distinct argument.",
)
u(
    "with competitive and transparent pricing",
    "p.5",
    ["2.2"],
    "Continuation of platform sentence; pricing as differentiation.",
)
u(
    "resulting in a strong market position in Switzerland",
    "p.5",
    ["2.1"],
)
u(
    "We are now expanding it to additional geographies in order to provide more customers with access to our platform",
    "p.5",
    ["7.1"],
)
u(
    "and to continue our story of profitable and sustainable growth",
    "p.5",
    ["13.1"],
    "'Sustainable' used in growth/business sense, not ESG product claim.",
)

# ── Page 6 ──
u(
    "We want to be the first bank for digital-first mass affluent traders and investors.",
    "p.6",
    ["7.3"],
    "Vision statement targeting digital-first mass affluent segment.",
)
u(
    "We challenge the code to deliver innovative services and products that make financial opportunities accessible to ambitious, self-directed people.",
    "p.6",
    ["5.1"],
)

# ── Page 7 ──
u(
    "We have a strong track record of profitable growth",
    "p.7",
    ["12.3"],
)
u(
    "Over the last few years our platform has attracted large numbers of new clients and new assets.",
    "p.7",
    ["3.4"],
)
u(
    "In benign markets, economies of scale kicked in and margin increased.",
    "p.7",
    ["1.3", "12.3"],
    "Economies of scale (1.3) with achieved margin increase (12.3); dual coding.",
)
u(
    "In less benign markets, we managed to maintain a fair margin level.",
    "p.7",
    ["10.1", "12.3"],
    "Earnings resilience in weaker markets with maintained margin.",
)

# ── Page 10 ──
u(
    "Clients trust us, helped by strong capital and solid regulatory ratios",
    "p.10",
    ["3.2", "12.1"],
    "Trust (3.2) linked to strong capital/regulatory ratios (12.1).",
)
u(
    "Liquidity ratio LCR 309% / min 100%",
    "p.10",
    ["12.2"],
    "Figure label arguing liquidity strength vs regulatory minimum.",
)
u(
    "Funding ratio NSFR 263% / min 100%",
    "p.10",
    ["12.2"],
)
u(
    "Leverage ratio 7.2% / min 3.0%",
    "p.10",
    ["12.1"],
)
u(
    "Interest rate risk IRRBB 3.6% / max 15%",
    "p.10",
    ["10.4"],
    "Interest-rate risk metric presented as comfortably within limit.",
)
u(
    "Total equity / Capital ratio trajectory with FINMA requirement (chart)",
    "p.10",
    ["12.1"],
    "Chart labels (Total equity, Capital ratio, FINMA requirement) argue capitalization strength; numerical series not invented beyond extract labels.",
)
u(
    'category of "World\'s Most Trustworthy Companies" by Newsweek',
    "p.10",
    ["2.3"],
    "OCR figure text (partially noisy) presenting Newsweek Most Trustworthy Companies recognition as trust/brand signal; statement reconstructed from OCR labels.",
    1,
)

# ── Page 13 (OCR) ──
u(
    "Our clients like our broad offering based on 3 pillars",
    "p.13",
    ["1.1"],
    "Pillar detail coded in subsequent OCR units (TRADE / INVEST / BANK + Yuh).",
)
u(
    "TRADE pillar: Become the most extensive and intuitive app for all trading needs. — SECURITIES TRADING, FOREX/CFD, CRYPTO, THEMES TRADING",
    "p.13",
    ["7.2", "5.1"],
    "OCR: multi-asset trading-app ambition as offering expansion (7.2) and product/service innovation (5.1).",
)
u(
    "INVEST pillar: Become recognised for easy accessible investment and saving solutions. — PORTFOLIO MANAGEMENT, SUSTAINABLE INVESTING, SAVINGS",
    "p.13",
    ["7.2", "3.3"],
    "OCR: accessible investment/saving solutions as offering expansion and value-added solutions; Sustainable Investing also supports 9.4 (noted; dual already at two-code limit).",
    1,
)
u(
    "INVEST pillar label: SUSTAINABLE INVESTING",
    "p.13",
    ["9.4"],
    "OCR product label within INVEST pillar; sustainable investing offering for clients.",
)
u(
    "BANK pillar: Become the top-of-wallet bank of our clients, with 360° services, 365 days a year. — DEBIT CARDS, ONLINE PAYMENTS, eMORTGAGES, LOANS",
    "p.13",
    ["7.2", "3.3"],
    "OCR: full banking-service package as offering expansion and value-added 360° services.",
)
u(
    "Joint venture expands Swissquote's reach and offering",
    "p.13",
    ["8.1", "7.2"],
    "OCR: Yuh JV framed as expanding reach/offering (inorganic JV + offering expansion).",
)
u(
    "Yuh, our joint venture with PostFinance, for instance, continues to provide very low cost banking, saving and investing options.",
    "p.13",
    ["2.2", "6.2"],
    "OCR: low-cost banking/saving/investing as competitive differentiation and cost efficiency of the offering.",
)
u(
    "Its main target audience is young, mobile-first users in Switzerland.",
    "p.13",
    ["7.3"],
    "OCR: Yuh target segment (young, mobile-first) in Switzerland.",
)

# ── Page 14 ──
u(
    "Our strong client service and savvy marketing translate into…",
    "p.14",
    ["3.1"],
    "Heading claim of strong client service (3.1). 'Savvy marketing' has no clear subcategory fit and is not coded as 6.1; outcomes coded in following units.",
)
u(
    "swissquote.ch as the leading online source of financial information in Switzerland since 2016",
    "p.14",
    ["2.1"],
)
u(
    "Top-of-mind online trading bank in Switzerland since 2021",
    "p.14",
    ["2.3"],
)
u(
    "Top rating recognition as one of Switzerland's most innovative companies by Bilanz, PME and Statista",
    "p.14",
    ["5.1"],
)
u(
    "A STRONG BRAND",
    "p.14",
    ["2.3"],
)
u(
    "Positive Net Promoter Score (NPS®) of 39 in 2024 (+2 compared to last year)",
    "p.14",
    ["3.2"],
)
u(
    "Client satisfaction score of 5.9 out of 7 (2023: 6.0) with regulated operator status",
    "p.14",
    ["3.2"],
)

# ── Page 15 ──
u(
    "Key long-term driver: New clients in the right segments",
    "p.15",
    ["7.3"],
)
u(
    "OUR TARGET SEGMENT: Affluent and mass affluent — Generate higher revenues per client",
    "p.15",
    ["7.3", "3.4"],
    "Target-segment strategy (7.3) with higher revenue-per-client customer quality (3.4).",
)
u(
    "Trustworthiness, confidentiality, an innovative offering and convenience are important to them",
    "p.15",
    ["3.1"],
    "Customer-need orientation of target segment proposition.",
)
u(
    "Willing to pay competitive and transparent fees for these services (clients buy satisfaction).",
    "p.15",
    ["2.2"],
    "Willingness to pay for competitive/transparent fee proposition as commercial differentiation.",
)

# ── Page 17 ──
u(
    "Our revenue streams are diverse",
    "p.17",
    ["1.1"],
)
u(
    "Net revenues by asset class (Cash, Shares, FX/precious metals, Fixed income/funds, Structured products/derivatives, Crypto, eFX/CFDs, Others)",
    "p.17",
    ["1.1"],
    "Chart labels evidencing multi-asset revenue diversification; supports diversified business model.",
)
u(
    "Net revenues by customer domicile (Switzerland, Europe, MEA, Rest of the world)",
    "p.17",
    ["1.1"],
    "Geographic revenue mix presented as diversity of streams; no explicit risk-mitigation wording (not 10.3).",
)

# ── Page 18 ──
u(
    "Non-transaction-based income has increased in importance",
    "p.18",
    ["1.2"],
    "Custody/management fees and interest-type non-transaction income framed as growing share of model (recurring/ongoing).",
)
u(
    "Non-transaction-based income: Stable in the short term",
    "p.18",
    ["10.2"],
    "Label contrasting stability of non-transaction income vs sentiment/volatility-driven transaction income.",
)

# ── Page 19 ──
u(
    "We have a strong platform, and third parties like it",
    "p.19",
    ["2.2"],
)
u(
    "However, our platform is attractive, and for a number of years we have had third parties using our platform, partially white labelled.",
    "p.19",
    ["2.2", "7.2"],
    "Platform attractiveness (2.2) and white-label/third-party offering expansion (7.2).",
)
u(
    "Overall fees from B2B are slightly less transaction-dependent than the B2C business.",
    "p.19",
    ["10.1"],
)
u(
    "We expect similar growth rates for B2B and B2C business.",
    "p.19",
    ["13.1"],
)
u(
    "B2C: FULL CAPTURE OF REVENUES; B2B/B2B2C: PROFIT SHARING — client profiles (Retail/Affluent/HNWI; Corporate/Family offices; Fund managers; Asset managers; Brokers/Banks)",
    "p.19",
    ["1.1", "3.4"],
    "Figure labels: dual B2C/B2B model structure (1.1) and diversified client/partner profiles (3.4).",
)

# ── Page 22 ──
u(
    "Future growth is expected from our 'growth portfolio'",
    "p.22",
    ["7.2", "13.1"],
    "Growth-portfolio framing as offering/project expansion with expected growth outcome.",
)
u(
    "Multiple growth opportunities: Multi asset European strategy Roll-out",
    "p.22",
    ["7.1"],
)
u(
    "Multiple growth opportunities: SQX Crypto Exchange",
    "p.22",
    ["7.2", "5.1"],
    "New crypto exchange offering as product expansion and service innovation.",
)
u(
    "Multiple growth opportunities: Yuh mobile app with over 300,000 'yuhsers'",
    "p.22",
    ["5.1", "3.4"],
    "Mobile app innovation with sizable user base.",
)
u(
    "Multiple growth opportunities: Further expansion of B2B/B2B2C Business",
    "p.22",
    ["7.3"],
)
u(
    "Multiple growth opportunities: Universal bank with full daily banking package (e.g. Invest Easy)",
    "p.22",
    ["7.2"],
)

# ── Page 23 ──
u(
    "Our tech DNA is reflected in our employee mix — Technology 36% of headcount (1,217 FTE)",
    "p.23",
    ["4.1", "4.2"],
    "Tech-heavy workforce as specialized technical capability (4.1) and skilled talent mix (4.2); dual coding.",
)

# ── Page 24 ──
u(
    "And our employees like to work with us!",
    "p.24",
    ["4.3"],
)
u(
    "Our company culture is amazing and our Great Place to Work Certification™ proves it.",
    "p.24",
    ["4.4", "4.3"],
    "OCR quote: explicit corporate-culture claim (4.4) evidenced by GPTW certification / engagement (4.3).",
)
u(
    "Certified Great Place to Work early 2024",
    "p.24",
    ["4.3"],
    "Workplace certification as employee engagement/retention strength.",
)
u(
    "Employer Net Promoter Score 34 (2023: 14)",
    "p.24",
    ["4.3"],
)

# ── Page 26 ──
u(
    "Guidance 2025: expected pre-tax profit of CHF 355m",
    "p.26",
    ["13.2"],
)
u(
    "Guidance 2025 chart: Net revenues ~CHF 675m; Pre-tax profit ~CHF 355m",
    "p.26",
    ["13.1", "13.2"],
    "Chart labels for 2025G expected revenues and pre-tax profit.",
)

# ── Pages 27–28 ──
u(
    "Mid-term outlook 2028: ~50K accounts opened on a yearly basis",
    "p.27",
    ["13.1", "7.3"],
    "Expected account growth as earnings/growth outlook and customer expansion.",
)
u(
    "Mid-term outlook 2028: Net new money ~CHF 7 BLN / year (representing ~0.90% of client assets)",
    "p.27",
    ["13.1"],
    "Extract shows '~CH 7 BLN / year' under Net new money.",
)
u(
    "Mid-term outlook 2028: Net revenues ~CHF 900M",
    "p.27",
    ["13.1"],
)
u(
    "Mid-term outlook 2028: Pre-tax profit ~CHF 500M",
    "p.27",
    ["13.1", "13.2"],
)
u(
    "Mid-term outlook 2028: ~55% Pre-tax profit margin assumed",
    "p.27",
    ["13.2"],
)
u(
    "Stable non-transaction based revenues in a low-rate environment and a higher share of banking revenues in transaction based revenues.",
    "p.27",
    ["10.2", "1.2"],
    "Outlook assumption on stable non-transaction revenues (10.2/1.2).",
)
u(
    "Mid-term outlook 2028 chart path: revenues 675→900 (+33%) and pre-tax profit 355→500 (+41%) with margin ~55.6%",
    "p.28",
    ["13.1", "13.2"],
    "Chart labels 2025G/2028O with +33%/+41% growth annotations.",
)

# ── Page 29 ──
u(
    "Thanks to solid equity position, payout ratio at 30%",
    "p.29",
    ["12.1", "13.4"],
)

# ── Page 30 ──
u(
    "As Swissquote maintains a solid capital position (23.5%), the capital exceeding this 18% internal limit is approximately CHF 230 million ('growth buffer').",
    "p.30",
    ["12.1"],
)
u(
    "The growth buffer could be used for selected inorganic opportunities in the coming years. The main objective would be to grow the existing business model.",
    "p.30",
    ["8.1"],
)
u(
    "Swissquote is committed to a dividend distribution representing 30% of the net profit.",
    "p.30",
    ["13.4"],
)
u(
    "Additional distributions (amount, form and timing to be determined at a later stage) may be contemplated when the 'growth buffer' will have reached at least the double of the current size.",
    "p.30",
    ["13.4"],
)

# ── Page 33 ──
u(
    "We are getting recognised for our efforts in Sustainability",
    "p.33",
    ["99"],
    "VERBATIM: We are getting recognised for our efforts in Sustainability. Meets selection criterion (ESG recognition as investment attractiveness) but does not specify climate (9.1), social (9.2), governance (9.3), or sustainable products (9.4).",
)
u(
    "Sustainalytics improved Swissquote's ESG risk rating from 'medium risk' to 'low risk' in the first half of 2023.",
    "p.33",
    ["99"],
    "VERBATIM: Sustainalytics improved Swissquote's ESG risk rating from 'medium risk' to 'low risk' in the first half of 2023. Overall ESG risk rating spans E/S/G; no single 9.x subcategory definition is satisfied.",
)
u(
    "KEY TOPICS E: Energy use and climate resilience; Energy consumption; Greenhouse gas emissions (scope 1, 2, 3)",
    "p.33",
    ["9.1"],
)
u(
    "KEY TOPICS S: Diversity, equity and inclusion; Talent recruitment, development and retention; Social engagements; Protection of human rights; Compensation and benefits",
    "p.33",
    ["9.2"],
)
u(
    "KEY TOPICS G: Compliance, governance and ethics; Transparency and credibility in the market",
    "p.33",
    ["9.3"],
)
u(
    "MSCI ESG Rating stable at AA",
    "p.33",
    ["99"],
    "VERBATIM: MSCI ESG Rating stable at AA. Composite ESG rating without a climate-, social-, governance-, or sustainable-product-specific claim under 9.1–9.4.",
)
u(
    "zRating® ranked Swissquote 8th out of 164 Swiss listed companies in 2024. Ranked first in the financial services industry in Switzerland.",
    "p.33",
    ["9.3"],
    "External ranking listed under G / transparency & credibility on the sustainability slide; fits Governance & Business Ethics (9.3).",
)

# ── Page 34 summary (each bullet is a packaged investment argument) ──
u(
    "Summary: Trusted Swiss bank (capital and liquidity) with a strong brand",
    "p.34",
    ["12.1", "12.2", "2.3"],
    ">2 subcategories apply and unit is summary packaging; coded two best-fitting financial (12.1+12.2) and brand (2.3) — uncertainty for multi-claim summary bullet.",
    1,
)
# Fix: guide says code two best-fitting if >2 and inseparable. So only two:
UNITS[-1] = {
    "statement": "Summary: Trusted Swiss bank (capital and liquidity) with a strong brand",
    "location": "p.34",
    "codes": ["12.1", "2.3"],
    "note": "Summary bullet packs capital/liquidity trust and brand; coded 12.1 + 2.3; additional candidate 12.2 (liquidity) noted for review.",
    "uncertainty": 1,
}
u(
    "Summary: Tech DNA: Track record of innovative products that clients love, with more to come",
    "p.34",
    ["5.1", "3.2"],
)
u(
    "Summary: Diverse revenue streams with increasing importance of non-transaction-based income",
    "p.34",
    ["1.1", "1.2"],
)
u(
    "Summary: Strong growth in target segment: Affluent clients who value our offering",
    "p.34",
    ["7.3", "3.4"],
)
u(
    "Summary: Leading multi-asset platform with top-notch customer service",
    "p.34",
    ["2.1", "3.1"],
)
u(
    "Summary: Strong growth portfolio: focused portfolio of growth opportunities and projects",
    "p.34",
    ["7.2"],
)
u(
    "Summary: Convenient digital platform with competitive and transparent pricing",
    "p.34",
    ["2.2", "5.1"],
    "Digital platform convenience as innovation/offering characteristic plus pricing differentiation.",
)
u(
    "Summary: Considering sustainability as an integral part of Swissquote's strategy",
    "p.34",
    ["99"],
    "VERBATIM: Considering sustainability as an integral part of Swissquote's strategy. Meets selection criterion but does not specify climate, social, governance, or sustainable-product content under 9.1–9.4.",
)

# ── Page 36 ──
u(
    "Swissquote is growing in its domestic market",
    "p.36",
    ["7.1"],
)
u(
    "Swissquote average market share is 2.1% (at 31 December 2023)",
    "p.36",
    ["2.1"],
    "Market-share figure as established domestic position (not leadership claim).",
)
u(
    "Approximately 3% continued growth in wealth per year",
    "p.36",
    ["11.2"],
)
u(
    "Evolution of addressable market — Total net wealth of Swiss households CHF 4.7 trillion; 38% in financial assets (CHF 1.8 trillion)",
    "p.36",
    ["11.2", "11.4"],
    "Addressable-market expansion / wealth pool as external growth opportunity; dual macro (11.2) and structural market (11.4).",
)


EXCLUSIONS = [
    {"page": 1, "text": "2024 FULL YEAR RESULTS / PRESS CONFERENCE / SWISSQUOTE EQUITY STORY", "reason": "Title/cover slide; organizational framing without a self-contained investment argument."},
    {"page": 2, "text": "01 SWISSQUOTE STORY", "reason": "Section divider / navigation."},
    {"page": 3, "text": "STORY / 01", "reason": "Empty/section placeholder; no investment argument."},
    {"page": 4, "text": "Nav labels: Swissquote story / Our USP / Diversified net revenues / Innovation in our DNA / Outlook & payout strategy / Sustainability & summary / Trust & transparency", "reason": "Repeated page navigation chrome; organizational, not coded (applies to all content pages)."},
    {"page": 4, "text": "Marc Bürki, CEO", "reason": "Attribution/label only; not an investment argument."},
    {"page": 7, "text": "Chart axis years 2019–2024 and raw numeric series without standalone textual investment claim beyond coded headings/captions", "reason": "Graphics: only textual labels/captions coded; numeric series alone not separate units beyond track-record statements already coded."},
    {"page": 8, "text": "03 TRUST & TRANSPARENCY", "reason": "Section divider."},
    {"page": 9, "text": "TRANSPARENCY / 02", "reason": "Empty/section placeholder."},
    {"page": 11, "text": "02 OUR USP / DIVERSIFIED NET REVENUES", "reason": "Section divider; substance coded on later USP/revenue pages."},
    {"page": 12, "text": "USP / 03", "reason": "Empty/section placeholder."},
    {"page": 15, "text": "Distribution of accounts / client assets / net revenues pie chart percentages (10%/36%/54% etc.)", "reason": "Graphic percentages without additional textual investment claim beyond 'Key long-term driver: New clients in the right segments' (coded)."},
    {"page": 15, "text": "* 52% of new clients 2024 acquired in H2-2024", "reason": "Timing footnote; no independent investment argument."},
    {"page": 16, "text": "04 NET REVENUES", "reason": "Section divider."},
    {"page": 18, "text": "Component lists (FX transactions and eForex; Commissions on trades; Custody and management fees; Interest income; Crypto-related revenues; Other non-transactional income) as bare source lists", "reason": "Descriptive components of revenue types; investment arguments coded via 'increased in importance' and 'Stable in the short term'."},
    {"page": 18, "text": "Transaction-based income: Sentiment-driven / Dependent on market volatility / Impacted by asset valuation", "reason": "Describes characteristics of transaction income as contrast, not a positive investment-attractiveness argument in itself; stability claim for non-transaction income coded separately."},
    {"page": 19, "text": "Our ultimate customer is always an individual investor, a family office or a corporate entity.", "reason": "Descriptive customer definition without investment-attractiveness justification beyond later platform/B2B arguments."},
    {"page": 20, "text": "04 INNOVATION IN OUR DNA / OUTLOOK & PAYOUT STRATEGY", "reason": "Section divider."},
    {"page": 21, "text": "IN OUR DNA / 05", "reason": "Empty/section placeholder."},
    {"page": 23, "text": "Nationalities 69; Female 32%; Part-time 11%; Average age 35; Turnover 9%; 3% temporary positions; non-tech headcount mix slices", "reason": "Workforce demographics without explicit investment-strength framing (tech mix / engagement coded separately on p.23–24)."},
    {"page": 25, "text": "06 & PAYOUT STRATEGY", "reason": "Section divider."},
    {"page": 29, "text": "EPS / DPS numeric chart series 2019–2024 without additional claim beyond payout-ratio headline", "reason": "Chart numerics; investment argument coded via 'solid equity position, payout ratio at 30%'."},
    {"page": 30, "text": "The minimum set by FINMA is 11.2%. The 18% threshold is an internal limit set by the Board of Directors triggering immediate attention of the latter.", "reason": "Regulatory/internal-limit explanation; investment argument coded via solid capital position / growth buffer statements."},
    {"page": 31, "text": "06 SUSTAINABILITY & SUMMARY", "reason": "Section divider."},
    {"page": 32, "text": "& SUMMARY / 07", "reason": "Empty/section placeholder."},
    {"page": 33, "text": "QUANTITATIVE DATA metric lists (FTE dedicated to technology; App rating; call-answer %; workforce composition metrics; capital and liquidity ratios as data catalogue)", "reason": "ESG data catalogue / reporting fields; recognition and KEY TOPICS investment claims coded separately."},
    {"page": 33, "text": "ASSURED labels", "reason": "Assurance stamps without standalone investment argument."},
    {"page": 35, "text": "08", "reason": "Empty/section placeholder."},
    {"page": 36, "text": "Wealth inequality is moderate. Over 70% of wealth distributed between 10K and 1M", "reason": "External market context footnote without clear favorable-driver framing for Swissquote's investment case (growth/addressable market coded separately)."},
    {"page": 36, "text": "Source footnotes (Swiss National Bank; Global wealth Databook)", "reason": "Citations/sources; not investment arguments."},
    {"page": 37, "text": "THANK YOU / office locations / investorrelations@swissquote.ch / Contact", "reason": "Contact / closing slide; excluded per selection criterion."},
]


def assign_ids():
    for i, unit in enumerate(UNITS, 1):
        unit["id"] = f"Swissquote-{i:03d}"


def write_excel():
    if MASTER.exists():
        MASTER.unlink()
    shutil.copy2(TEMPLATE, MASTER)
    wb = load_workbook(MASTER)
    ws = wb["Coding"]

    # Verify formulas present before write
    sample_formula = ws.cell(3, 11).value
    assert isinstance(sample_formula, str) and sample_formula.startswith("="), "Derived formula missing in template copy"

    for i, unit in enumerate(UNITS):
        r = 3 + i
        # Ensure formulas exist for this row (template has 200); if missing, copy from row 3 pattern
        for col in DERIVED_COLS:
            cell = ws.cell(r, col)
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                # replicate formula from row 3 with adjusted row number
                proto = ws.cell(3, col).value
                if isinstance(proto, str) and proto.startswith("="):
                    cell.value = proto.replace("3", str(r))

        # Metadata + unit fields
        ws.cell(r, 1).value = META["company"]
        ws.cell(r, 2).value = META["source_url"]
        ws.cell(r, 3).value = META["source_type"]
        ws.cell(r, 4).value = META["access_date"]
        ws.cell(r, 5).value = META["language"]
        ws.cell(r, 6).value = META["document_format"]
        ws.cell(r, 7).value = unit["id"]
        ws.cell(r, 8).value = unit["statement"]
        ws.cell(r, 9).value = META["source_document_type"]
        ws.cell(r, 10).value = unit["location"]

        # Initialize all subcategory columns to 0 for this row
        for col in CODE_COL.values():
            ws.cell(r, col).value = 0

        for code in unit["codes"]:
            if code not in CODE_COL:
                raise ValueError(f"Unknown code {code} in {unit['id']}")
            ws.cell(r, CODE_COL[code]).value = 1

        ws.cell(r, 79).value = unit.get("note") or None
        ws.cell(r, 80).value = META["coder"]
        ws.cell(r, 81).value = META["coding_date"]
        ws.cell(r, 82).value = int(unit.get("uncertainty") or 0)

        # Critical: do not overwrite derived formula cells
        for col in DERIVED_COLS:
            val = ws.cell(r, col).value
            assert isinstance(val, str) and val.startswith("="), f"Derived col {col} row {r} lost formula"

    wb.save(MASTER)
    wb.close()


def write_json():
    payload = {
        "meta": {
            **META,
            "source_pdf": "themen/equity-stories/quellen/21-swissquote.pdf",
            "source_extract": "themen/equity-stories/quellen/21-swissquote-extract.txt",
            "ocr_extract": "themen/equity-stories/quellen/21-swissquote-ocr-extract.txt",
            "guide": "themen/equity-stories/leitfaden/neue-coding-guidelines-extract.md",
            "unit_count": len(UNITS),
            "residual_99_count": sum(1 for u in UNITS if "99" in u["codes"]),
            "uncertainty_count": sum(1 for u in UNITS if u.get("uncertainty")),
            "exclusion_count": len(EXCLUSIONS),
        },
        "units": [
            {
                "id": u["id"],
                "statement": u["statement"],
                "location": u["location"],
                "codes": u["codes"],
                "note": u.get("note") or "",
                "uncertainty": int(u.get("uncertainty") or 0),
            }
            for u in UNITS
        ],
        "exclusions": EXCLUSIONS,
    }
    JSON_OUT.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def write_exclusion_log():
    lines = [
        "# Swissquote — Exclusion Log",
        "",
        f"- Company: {META['company']}",
        f"- Source: `{META['source_url']}`",
        f"- Access Date: {META['access_date']}",
        f"- Coder: {META['coder']}",
        f"- Coding Date: {META['coding_date']}",
        f"- Exclusion count: {len(EXCLUSIONS)}",
        "",
        "Excluded content does not meet the selection criterion (no self-contained investment-attractiveness argument), is contact/legal/organizational, or is non-textual graphic material beyond allowed labels/captions.",
        "",
        "| Page | Text (verbatim / summary) | Reason |",
        "| --- | --- | --- |",
    ]
    for e in EXCLUSIONS:
        text = e["text"].replace("|", "\\|").replace("\n", " ")
        reason = e["reason"].replace("|", "\\|")
        lines.append(f"| {e['page']} | {text} | {reason} |")
    lines.append("")
    EXCL_OUT.write_text("\n".join(lines), encoding="utf-8")


def write_review_pack():
    residual = [u for u in UNITS if "99" in u["codes"]]
    uncertain = [u for u in UNITS if u.get("uncertainty")]
    code_counts: dict[str, int] = {}
    for unit in UNITS:
        for c in unit["codes"]:
            code_counts[c] = code_counts.get(c, 0) + 1

    lines = [
        "# Swissquote — Coding Review Pack",
        "",
        f"- Company: {META['company']}",
        f"- Source: `{META['source_url']}`",
        f"- Source Type: {META['source_type']}",
        f"- Access Date: {META['access_date']}",
        f"- Language: {META['language']}",
        f"- Document Format: {META['document_format']}",
        f"- Coder: {META['coder']}",
        f"- Coding Date: {META['coding_date']}",
        "",
        "## Counts",
        "",
        f"- Coding units: **{len(UNITS)}**",
        f"- Residual 99: **{len(residual)}**",
        f"- Uncertainty / Review Flag = 1: **{len(uncertain)}**",
        f"- Exclusions logged: **{len(EXCLUSIONS)}**",
        "",
        "## Subcategory frequency",
        "",
        "| Code | Count |",
        "| --- | ---: |",
    ]
    for code in sorted(code_counts.keys(), key=lambda x: (x == "99", [int(p) for p in x.split(".")] if x != "99" else [99])):
        lines.append(f"| {code} | {code_counts[code]} |")

    lines += [
        "",
        "## Residual 99 list",
        "",
    ]
    if not residual:
        lines.append("_None._")
    else:
        for u in residual:
            lines.append(f"- `{u['id']}` ({u['location']}): {u['statement']}")

    lines += [
        "",
        "## Uncertainty / Review Flag list",
        "",
    ]
    if not uncertain:
        lines.append("_None._")
    else:
        for u in uncertain:
            lines.append(f"- `{u['id']}` ({u['location']}; codes={','.join(u['codes'])}): {u['statement']}")
            if u.get("note"):
                lines.append(f"  - Note: {u['note']}")

    lines += [
        "",
        "## All units (summary)",
        "",
        "| ID | Loc | Codes | Statement |",
        "| --- | --- | --- | --- |",
    ]
    for u in UNITS:
        stmt = u["statement"].replace("|", "\\|").replace("\n", " ")
        if len(stmt) > 160:
            stmt = stmt[:157] + "..."
        codes = ", ".join(u["codes"])
        flag = " ⚑" if u.get("uncertainty") else ""
        lines.append(f"| {u['id']}{flag} | {u['location']} | {codes} | {stmt} |")

    lines += [
        "",
        "## Schema / process notes",
        "",
        "- Master Excel started from a **copy** of the template; template file not mutated.",
        "- Derived formula columns left untouched: Value Driver, Investment Outcome, all `* Present (derived)`.",
        "- OCR extract used: `21-swissquote-ocr-extract.txt` (especially p.13 pillars + Yuh; also p.5 tech-leadership labels, p.10 trust recognition, p.24 culture quote).",
        "- Residual 99 applied where selection criterion is met but no 9.x (or other) subcategory definition is satisfied (overall ESG ratings / generic sustainability strategy).",
        "",
    ]
    REVIEW_OUT.write_text("\n".join(lines), encoding="utf-8")


def main():
    assign_ids()
    # Validate codes and dual-coding rule (<=2 codes unless 99 alone)
    for unit in UNITS:
        assert unit["codes"], unit
        for c in unit["codes"]:
            assert c in CODE_COL, c
        if len(unit["codes"]) > 2 and "99" not in unit["codes"]:
            raise AssertionError(f"{unit.get('id','?')} has >2 codes: {unit['codes']}")

    write_excel()
    write_json()
    write_exclusion_log()
    write_review_pack()

    # Post-verify formulas untouched on written rows
    wb = load_workbook(MASTER)
    ws = wb["Coding"]
    for i in range(len(UNITS)):
        r = 3 + i
        for col in DERIVED_COLS:
            val = ws.cell(r, col).value
            assert isinstance(val, str) and val.startswith("="), (r, col, val)
        # spot-check a code cell is 0/1 int
        assert ws.cell(r, 7).value == UNITS[i]["id"]
    wb.close()

    print(
        json.dumps(
            {
                "units": len(UNITS),
                "residual_99": sum(1 for u in UNITS if "99" in u["codes"]),
                "uncertainty": sum(1 for u in UNITS if u.get("uncertainty")),
                "exclusions": len(EXCLUSIONS),
                "master": str(MASTER),
                "json": str(JSON_OUT),
                "exclusion_log": str(EXCL_OUT),
                "review_pack": str(REVIEW_OUT),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
