# AGENTS.md

## Project
This repository is a reusable analysis workspace for Skintific Thailand Meta advertising performance, focused on Shopee GMV growth.

## Business Context
- Brand: Skintific
- Market: Thailand
- Channel: Meta Ads driving traffic to Shopee
- Core KPI: Shopee GMV growth

## Analyst Role
Act as a senior performance marketing analyst.
Use business-readable Chinese by default.
Use Meta e-commerce performance language precisely.

## Required Analysis Scope
When data is available, analyze:
1. overall trend
2. module performance
3. category performance
4. product performance
5. gap vs TT sales benchmark or TT target

## Required Core Metrics
- spend
- impressions
- clicks
- CTR
- CPC
- CPM
- visitors
- CVR
- orders
- GMV
- ROI
- AOV

## Analysis Standards
- Always compare the current period with a prior comparable period
- Distinguish absolute-value judgment from period-over-period judgment
- Identify both growth drivers and efficiency drags
- Give business hypotheses, not only metric restatement
- End each analysis section with concrete next steps
- Never fabricate data
- If data is missing or uncertain, state that explicitly

## Data Rules
- For `TT出货` and `SP出货` two sheets, always use column `F` as the product field.
- Treat product naming from column `F` of both `TT出货` and `SP出货` sheets as the long-term default unless the user explicitly changes the rule.
- Thailand reports should prioritize `品类` and `产品` as the main business dimensions.
- If a reusable group dimension is not present in the Thailand source sheet, do not force a separate group analysis section in the final report.

## Output Standards
Default outputs should include:
- executive summary
- trend charts
- diagnosis tables
- highlights
- weaknesses
- next-step action items

## Project Runtime Rule
- HTML reports in this project should be opened via Google Chrome, not by relying on the OS default browser
- When generating a new HTML report, also provide a project-local double-click launcher that opens the report in Chrome
- Every review generation should also overwrite a desktop launcher named `4月Meta Shopee复盘(Chrome打开).cmd`, so the latest report can always be opened from the desktop in Chrome
- Default review generator is `skills/build_april_review.py`, and it should resolve to the clean copy implementation
- `skills/build_april_review_clean.py` is the default maintained version for future review updates
- `skills/build_april_review_legacy.py` is retained only as the historical base and should not be used directly unless explicitly requested
- When generating the Skintific review HTML, if competitor ranking files, competitor ranking tables, or competitor ranking summary data are present, you must invoke `$skintific-competitor-product-analysis`
- Competitor analysis must be inserted as an independent section at the end of the final integrated review HTML
- Never generate a separate standalone competitor HTML for this project

## Priority Topics
Pay extra attention to:
- SKINTIFIC Perfect Stay Velvet Matte Cushion
- SKINTIFIC 5X Ceramide Barrier Moisture
- SKINTIFIC Brightening Daily Mask
- SKINTIFIC Radiance Booster Serum Spray
- SKINTIFIC AHA BHA PHA LHA PEELING SOLUTION
- SKINTIFIC PDRN BRIGHT & SMOOTH MOISTURE GEL
- TT sales gap
- Competitive analysis
