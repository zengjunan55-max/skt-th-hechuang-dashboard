# Skintific Competitor Product Analysis Rules

## Main Goal

Generate only the competitor product analysis section for the main Skintific review HTML.

Do not generate a separate HTML report.

## Board Types

Prefer ranking boards such as:

- Sales
- Units
- Views

## Product Merge Priority

Merge the same product using:

1. Product Link
2. Product Name

## Important Derived Fields

When possible, derive:

- sales_rank
- units_rank
- views_rank
- board_count
- board_list
- avg_rank
- best_rank
- rank_spread

## Cross-board Logic

- Products present in Sales + Units + Views are cross-board winners.
- Products present only in Views are likely traffic-oriented.
- Products present in Sales but weak in Views are likely conversion-oriented niche products.

## Brand Concentration Logic

Summarize:

- brand count
- top brand ranked slots
- top 3 brand share
- whether the market is concentrated or fragmented

## Keyword Logic

Extract repeated keywords from Product Name.

Prioritize words linked to:

- ingredient
- efficacy
- texture
- claim
- format
- bundle

Ignore generic platform words when possible.

## Selling Point Buckets

Common buckets may include:

- brightening
- acne care
- soothing
- barrier repair
- moisturizing
- anti-aging
- no white cast
- lightweight
- oil control
- sunscreen claims
- bundle value

## Skintific Own-brand Logic

Use these signals in order:

1. Is it My Shop
2. Shop Name containing Skintific
3. Product Name containing Skintific

## Required Section Order

1. Executive summary
2. Cross-board winners
3. Brand concentration
4. Keyword trends
5. Selling point distribution
6. Traffic-heavy products
7. Conversion-efficient products
8. Skintific positioning
9. Recommendations
