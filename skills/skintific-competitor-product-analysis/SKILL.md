---
name: skintific-competitor-product-analysis
description: Generate the competitor product analysis section for the Skintific review HTML report.
---

# Skintific Competitor Product Analysis

## Purpose

Use this skill only for the competitor product analysis part of the main Skintific review report.

This skill does not produce a standalone report.

Its job is to generate a structured competitor-analysis section that can be inserted into the final monthly or periodic review HTML.

## Scope

Only analyze competitor product ranking data, such as ranking workbooks or extracted ranking tables.

Do not perform:

- creator content analysis
- video analysis
- OCR
- ad creative analysis
- separate competitor HTML generation

## Brand Context

The own brand is **Skintific**.

When identifying own-brand products or shops, prioritize:

1. `Is it My Shop`
2. shop names containing `Skintific`
3. product names containing `Skintific`

If multiple own-brand aliases exist, treat them as the same brand.

## Expected Inputs

Use competitor ranking workbooks or structured ranking tables containing boards such as:

- Sales
- Units
- Views

Preferably also read a `Definitions` sheet to identify:

- Region
- Category
- Time period

Useful fields may include:

- Ranking
- Change of Rank
- Product Name
- Product Link
- Product ID
- Shop Name
- Shop Link
- Is it My Shop

## Core Analysis Tasks

### 1. Cross-board product analysis

Merge products across Sales / Units / Views using:

1. Product Link
2. Product Name

Identify:

- products appearing in all three boards
- products appearing in two boards
- products appearing in only one board

### 2. Brand concentration

Summarize:

- top brands by ranked product count
- top 3 brand share
- whether the category is concentrated or fragmented
- which brands dominate shelf visibility

### 3. Title keyword trends

Extract repeated title keywords from product names.

Focus on words related to:

- ingredients
- efficacy
- texture
- format
- claim language
- bundle language

Avoid over-counting generic marketplace words.

### 4. Selling point mapping

Group repeated claims into buckets such as:

- brightening
- barrier repair
- acne care
- soothing
- moisturizing
- anti-aging
- no white cast
- lightweight texture
- oil control
- sunscreen claims
- bundle value

### 5. Traffic vs conversion structure

Identify:

- traffic-heavy products: better Views rank than Sales rank
- conversion-efficient products: better Sales rank than Views rank

Explain what this suggests about upper-funnel attraction vs lower-funnel selling power.

### 6. Skintific positioning

Summarize:

- where Skintific appears in Sales / Units / Views
- which Skintific products are carrying ranking presence
- which claims Skintific is competing on
- where Skintific is underrepresented versus competitors

## Required Output Format

This skill should output content that can be inserted into the final review HTML.

Do not output a full standalone HTML document.

Return one structured section with the following order:

1. Section title
2. Executive summary
3. Cross-board winners
4. Brand concentration
5. Keyword trends
6. Selling point distribution
7. Traffic-heavy products
8. Conversion-efficient products
9. Skintific positioning
10. Action recommendations

## Output Style

The output should be concise, structured, and presentation-ready.

Prefer:

- clear subheadings
- short insight cards
- compact tables
- direct business interpretation

Avoid generic commentary.

Each conclusion should point to a specific product, brand, title pattern, or ranking structure.

## HTML Embedding Rule

If the caller is generating HTML, write this skill's content as an embeddable section only, such as:

`<section id="competitor-product-analysis"> ... </section>`

Never generate a full document with `<html>`, `<head>`, or `<body>` unless explicitly requested.

## Final Reminder

This skill is a reusable competitor-analysis module for the main Skintific review report, not a standalone report generator.
