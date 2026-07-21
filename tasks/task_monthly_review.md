Task: Build a monthly Meta x Shopee review.

Inputs:
- current period data
- prior comparable period data
- metric definitions from config/metrics.yaml
- dimension mappings from config/dimensions.yaml

Required steps:
1. load data
2. validate required columns
3. compute KPI summary
4. generate trend charts
5. analyze by module
6. analyze by category
7. analyze by product
8. analyze bundle performance
9. compare with TT benchmark if available
10. if competitor ranking files, competitor ranking tables, or competitor ranking summary data exist, call `$skintific-competitor-product-analysis`
11. append the competitor analysis section to the end of the final integrated review HTML
12. generate markdown summary
13. generate html report
14. save outputs to /output

Acceptance criteria:
- includes charts
- includes diagnosis tables
- includes highlights and weaknesses
- includes action items
- if competitor ranking data exists, includes a competitor analysis section at the end of the same HTML
- uses Chinese
- states missing data explicitly
- does not fabricate data
