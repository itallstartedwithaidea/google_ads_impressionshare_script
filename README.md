# Google Ads Impression Share Aggregation Script

This repository contains a Google Ads Script for **aggregating impression share metrics** in a more accurate, weighted manner—matching the methodology used in the Google Ads interface. It ensures that metrics such as **Search Impression Share**, **Click Share**, **Search Exact Match Impression Share**, **Absolute Top Impression Share**, and **Top Impression Share** are aggregated properly across various campaign types.

## Features

1. **Customizable Date Range**  
   - Choose a custom date range or use common predefined ranges like `LAST_30_DAYS`, `THIS_MONTH`, etc.

2. **Accurate Aggregation**  
   - Calculates total eligible impressions (denominator) to compute realistic overall impression share figures.

3. **Campaign Filters**  
   - Easily include or exclude campaigns by name or type (`SEARCH`, `SHOPPING`, `DISPLAY`, `VIDEO`, `PERFORMANCE_MAX`).

4. **Flexible Output**  
   - Optionally specify a spreadsheet URL to output data or let the script create a new spreadsheet for each run.
   - Automatic email notification summarizing the impression share results for quick insights.

5. **Handles Multiple Metrics**  
   - Impression share metrics: 
     - **Search Impression Share** (`metrics.search_impression_share`)
     - **Search Exact Match Impression Share** (`metrics.search_exact_match_impression_share`)
     - **Search Click Share** (`metrics.search_click_share`)
     - **Search Top Impression Share** (`metrics.search_top_impression_share`)
     - **Search Absolute Top Impression Share** (`metrics.search_absolute_top_impression_share`)
   - Provides standard performance stats: impressions, clicks, cost, conversions, CTR, etc.
   - Includes additional metrics for **VIDEO** and **PERFORMANCE_MAX** campaigns (video views, engagements).

6. **Formatting & Display Options**  
   - Convert very small metrics to `< 10%` (optional).
   - Round percentages or keep them at a specific number of decimal places.

## How It Works

1. **Retrieve Campaign Data**  
   - Queries the Google Ads API (via `AdsApp.report()`) to fetch campaigns with impressions in your chosen date range.

2. **Filter & Process**  
   - Applies inclusion or exclusion filters to campaigns (e.g., partial name matches, campaign types).

3. **Compute Weighted Metrics**  
   - Uses the sum of eligible impressions or clicks to properly aggregate impression share.

4. **Output to Spreadsheet**  
   - Generates or updates a Google Spreadsheet with campaign-level results.
   - Also produces overall totals and (optionally) sends an email with a summary and spreadsheet link.

5. **Clean, Configurable Code**  
   - Easily adjust date ranges, filters, and email/spreadsheet settings in the `CONFIG` object.

---

## Contributing

Contributions, bug reports, and feature requests are welcome. Feel free to open an issue or submit a pull request.

## License

Use the license of your choice (e.g., MIT, Apache 2.0). If none is specified, it defaults to "All Rights Reserved."

