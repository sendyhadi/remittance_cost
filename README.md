# Remittance Cost Viewer

An interactive viewer of the cost of sending remittances, built from the World Bank's **Remittance Prices Worldwide (RPW)** database, 2016 Q2 to 2025 Q3. It covers every corridor in the database. You can filter by the UN WESP region, subregion and development status of both the receiving and the sending country.

## What it shows

**Filters.** All filters are multi-select, and an empty filter means "All".
- **Receiving country:** region, subregion, development status, country.
- **Sending country:** region, subregion, development status, country.
- **Service:** provider type and payment instrument.
- **View:** quarter and amount (USD 200 or USD 500). These are single choices.

**Charts and figures.**
- **Headline figures for the selected quarter:** average cost, year-on-year change, number of services quoted, and the cost relative to the 3% SDG target.
- **Average cost by quarter:** USD 200 and USD 500, with the 3% target as a reference line. A data table is available below the chart.
- **Breakdown chart:** group by region, subregion, development status or country, on the receiving or sending side.
- **Breakdowns by payment instrument and by provider type.**
- **Corridor table:** sender → receiver, searchable and sortable.
- **Download PNG:** each chart has a button that saves it as a PNG (2× resolution, light background). The image includes the title, the active filters, the legend and the source.
- **Download data:** each chart, and the corridor table, has a button that saves exactly what it shows as an Excel file, with an About sheet listing the filters, method and source. Built in the browser with SheetJS (loaded from cdnjs).
- **Download full dataset (Excel):** top right. All 204,469 service quotes (Q2 2016 to Q3 2025, all corridors), cleaned, with WESP region, subregion and development status for sender and receiver (`data/rpw_services_2016q2_2025q3.xlsx`, 14 MB).

With no filters, the series reproduces the World Bank's published global average: 6.36% in Q3 2025 and 6.49% in Q1 2025.

## Method

- **Source:** World Bank RPW Excel file, sheet "Dataset (from Q2 2016)".
- **Measure:** total cost (fee + FX margin) as a % of the amount sent.
- **Average:** simple average across **transparent services**, the same method as the World Bank's SDG indicator 10.c.1. It is not volume-weighted.
- **What an observation is:** a price quote for one service (provider × corridor × payment and pickup method) for a hypothetical transfer, collected by mystery shopping. These are not actual transactions.
- **Regions and status:** UN *World Economic Situation and Prospects* (WESP) Annex.
  - Development status is developed, economies in transition, or developing.
  - Where a country has no subregion, its region is used.
  - Kosovo and Tonga are not in the classification and show as "Unclassified".
- **Cleaning:** in the original file, values for 2024 Q3 and Q4 were stored as text. They are converted to numbers.
- **Q2 2025:** the World Bank published no data for this quarter, so it shows as a gap in the chart.

## Publish on GitHub Pages

1. Create a new GitHub repository, for example `remittance-cost-viewer`.
2. Upload **the contents of this folder**: `index.html`, `README.md`, and the `data/` and `scripts/` folders. `index.html` must sit at the repository root.
3. In the repository, go to **Settings → Pages**.
4. Under *Build and deployment*, choose **Source: Deploy from a branch**, branch **main**, folder **/ (root)**, and save.
5. After a minute or two the site is live at `https://<your-username>.github.io/remittance-cost-viewer/`.

You can also open `index.html` directly in a browser without a server. The data loads from `data/rpw_costs.js`.

## Updating the data

When the World Bank publishes a new quarter:

```bash
pip install pandas openpyxl xlsxwriter
python scripts/build_data.py path/to/rpw_dataset_2011_2025_qX.xlsx
```

The script regenerates `data/rpw_costs.js` and the full dataset `data/rpw_services_2016q2_2025q3.xlsx`. If the quarter range changes, rename the Excel file in the script and in `index.html` (`FULL_FILE`). To change the regional classification, edit `data/wesp_regions.csv`, which has the columns ISO, country, region, subregion and development_status.

## Files

| File | Contents |
|---|---|
| `index.html` | The viewer: HTML, CSS and JavaScript in one file, with no external libraries |
| `data/rpw_costs.js` | Aggregated data (sums and counts by quarter, sender, receiver, provider type and payment instrument) plus the WESP classification of each country |
| `data/rpw_services_2016q2_2025q3.xlsx` | Full cleaned service-level dataset behind the top-right download button |
| `data/wesp_regions.csv` | WESP classification by country: region, subregion, development status |
| `scripts/build_data.py` | Rebuilds the data from the World Bank Excel file |

## Data source and license

World Bank, *Remittance Prices Worldwide*, <https://remittanceprices.worldbank.org>. The dataset is classified as public in the [World Bank Data Catalog](https://datacatalog.worldbank.org/search/dataset/0037898/remittance-prices-worldwide), under a **Creative Commons Attribution 4.0 (CC BY 4.0)** license. Regional groupings follow the UN *World Economic Situation and Prospects* (WESP) Annex.
