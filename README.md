# Higher-Education-Institutions-in-France-Dataset

This repository contains a dataset of higher education institutions in France. This includes 349 higher education institutions in France, including universities,  universities of applied sciences and Higher Institutes as Higher Institute of Engineering, Higher Institute of biotechnologies and few others. This dataset was compiled in response to a cybersecurity investigation of France higher education institutions' websites [1]. The data is being made publicly available to promote open science principles [2].

## Data

The data includes the following fields for each institution:

- ETER_Id: A unique identifier assigned to each institution.
- Name: The full name of the institution.
- Category: Indicates whether the institution is public or private.
- Institution_Category_Standardized: Indicates whether the institution is University, University of applied sciences or other.
- Member_of_European_University_alliance: Indicates if the institution is member of European University Alliance (A kind of collaborative higher education institutions network in Europe).
- Url: The website of the institution.
- NUTS2: Nomenclature of Territorial Units for Statistics (NUTS): A classification by the European Union to divide member states' territories into statistical units. The NUTS system has three hierarchical levels, with NUTS2 being the second level.
- NUTS2_Label_2016: Refers to the classification of regions at the NUTS2 level according to the 2016 criteria set by the European Union.
- NUTS2_Label_2021: Refers to the classification of regions at the NUTS2 level according to the 2021 criteria set by the European Union.
- NUTS3: Nomenclature of Territorial Units for Statistics (NUTS): A classification by the European Union to divide member states' territories into statistical units. The NUTS system has three hierarchical levels, with NUTS3 being the third level.
- NUTS3_Label_2016: Refers to the classification of regions at the NUTS3 level according to the 2016 criteria set by the European Union.
- NUTS3_Label_2021: Refers to the classification of regions at the NUTS3 level according to the 2021 criteria set by the European Union.

## Methodology
The methodology for creating the dataset involved obtaining data from two sources: The European Higher Education Sector Observatory (ETER)[3]. The data was collected on December 26, 2024, the Eurostat for NUTS - Nomenclature of territorial units for statistics 2013-16[4] and 2021[5].

This section outlines the methodology used to create the dataset for Higher Education Institutions (HEIs) in France. The dataset consolidates information from various sources, processes the data, and enriches it to provide accurate and reliable insights.

**Data Sources**
1. **ETER Database**: The primary dataset was sourced from the ETER database, containing detailed information about HEIs in Europe.
   - File: `eter-export-2021-FR.xlsx`
2. **Eurostat NUTS Data**: Two datasets from Eurostat were used for regional information:
   - NUTS 2013-2016 regions: `NUTS2013-NUTS2016.xlsx`
   - NUTS 2021 regions: `NUTS2021.xlsx`

**Data Cleaning and Preprocessing**
**Column Renaming**
Columns in the raw dataset were renamed for consistency and readability. Examples include:
- `ETER ID` → `ETER_ID`
- `Institution Name` → `Name`
- `Legal status` → `Category`

**Value Replacement**
1. **HEI Categories**: The `Category` column was cleaned, with government-dependent institutions classified as "public."
2. **Standardized Institution Categories**: Mapped numerical values to descriptive labels such as "University" and "University of applied sciences."
3. **European University Alliance Membership**: Replaced binary values with "Yes" or "No."

**Handling Missing or Incorrect Data**
1. Specific entries with missing or incorrect data were updated manually based on their `ETER_ID`. For instance:
   - Adjusted URLs for entries like `FR0333` (updated to `www.icam.fr`)
   - Adjusted URLs for entries like `FR0906` (updated to `epss.fr`)
   - Adjusted URLs for entries like `FR0104` (updated to `www.ensa-nancy.fr`)
   - Adjusted URLs for entries like `FR0466` (updated to `www.clermont-auvergne-inp.fr`)
   - Adjusted URLs for entries like `FR0907` (updated to `insp.gouv.fr`) - This universety also changed your name for `Institut national du service public`
   - Removed entries such as `FR0129` and `FR0944` due to insufficient or invalid information.
   - Removed FR0513 Institut supérieur européen de gestion Lyon because it's the same url and school of Paris, so remains only the main campus in Paris
   - Remove FR0235 Institut supérieur de l'électronique et du numérique Toulon, because it's the same url of Institut supérieur de l'électronique et du numérique Lille, so remains only the main campus
   - Remove FR0106 and FR010 École spéciale militaire, because it's the url returns 403 forbiden
   - Remove FR0970 École nationale de la meteorologie  beucase of invalid HTTPS

**Regional Data Integration**
1. Merged NUTS 2016 and NUTS 2021 data to enrich the dataset with regional labels.

**Final Dataset**
The final dataset was saved as a CSV file: `france-heis.csv`, encoded in UTF-8 for compatibility. It includes detailed information about HEIs in France, their categories, regional affiliations, and membership in European alliances.

**Summary**
This methodology ensures that the dataset is accurate, consistent, and enriched with valuable regional and institutional details. The final dataset is intended to serve as a reliable resource for analyzing French HEIs.

## Usage

This data is available under the Creative Commons Zero (CC0) license and can be used for any purpose, including academic research purposes. We encourage the sharing of knowledge and the advancement of research in this field by adhering to open science principles [2].

If you use this data in your research, please cite the source and include a link to this repository. To properly attribute this data, please use the following DOI: 10.5281/zenodo.7614862

## Contribution

If you have any updates or corrections to the data, please feel free to open a pull request or contact us directly. Let's work together to keep this data accurate and up-to-date.

## Acknowledgment

We would like to acknowledge the support of the Norte Portugal Regional Operational Programme (NORTE 2020), under the PORTUGAL 2020 Partnership Agreement, through the European Regional Development Fund (ERDF), within the project "Cybers SeC IP" (NORTE-01-0145-FEDER-000044). This study was also developed as part of the Master in Cybersecurity Program at the Instituto Politécnico de Viana do Castelo, Portugal.

## References

1. Pending
2. S. Bezjak, A. Clyburne-Sherin, P. Conzett, P. Fernandes, E. Görögh, K. Helbig, B. Kramer, I. Labastida, K. Niemeyer, F. Psomopoulos, T. Ross-Hellauer, R. Schneider, J. Tennant, E. Verbakel, H. Brinken, and L. Heller, Open Science Training Handbook. Zenodo, Apr. 2018. [Online]. Available: [https://doi.org/10.5281/zenodo.1212496]
3. The European Higher Education Sector Observatory, Dec 2024. Available: [ETER](https://eter-project.com/data/data-for-download-and-visualisations/database/)
4. NUTS - Nomenclature of territorial units for statistics, Dec 2024. Available: [NUTS-2013-2016](https://ec.europa.eu/eurostat/documents/345175/629341/NUTS2013-NUTS2016.xlsx)
5. NUTS - Nomenclature of territorial units for statistics, Dec 2024. Available: [NUTS-2021](https://ec.europa.eu/eurostat/documents/345175/629341/NUTS2021.xlsx).

