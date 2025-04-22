# Tenant Registration Data Cleaning & Dashboard

This project automates the cleaning and consolidation of tenant registration data using Excel VBA. It cross-references raw tenant information from Yardi and Rent Cafe, filters and prioritizes data by registration status, and prepares a clean dataset for visual analysis in Tableau.

For a more detailed look into the project, check out my [Notion case study!](https://www.notion.so/Resident-Portal-Registration-Analysis-Excel-VBA-Tableau-1d9fcf2408d080d18c8dca641c623dd2)

To have a more interactive look at my visualization, go directly to my [Dashboard!](https://public.tableau.com/app/profile/didyouseegrant/viz/RCRegistrationApril2025/Dashboard1)

---

## Features

- **Cross-references tenants and roommates** from Yardi and Rent Cafe using name and unit logic
- **Fills missing property/unit info** in Rent Cafe report using Yardi as the source of truth
- **Consolidates multiple rows per unit** into one row using prioritized status logic:
  - Keeps "Registered" if available
  - Falls back to "Invited" if no one is registered
  - Defaults to one "Unregistered" row otherwise
- **Cleans and structures final dataset** for easy import into Tableau
- **Tableau dashboard** shows registration KPIs and ratios by property

---

## Tools & Technologies

- Excel VBA  
- Tableau  


---

## Structure

```plaintext
├── tenant_registration/
│   ├── 1. raw_data/
│   │   ├── yardi_export.xlsx
│   │   └── rentcafe_export.xlsx
│   ├── 2. scripts/
│   │   └── registration_cleaning_macro.bas
│   ├── 3. cleaned_output/
│   │   └── tenant_registration_summary.xlsx
├── dashboard/
│   └── registration_dashboard.twbx
└── README.md
```

## Dashboard

The Tableau dashboard summarizes registration status across 16 properties, including:
- **KPIs** for total units, registered units, invited/unregistered counts
- Bar chart of **registration status by property**, with filterable view
- Designed to guide **outreach efforts** and improve **portal adoption**

## Disclaimer

All sample files contain **placeholder data** and modified scripts to protect sensitive company information. Values and names used in this repository are for demonstration purposes only.
