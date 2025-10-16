# SharePoint Permissions Risk Analysis Tool

A comprehensive PowerShell solution for analyzing SharePoint permissions data and generating prioritized risk assessment reports.

## Overview

This PowerShell tool analyzes SharePoint permissions data and generates prioritized risk assessment reports with interactive HTML output.

## Features

### Risk Analysis
- **Customizable scoring system** with configurable weights
- **Interactive scoring configuration** at runtime
- **Risk categorization**: Critical (10+), High (7-9), Medium (4-6), Low (1-3), No Risk (0)
- **Professional HTML reports** with sortable columns and color-coding

### Sample Data
- Includes anonymized sample data with `contoso.sharepoint.com` domains
- Realistic data structure for testing and demonstration
- No customer-identifiable information

### Report Features
- **Sortable columns** (click headers to sort)
- **Color-coded risk levels** with visual badges
- **Search functionality** (filter by site name, URL, etc.)
- **Risk level filtering** dropdown
- **Risk distribution chart** (interactive doughnut chart)
- **Export capabilities** (JSON format)
- **Summary statistics** dashboard
- **Mobile-responsive design**

## Default Scoring Methodology

| Risk Factor | Default Score |
|-------------|---------------|
| Public Site | +3 points |
| EEEU Permissions Present | +3 points |
| Everyone Permissions Present | +3 points |
| Anyone Links Present | +2 points |
| No Sensitivity Label | +2 points |
| ≥500 Users with Access | +2 points |

## Usage

### Risk Analysis

```powershell
# Basic analysis with default scoring
.\Analyze-SharePointRisk.ps1 -CsvPath ".\your-report.csv"

# Specify custom output path
.\Analyze-SharePointRisk.ps1 -CsvPath ".\your-report.csv" -OutputPath ".\custom-report.html"
```

### Interactive Scoring

When you run the analysis script, you'll be prompted:

```
Scoring Configuration
===================
Default scoring weights:
- Public Site: +3 points
- EEEU Permissions: +3 points
- Everyone Permissions: +3 points
- Anyone Links: +2 points
- No Sensitivity Label: +2 points
- ≥500 Users: +2 points

Would you like to customize these scoring weights? (y/N): 
```

- Press **Enter** or **N** for default scoring
- Press **Y** to customize each weight interactively

## Input Data Format

The script expects a CSV file with the following columns (from SharePoint permissions export):

- `Tenant ID`
- `Site ID`
- `Site name`
- `Site URL`
- `Site template`
- `Primary admin`
- `Primary admin email`
- `External sharing`
- `Site privacy` (Public/Private/blank)
- `Site sensitivity` (sensitivity label or blank)
- `Number of users having access`
- `Guest user permissions`
- `External participant permissions`
- `Entra group count`
- `File count`
- `Items with unique permissions count`
- `PeopleInYourOrg link count`
- `Anyone link count`
- `EEEU permission count`
- `Everyone permission count`
- `Report date`

## Output

### Console Output
- Total sites analyzed
- High risk sites count
- Public sites count  
- Sites with anyone links
- Top 5 highest risk sites

### HTML Report
- Interactive table with all site data
- Risk scores and explanations
- Summary statistics dashboard with visual chart
- JSON export functionality
- Mobile-responsive design

## Sample Output

### Console Output Example
```
SharePoint Risk Analysis Tool
================================

Scoring Configuration
===================
Default scoring weights:
- Public Site: +3 points
- EEEU Permissions: +3 points
- Everyone Permissions: +3 points
- Anyone Links: +2 points
- No Sensitivity Label: +2 points
- 500+ Users: +2 points

Would you like to customize these scoring weights? (y/N): n

Loading CSV data...
Loaded 4105 sites

Analyzing risk scores...

Generating HTML report...

=== Analysis Complete ===
Report generated: .\risk_analysis.html
Total sites analyzed: 4105
High risk sites (score 7+): 156
Public sites: 189
Sites with anyone links: 0

Top 5 Highest Risk Sites:
  13 - Primary Reports
  10 - Primary Communications
  10 - Marketing Templates
  10 - Unit Finance
  10 - Common Team Site

Opening report in default browser...
```

### HTML Report Features

The generated HTML report includes:

**📊 Dashboard Summary Cards**
- Average Risk Score, Critical/High Risk Count, Highest Risk Score
- Color-coded statistics with visual indicators

**📈 Interactive Risk Distribution Chart**
- Doughnut chart showing breakdown by risk category
- Click legend items to filter data dynamically

**📋 Interactive Data Table**
```
Risk Score | Risk Level    | Site Name           | Site URL                    | Privacy | Users | Risk Factors
-----------|---------------|---------------------|-----------------------------|---------|---------|--------------
13         | Critical Risk | Primary Reports     | contoso.sharepoint.com/... | Public  | 1,247   | Public, EEEU, Everyone, No Label, 500+ Users
10         | High Risk     | Marketing Templates | contoso.sharepoint.com/... | Public  | 892     | Public, EEEU, Everyone, No Label
7          | High Risk     | Finance Dashboard   | contoso.sharepoint.com/... | Private | 634     | EEEU, Everyone, No Label, 500+ Users
4          | Medium Risk   | Team Collaboration | contoso.sharepoint.com/... | Private | 234     | Everyone, No Label
0          | No Risk       | Secure Archive      | contoso.sharepoint.com/... | Private | 12      | (none)
```

**🎨 Visual Elements**
- **Risk Badges**: Color-coded labels (Red=Critical, Orange=High, Yellow=Medium, Blue=Low, Green=No Risk)
- **Sortable Columns**: Click any header to sort (with ▲▼ indicators)
- **Search Bar**: Filter results by any text (site name, URL, etc.)
- **Risk Filter**: Dropdown to show only specific risk levels
- **Responsive Design**: Mobile-friendly layout

**⚡ Interactive Features**
- Click column headers to sort data
- Use search box to filter by keywords
- Select risk level from dropdown filter
- Export data to JSON format
- Hover effects and smooth transitions

> **💡 See it in Action**: Run the tool with the included sample data to see the full interactive HTML report:
> ```powershell
> .\Analyze-SharePointRisk.ps1 -CsvPath ".\Permissioned_Users_Count_SharePoint_report_2025-09-29_scrubbed.csv"
> ```
> The report will automatically open in your default browser showing all interactive features with real data.

## Risk Categories

| Category | Score Range | Color | Description |
|----------|-------------|-------|-------------|
| Critical Risk | 10+ | Red | Immediate attention required |
| High Risk | 7-9 | Orange | Should be reviewed soon |
| Medium Risk | 4-6 | Yellow | Monitor and plan remediation |
| Low Risk | 1-3 | Blue | Low priority for review |
| No Risk | 0 | Green | No risk factors identified |

## Requirements

- PowerShell 5.1 or later
- Windows with default browser for report viewing
- CSV export from SharePoint permissions report

## Examples

### Basic Usage
```powershell
# Analyze your SharePoint permissions data
.\Analyze-SharePointRisk.ps1 -CsvPath ".\your-sharepoint-report.csv"

# The script will prompt interactively for custom scoring if desired
```

### Custom Scoring Example
```
Would you like to customize these scoring weights? (y/N): y

=== Custom Scoring Configuration ===
Enter custom scores for each risk factor (press Enter to keep default):
Site Privacy = Public (default: 3): 5
EEEU Permissions Present (default: 3): 4
Everyone Permissions Present (default: 3): 4
Anyone Links Present (default: 2): 3
No Sensitivity Label (default: 2): 1
High User Count (default: 2): 2
User Count Threshold (default: 500): 1000
```

## Sample Data

The included sample CSV file contains anonymized data with:
- Generic `https://contoso.sharepoint.com/sites/[sitename]` URLs
- Anonymized site names (e.g., "HR Dashboard", "Marketing Portal")  
- Generic `user###@contoso.com` email addresses
- Generic admin names and tenant ID
- **Realistic numerical data** for testing the analysis tool

## License

This tool is provided as-is for SharePoint security analysis purposes.