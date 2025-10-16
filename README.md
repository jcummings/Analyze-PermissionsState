# 🔐 SharePoint Permissions Risk Analysis Tool

<div align="center">

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![Status](https://img.shields.io/badge/Status-Production%20Ready-brightgreen.svg)

**A comprehensive PowerShell solution for analyzing SharePoint permissions data and generating prioritized risk assessment reports.**

*Transform your SharePoint Advanced Management Data Governance reports into actionable security insights*

---

**📊 Analyze** • **🎯 Prioritize** • **📈 Visualize** • **🔍 Export**

</div>

## 📋 Table of Contents

- [🚀 Quick Start](#-quick-start)
- [📥 Data Source](#-data-source)
- [✨ Features](#-features)
- [⚙️ Risk Scoring](#️-risk-scoring)
- [🛠️ Usage](#️-usage)
- [📊 Sample Output](#-sample-output)
- [📁 Input Format](#-input-format)
- [🎨 Risk Categories](#-risk-categories)
- [📋 Requirements](#-requirements)
- [👨‍💻 Author](#-author)

## 🚀 Quick Start

```powershell
# 1. Generate SharePoint Site Permissions Report (see Data Source section)
# 2. Run analysis with default scoring
.\Analyze-SharePointRisk.ps1 -CsvPath ".\your-permissions-report.csv"

# 3. View interactive HTML report (opens automatically)
```

## 📥 Data Source

This tool is specifically designed to analyze the **SharePoint Advanced Management Data Governance "Site Permissions Report"**.

### 🔗 How to Generate the Input Data

1. **Navigate to SharePoint Advanced Management**
   - Go to the SharePoint Online Admin Center
   - Access SharePoint Advanced Management

2. **Generate Site Permissions Report**
   - Navigate to **Data access governance** → **Reports**
   - Select **"Site permissions report"**
   - Configure your report parameters and generate

3. **Download CSV Export**
   - Once generated, download the CSV file
   - This CSV file is your input for this analysis tool

📖 **Detailed Instructions**: [Microsoft Learn - SharePoint Data Access Governance Reports](https://learn.microsoft.com/en-us/sharepoint/data-access-governance-reports)

> ⚠️ **Important**: This tool requires the specific CSV format from SharePoint Advanced Management's Site Permissions Report. Other SharePoint exports may not contain the required columns.

## ✨ Features

<table>
<tr>
<td width="50%">

### 🎯 **Risk Analysis**
- 🔧 **Customizable scoring system** with configurable weights
- ⚡ **Interactive scoring configuration** at runtime
- 🏷️ **Risk categorization**: Critical (10+), High (7-9), Medium (4-6), Low (1-3), No Risk (0)
- 📄 **Professional HTML reports** with sortable columns and color-coding

### 📊 **Visual Dashboard**
- 🎨 **Color-coded risk levels** with visual badges
- 🔍 **Search functionality** (filter by site name, URL, etc.)
- 📋 **Risk level filtering** dropdown
- 📈 **Risk distribution chart** (interactive doughnut chart)
- 📱 **Mobile-responsive design**

</td>
<td width="50%">

### 🗂️ **Sample Data**
- 🏢 Includes anonymized sample data with `contoso.sharepoint.com` domains
- 📈 Realistic data structure for testing and demonstration
- 🔒 No customer-identifiable information

### 📤 **Export & Reporting**
- 💾 **Export capabilities** (JSON format)
- 📊 **Summary statistics** dashboard
- 📋 **Interactive table** with all site data
- 🎛️ **Risk scores and explanations**

</td>
</tr>
</table>

## ⚙️ Default Scoring Methodology

<div align="center">

| 🚨 Risk Factor | 📊 Default Score | 💡 Why It Matters |
|----------------|:----------------:|-------------------|
| 🌐 **Public Site** | **+3 points** | Visible to anyone on the internet |
| 👥 **EEEU Permissions Present** | **+3 points** | All organization users can access |
| 🔓 **Everyone Permissions Present** | **+3 points** | External users may have access |
| 🔗 **Anyone Links Present** | **+2 points** | Anonymous sharing links exist |
| 🏷️ **No Sensitivity Label** | **+2 points** | Missing data classification |
| 📈 **≥500 Users with Access** | **+2 points** | Large user base increases risk |

</div>

> 💡 **Customizable**: All scoring weights can be adjusted interactively when running the tool!

## 🛠️ Usage

### 🚀 **Risk Analysis**

```powershell
# 📊 Basic analysis with default scoring
.\Analyze-SharePointRisk.ps1 -CsvPath ".\your-permissions-report.csv"

# 🎯 Specify custom output path
.\Analyze-SharePointRisk.ps1 -CsvPath ".\your-permissions-report.csv" -OutputPath ".\custom-report.html"
```

### ⚙️ **Interactive Scoring Configuration**

When you run the analysis script, you'll be prompted:

<details>
<summary>📋 <strong>Click to expand scoring configuration example</strong></summary>

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

- 💚 Press **Enter** or **N** for default scoring
- 🎛️ Press **Y** to customize each weight interactively

</details>

## 📁 Input Data Format

> 📋 **Required Source**: SharePoint Advanced Management - Site Permissions Report CSV

<details>
<summary>📊 <strong>Click to view required CSV columns</strong></summary>

The script expects a CSV file with the following columns (from SharePoint Advanced Management permissions export):

- 🆔 `Tenant ID`
- 🆔 `Site ID`
- 📝 `Site name`
- 🔗 `Site URL`
- 📄 `Site template`
- 👤 `Primary admin`
- 📧 `Primary admin email`
- 🔄 `External sharing`
- 🔒 `Site privacy` (Public/Private/blank)
- 🏷️ `Site sensitivity` (sensitivity label or blank)
- 👥 `Number of users having access`
- 👤 `Guest user permissions`
- 🌐 `External participant permissions`
- 📊 `Entra group count`
- 📁 `File count`
- 🔐 `Items with unique permissions count`
- 🔗 `PeopleInYourOrg link count`
- 🔗 `Anyone link count`
- 🔓 `EEEU permission count`
- 🔓 `Everyone permission count`
- 📅 `Report date`

</details>

## 📊 Sample Output

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

## 📊 Sample Output

### 💻 **Console Output Example**

<details>
<summary>🖥️ <strong>Click to view console output</strong></summary>

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

</details>

### 🎨 **HTML Report Features**

The generated HTML report includes:

**📊 Dashboard Summary Cards**
- Average Risk Score, Critical/High Risk Count, Highest Risk Score
- Color-coded statistics with visual indicators

**📈 Interactive Risk Distribution Chart**
- Doughnut chart showing breakdown by risk category
- Click legend items to filter data dynamically

**📋 Interactive Data Table**

<details>
<summary>📊 <strong>Sample table data preview</strong></summary>

```
Risk Score | Risk Level    | Site Name           | Site URL                    | Privacy | Users | Risk Factors
-----------|---------------|---------------------|-----------------------------|---------|---------|--------------
13         | Critical Risk | Primary Reports     | contoso.sharepoint.com/... | Public  | 1,247   | Public, EEEU, Everyone, No Label, 500+ Users
10         | High Risk     | Marketing Templates | contoso.sharepoint.com/... | Public  | 892     | Public, EEEU, Everyone, No Label
7          | High Risk     | Finance Dashboard   | contoso.sharepoint.com/... | Private | 634     | EEEU, Everyone, No Label, 500+ Users
4          | Medium Risk   | Team Collaboration | contoso.sharepoint.com/... | Private | 234     | Everyone, No Label
0          | No Risk       | Secure Archive      | contoso.sharepoint.com/... | Private | 12      | (none)
```

</details>

**🎨 Visual Elements**
- 🏷️ **Risk Badges**: Color-coded labels (🔴 Critical, 🟠 High, 🟡 Medium, 🔵 Low, 🟢 No Risk)
- 🔽 **Sortable Columns**: Click any header to sort (with ▲▼ indicators)
- 🔍 **Search Bar**: Filter results by any text (site name, URL, etc.)
- 📋 **Risk Filter**: Dropdown to show only specific risk levels
- 📱 **Responsive Design**: Mobile-friendly layout

**⚡ Interactive Features**
- 🔄 Click column headers to sort data
- 🔍 Use search box to filter by keywords
- 📊 Select risk level from dropdown filter
- 💾 Export data to JSON format
- ✨ Hover effects and smooth transitions

> **💡 See it in Action**: Run the tool with the included sample data to see the full interactive HTML report:
> ```powershell
> .\Analyze-SharePointRisk.ps1 -CsvPath ".\Permissioned_Users_Count_SharePoint_report_2025-09-29_scrubbed.csv"
> ```
> The report will automatically open in your default browser showing all interactive features with real data.

## 🎨 Risk Categories

<div align="center">

| 🏷️ Category | 📊 Score Range | ⚠️ Priority Level | 📝 Description |
|-------------|:-------------:|:----------------:|----------------|
| 🔴 **Critical Risk** | **10+** | 🚨 **URGENT** | Immediate attention required |
| 🟠 **High Risk** | **7-9** | ⚡ **HIGH** | Should be reviewed soon |
| 🟡 **Medium Risk** | **4-6** | 📋 **MEDIUM** | Monitor and plan remediation |
| 🔵 **Low Risk** | **1-3** | 📝 **LOW** | Low priority for review |
| 🟢 **No Risk** | **0** | ✅ **SAFE** | No risk factors identified |

</div>

## 📋 Requirements

- 💻 **PowerShell 5.1 or later**
- 🪟 **Windows** with default browser for report viewing
- 📊 **SharePoint Advanced Management** - Site Permissions Report CSV

## 🚀 Examples

### 🔧 **Basic Usage**
```powershell
# 📊 Analyze your SharePoint permissions data
.\Analyze-SharePointRisk.ps1 -CsvPath ".\your-sharepoint-permissions-report.csv"

# 🎛️ The script will prompt interactively for custom scoring if desired
```

### ⚙️ **Custom Scoring Example**

<details>
<summary>🎛️ <strong>Advanced configuration example</strong></summary>

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

</details>

## 🗂️ Sample Data

The included sample CSV file contains anonymized data with:
- 🔗 Generic `https://contoso.sharepoint.com/sites/[sitename]` URLs
- 🏢 Anonymized site names (e.g., "HR Dashboard", "Marketing Portal")  
- 📧 Generic `user###@contoso.com` email addresses
- 👤 Generic admin names and tenant ID
- 📊 **Realistic numerical data** for testing the analysis tool

## 👨‍💻 Author

<div align="center">

**John Cummings**  
📧 [john@jcummings.net](mailto:john@jcummings.net)  
📅 Published: October 16, 2025

---

*Built with ❤️ for SharePoint security professionals*

</div>

## 📄 License

This tool is provided as-is for SharePoint security analysis purposes under the MIT License.

---

<div align="center">

### 🛡️ Security • 📊 Analytics • 🚀 Efficiency

**Star ⭐ this repository if it helped you secure your SharePoint environment!**

[![GitHub issues](https://img.shields.io/github/issues/jcummings/Analyze-PermissionsState)](https://github.com/jcummings/Analyze-PermissionsState/issues)
[![GitHub stars](https://img.shields.io/github/stars/jcummings/Analyze-PermissionsState)](https://github.com/jcummings/Analyze-PermissionsState/stargazers)

</div>