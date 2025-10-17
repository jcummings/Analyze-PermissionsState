<#
.SYNOPSIS
    SharePoint Permissions Risk Analysis Tool
    
.DESCRIPTION
    Analyzes SharePoint permissions data from CSV export and generates a prioritized risk assessment report.
    
.PARAMETER CsvPath
    Path to the SharePoint permissions CSV file
    
.PARAMETER OutputPath
    Path for the output HTML report (optional - defaults to same directory as input)
    
.EXAMPLE
    .\Analyze-SharePointRisk.ps1 -CsvPath ".\report.csv"
    
.EXAMPLE
    .\Analyze-SharePointRisk.ps1 -CsvPath ".\report.csv" -OutputPath ".\custom-report.html"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Default scoring configuration
$DefaultScoring = @{
    PublicSite = 3
    EEEUPermissions = 3
    EveryonePermissions = 3
    AnyoneLinks = 2
    NoSensitivityLabel = 2
    HighUserCount = 2
    UserCountThreshold = 500
}

# Risk level definitions
$RiskLevels = @{
    Critical = @{ Min = 10; Max = 999; Color = "#dc3545"; Label = "Critical Risk" }
    High = @{ Min = 7; Max = 9; Color = "#fd7e14"; Label = "High Risk" }
    Medium = @{ Min = 4; Max = 6; Color = "#ffc107"; Label = "Medium Risk" }
    Low = @{ Min = 1; Max = 3; Color = "#17a2b8"; Label = "Low Risk" }
    NoRisk = @{ Min = 0; Max = 0; Color = "#28a745"; Label = "No Risk" }
}

function Get-ScoringPreference {
    <#
    .SYNOPSIS
    Asks user if they want to use custom scoring
    #>
    
    Write-Host "`nScoring Configuration" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    Write-Host "Default scoring weights:" -ForegroundColor Yellow
    Write-Host "- Public Site: +$($DefaultScoring.PublicSite) points" -ForegroundColor White
    Write-Host "- EEEU Permissions: +$($DefaultScoring.EEEUPermissions) points" -ForegroundColor White
    Write-Host "- Everyone Permissions: +$($DefaultScoring.EveryonePermissions) points" -ForegroundColor White
    Write-Host "- Anyone Links: +$($DefaultScoring.AnyoneLinks) points" -ForegroundColor White
    Write-Host "- No Sensitivity Label: +$($DefaultScoring.NoSensitivityLabel) points" -ForegroundColor White
    Write-Host "- $($DefaultScoring.UserCountThreshold)+ Users: +$($DefaultScoring.HighUserCount) points" -ForegroundColor White
    
    do {
        $response = Read-Host "`nWould you like to customize these scoring weights? (y/N)"
        $response = $response.Trim().ToLower()
        if ($response -eq '' -or $response -eq 'n' -or $response -eq 'no') {
            return $false
        } elseif ($response -eq 'y' -or $response -eq 'yes') {
            return $true
        } else {
            Write-Host "Please enter 'y' for yes or 'n' for no (or press Enter for default)." -ForegroundColor Red
        }
    } while ($true)
}

function Get-CustomScoring {
    <#
    .SYNOPSIS
    Prompts user for custom scoring configuration
    #>
    
    Write-Host "`n=== Custom Scoring Configuration ===" -ForegroundColor Cyan
    Write-Host "Enter custom scores for each risk factor (press Enter to keep default):" -ForegroundColor Yellow
    
    $customScoring = $DefaultScoring.Clone()
    
    $prompt = "Site Privacy = Public (default: $($DefaultScoring.PublicSite)): "
    $input = Read-Host $prompt
    if ($input -and $input -match '^\d+$') { $customScoring.PublicSite = [int]$input }
    
    $prompt = "EEEU Permissions Present (default: $($DefaultScoring.EEEUPermissions)): "
    $input = Read-Host $prompt
    if ($input -and $input -match '^\d+$') { $customScoring.EEEUPermissions = [int]$input }
    
    $prompt = "Everyone Permissions Present (default: $($DefaultScoring.EveryonePermissions)): "
    $input = Read-Host $prompt
    if ($input -and $input -match '^\d+$') { $customScoring.EveryonePermissions = [int]$input }
    
    $prompt = "Anyone Links Present (default: $($DefaultScoring.AnyoneLinks)): "
    $input = Read-Host $prompt
    if ($input -and $input -match '^\d+$') { $customScoring.AnyoneLinks = [int]$input }
    
    $prompt = "No Sensitivity Label (default: $($DefaultScoring.NoSensitivityLabel)): "
    $input = Read-Host $prompt
    if ($input -and $input -match '^\d+$') { $customScoring.NoSensitivityLabel = [int]$input }
    
    $prompt = "High User Count (default: $($DefaultScoring.HighUserCount)): "
    $input = Read-Host $prompt
    if ($input -and $input -match '^\d+$') { $customScoring.HighUserCount = [int]$input }
    
    $prompt = "User Count Threshold (default: $($DefaultScoring.UserCountThreshold)): "
    $input = Read-Host $prompt
    if ($input -and $input -match '^\d+$') { $customScoring.UserCountThreshold = [int]$input }
    
    Write-Host "`nCustom scoring configuration applied!" -ForegroundColor Green
    return $customScoring
}

function Calculate-RiskScore {
    <#
    .SYNOPSIS
    Calculates risk score for a SharePoint site
    #>
    param(
        [PSCustomObject]$Site,
        [hashtable]$ScoringConfig
    )
    
    $score = 0
    $reasons = @()
    
    # Site Privacy = Public (+points)
    if ($Site.'Site privacy' -eq 'Public') {
        $score += $ScoringConfig.PublicSite
        $reasons += "Public site (+$($ScoringConfig.PublicSite))"
    }
    
    # EEEU Permissions Present (+points)
    if ($Site.'EEEU permission count' -gt 0) {
        $score += $ScoringConfig.EEEUPermissions
        $reasons += "EEEU permissions ($($Site.'EEEU permission count')) (+$($ScoringConfig.EEEUPermissions))"
    }
    
    # Everyone Permissions Present (+points)
    if ($Site.'Everyone permission count' -gt 0) {
        $score += $ScoringConfig.EveryonePermissions
        $reasons += "Everyone permissions ($($Site.'Everyone permission count')) (+$($ScoringConfig.EveryonePermissions))"
    }
    
    # Anyone Links Present (+points)
    if ($Site.'Anyone link count' -gt 0) {
        $score += $ScoringConfig.AnyoneLinks
        $reasons += "Anyone links ($($Site.'Anyone link count')) (+$($ScoringConfig.AnyoneLinks))"
    }
    
    # No Sensitivity Label (+points)
    if ([string]::IsNullOrWhiteSpace($Site.'Site sensitivity')) {
        $score += $ScoringConfig.NoSensitivityLabel
        $reasons += "No sensitivity label (+$($ScoringConfig.NoSensitivityLabel))"
    }
    
    # High User Count (+points)
    if ($Site.'Number of users having access' -ge $ScoringConfig.UserCountThreshold) {
        $score += $ScoringConfig.HighUserCount
        $reasons += "High user count ($($Site.'Number of users having access')) (+$($ScoringConfig.HighUserCount))"
    }
    
    return @{
        Score = $score
        Reasons = $reasons -join '; '
    }
}

function Get-RiskCategory {
    <#
    .SYNOPSIS
    Determines risk category based on score
    #>
    param([int]$Score)
    
    foreach ($level in $RiskLevels.GetEnumerator()) {
        if ($Score -ge $level.Value.Min -and $Score -le $level.Value.Max) {
            return @{
                Category = $level.Key
                Label = $level.Value.Label
                Color = $level.Value.Color
            }
        }
    }
    
    # Default to Critical if score is very high
    return @{
        Category = "Critical"
        Label = "Critical Risk"
        Color = "#dc3545"
    }
}

function Generate-HtmlReport {
    <#
    .SYNOPSIS
    Generates HTML report with risk analysis
    #>
    param(
        [array]$AnalyzedSites,
        [hashtable]$ScoringConfig,
        [hashtable]$Statistics,
        [string]$OutputPath
    )
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SharePoint Risk Analysis Report</title>
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background-color: #f8f9fa; 
        }
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            background: white; 
            border-radius: 10px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
            padding: 30px; 
        }
        h1 { 
            color: #343a40; 
            text-align: center; 
            margin-bottom: 30px; 
            border-bottom: 3px solid #007bff; 
            padding-bottom: 15px; 
        }
        .summary { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
            gap: 20px; 
            margin-bottom: 30px; 
        }
        .summary-card { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
            color: white; 
            padding: 20px; 
            border-radius: 10px; 
            text-align: center; 
            box-shadow: 0 4px 8px rgba(0,0,0,0.1); 
        }
        .summary-card.high-risk { 
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); 
        }
        .summary-card.public-sites { 
            background: linear-gradient(135deg, #fd7e14 0%, #e8690b 100%); 
        }
        .summary-card.anyone-links { 
            background: linear-gradient(135deg, #ffc107 0%, #e0a800 100%); 
            color: #212529; 
        }
        .summary-card.total-sites { 
            background: linear-gradient(135deg, #28a745 0%, #1e7e34 100%); 
        }
        .summary-card h3 { 
            margin: 0 0 10px 0; 
            font-size: 2em; 
        }
        .risk-distribution {
            background-color: #f8f9fa;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            border: 1px solid #e9ecef;
        }
        .risk-distribution h2 {
            color: #495057;
            margin-top: 0;
            margin-bottom: 20px;
            text-align: center;
        }
        .chart-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
            align-items: center;
        }
        .chart-wrapper {
            position: relative;
            height: 300px;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .risk-stats {
            display: flex;
            flex-direction: column;
            gap: 20px;
        }
        .risk-stat-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 15px 20px;
            background: white;
            border-radius: 8px;
            border-left: 4px solid #007bff;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .risk-stat-label {
            font-weight: 600;
            color: #495057;
        }
        .risk-stat-value {
            font-size: 1.5em;
            font-weight: bold;
            color: #007bff;
        }
        @media (max-width: 968px) {
            .chart-container {
                grid-template-columns: 1fr;
                gap: 20px;
            }
        }
        .summary-card p { 
            margin: 0; 
            opacity: 0.9; 
        }
        .scoring-info { 
            background-color: #e9ecef; 
            padding: 20px; 
            border-radius: 10px; 
            margin-bottom: 30px; 
        }
        .scoring-info h2 { 
            color: #495057; 
            margin-top: 0; 
        }
        .controls { 
            margin-bottom: 20px; 
            display: flex; 
            gap: 15px; 
            align-items: center; 
            flex-wrap: wrap; 
        }
        .controls input, .controls select { 
            padding: 8px 12px; 
            border: 1px solid #ddd; 
            border-radius: 5px; 
            font-size: 14px; 
        }
        .controls button { 
            background-color: #007bff; 
            color: white; 
            border: none; 
            padding: 10px 20px; 
            border-radius: 5px; 
            cursor: pointer; 
            font-size: 14px; 
        }
        .controls button:hover { 
            background-color: #0056b3; 
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 20px; 
            font-size: 14px; 
        }
        th, td { 
            padding: 12px; 
            text-align: left; 
            border-bottom: 1px solid #ddd; 
        }
        th { 
            background-color: #343a40; 
            color: white; 
            cursor: pointer; 
            user-select: none; 
            position: sticky; 
            top: 0; 
            z-index: 100; 
        }
        th:hover { 
            background-color: #495057; 
        }
        th.sort-asc::after {
            content: ' ^';
            font-weight: bold;
        }
        th.sort-desc::after {
            content: ' v';
            font-weight: bold;
        }
        tr:nth-child(even) { 
            background-color: #f8f9fa; 
        }
        tr:hover { 
            background-color: #e3f2fd; 
        }
        .risk-badge { 
            padding: 4px 12px; 
            border-radius: 20px; 
            color: white; 
            font-weight: bold; 
            font-size: 12px; 
            text-align: center; 
            display: inline-block; 
            min-width: 80px; 
        }
        .risk-reasons { 
            font-size: 12px; 
            color: #666; 
            max-width: 300px; 
            word-wrap: break-word; 
        }
        .site-url { 
            color: #007bff; 
            text-decoration: none; 
            max-width: 250px; 
            word-wrap: break-word; 
            display: block; 
        }
        .site-url:hover { 
            text-decoration: underline; 
        }
        .legend { 
            display: flex; 
            gap: 15px; 
            margin-bottom: 20px; 
            flex-wrap: wrap; 
        }
        .legend-item { 
            display: flex; 
            align-items: center; 
            gap: 5px; 
        }
        .legend-color { 
            width: 20px; 
            height: 20px; 
            border-radius: 10px; 
        }
        @media (max-width: 768px) {
            .container { 
                padding: 15px; 
            }
            .controls { 
                flex-direction: column; 
                align-items: stretch; 
            }
            table { 
                font-size: 12px; 
            }
            th, td { 
                padding: 8px; 
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>SharePoint Risk Analysis Report</h1>
        
        <div class="summary">
            <div class="summary-card total-sites">
                <h3>$($Statistics.TotalSites)</h3>
                <p>Total Sites Analyzed</p>
            </div>
            <div class="summary-card high-risk">
                <h3>$($Statistics.HighRiskSites)</h3>
                <p>High Risk Sites (7+ Score)</p>
            </div>
            <div class="summary-card public-sites">
                <h3>$($Statistics.PublicSites)</h3>
                <p>Public Sites</p>
            </div>
            <div class="summary-card anyone-links">
                <h3>$($Statistics.SitesWithAnyoneLinks)</h3>
                <p>Sites with Anyone Links</p>
            </div>
        </div>
        
        <div class="risk-distribution">
            <h2>Risk Distribution Analysis</h2>
            <div class="chart-container">
                <div class="chart-wrapper">
                    <canvas id="riskChart"></canvas>
                </div>
                <div class="risk-stats">
                    <div class="risk-stat-item">
                        <span class="risk-stat-label">Average Risk Score:</span>
                        <span class="risk-stat-value" id="avgRiskScore">0</span>
                    </div>
                    <div class="risk-stat-item">
                        <span class="risk-stat-label">Critical + High Risk:</span>
                        <span class="risk-stat-value" id="criticalHighRisk">0</span>
                    </div>
                    <div class="risk-stat-item">
                        <span class="risk-stat-label">Highest Risk Score:</span>
                        <span class="risk-stat-value" id="highestRiskScore">0</span>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="scoring-info">
            <h2>Scoring Methodology</h2>
            <p><strong>Risk factors and their weights:</strong></p>
            <ul>
                <li>Public Site: +$($ScoringConfig.PublicSite) points</li>
                <li>EEEU Permissions Present: +$($ScoringConfig.EEEUPermissions) points</li>
                <li>Everyone Permissions Present: +$($ScoringConfig.EveryonePermissions) points</li>
                <li>Anyone Links Present: +$($ScoringConfig.AnyoneLinks) points</li>
                <li>No Sensitivity Label: +$($ScoringConfig.NoSensitivityLabel) points</li>
                <li>$($ScoringConfig.UserCountThreshold)+ Users with Access: +$($ScoringConfig.HighUserCount) points</li>
            </ul>
        </div>
        
        <div class="legend">
            <div class="legend-item">
                <div class="legend-color" style="background-color: #dc3545;"></div>
                <span>Critical Risk (10+)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color" style="background-color: #fd7e14;"></div>
                <span>High Risk (7-9)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color" style="background-color: #ffc107;"></div>
                <span>Medium Risk (4-6)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color" style="background-color: #17a2b8;"></div>
                <span>Low Risk (1-3)</span>
            </div>
            <div class="legend-item">
                <div class="legend-color" style="background-color: #28a745;"></div>
                <span>No Risk (0)</span>
            </div>
        </div>
        
        <div class="controls">
            <input type="text" id="searchBox" placeholder="Search sites..." />
            <select id="riskFilter" onchange="filterByRisk(this.value)">
                <option value="">All Risk Levels</option>
                <option value="Critical">Critical Risk</option>
                <option value="High">High Risk</option>
                <option value="Medium">Medium Risk</option>
                <option value="Low">Low Risk</option>
                <option value="NoRisk">No Risk</option>
            </select>
            <button onclick="exportToJSON()">Export to JSON</button>
            <button onclick="exportToCSV()">Export to CSV</button>
        </div>
        
        <table id="riskTable">
            <thead>
                <tr>
                    <th onclick="sortTable(0)" id="riskScoreHeader">Risk Score</th>
                    <th onclick="sortTable(1)">Risk Level</th>
                    <th onclick="sortTable(2)">Site Name</th>
                    <th onclick="sortTable(3)">Site URL</th>
                    <th onclick="sortTable(4)">Privacy</th>
                    <th onclick="sortTable(5)">Users</th>
                    <th onclick="sortTable(6)">Anyone Links</th>
                    <th onclick="sortTable(7)">EEEU Perms</th>
                    <th onclick="sortTable(8)">Everyone Perms</th>
                    <th onclick="sortTable(9)">Risk Factors</th>
                </tr>
            </thead>
            <tbody>
"@
    
    foreach ($site in $AnalyzedSites) {
        $riskInfo = Get-RiskCategory -Score $site.RiskScore
        $privacy = if ([string]::IsNullOrWhiteSpace($site.'Site privacy')) { "Not Set" } else { $site.'Site privacy' }
        $siteName = if ([string]::IsNullOrWhiteSpace($site.'Site name')) { "Unnamed Site" } else { $site.'Site name' }
        
        $html += @"
                <tr data-risk="$($riskInfo.Category)">
                    <td style="font-weight: bold; font-size: 16px;">$($site.RiskScore)</td>
                    <td><span class="risk-badge" style="background-color: $($riskInfo.Color);">$($riskInfo.Label)</span></td>
                    <td>$(ConvertTo-HtmlSafe $siteName)</td>
                    <td><a href="$($site.'Site URL')" target="_blank" class="site-url">$(ConvertTo-HtmlSafe $site.'Site URL')</a></td>
                    <td>$privacy</td>
                    <td>$($site.'Number of users having access')</td>
                    <td>$($site.'Anyone link count')</td>
                    <td>$($site.'EEEU permission count')</td>
                    <td>$($site.'Everyone permission count')</td>
                    <td class="risk-reasons">$(ConvertTo-HtmlSafe $site.RiskReasons)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        let originalData = [];
        let riskChart = null;
        
        // Store original table data and initialize chart
        document.addEventListener('DOMContentLoaded', function() {
            const table = document.getElementById('riskTable');
            const tbody = table.querySelector('tbody');
            originalData = Array.from(tbody.querySelectorAll('tr')).map(row => ({
                element: row.cloneNode(true),
                text: row.textContent.toLowerCase()
            }));
            
            // Set up search functionality
            document.getElementById('searchBox').addEventListener('input', filterTable);
            document.getElementById('riskFilter').addEventListener('change', filterTable);
            
            // Initialize risk distribution chart
            initializeRiskChart();
        });
        
        function initializeRiskChart() {
            const table = document.getElementById('riskTable');
            const rows = table.querySelectorAll('tbody tr');
            
            // Count risk levels
            let riskCounts = {
                'Critical': 0,
                'High': 0, 
                'Medium': 0,
                'Low': 0,
                'No Risk': 0
            };
            
            let totalScore = 0;
            let maxScore = 0;
            
            rows.forEach(row => {
                const riskScore = parseInt(row.cells[0].textContent.trim());
                const riskCategory = row.getAttribute('data-risk');
                
                totalScore += riskScore;
                maxScore = Math.max(maxScore, riskScore);
                
                // Map data-risk attributes to display names
                if (riskCategory === 'Critical') riskCounts.Critical++;
                else if (riskCategory === 'High') riskCounts.High++;
                else if (riskCategory === 'Medium') riskCounts.Medium++;
                else if (riskCategory === 'Low') riskCounts.Low++;
                else riskCounts['No Risk']++;
            });
            
            const avgScore = rows.length > 0 ? (totalScore / rows.length).toFixed(1) : 0;
            const criticalHighCount = riskCounts.Critical + riskCounts.High;
            
            // Update statistics
            document.getElementById('avgRiskScore').textContent = avgScore;
            document.getElementById('criticalHighRisk').textContent = criticalHighCount.toLocaleString();
            document.getElementById('highestRiskScore').textContent = maxScore;
            
            // Create chart
            const ctx = document.getElementById('riskChart').getContext('2d');
            
            riskChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['Critical Risk', 'High Risk', 'Medium Risk', 'Low Risk', 'No Risk'],
                    datasets: [{
                        data: [riskCounts.Critical, riskCounts.High, riskCounts.Medium, riskCounts.Low, riskCounts['No Risk']],
                        backgroundColor: ['#dc3545', '#fd7e14', '#ffc107', '#17a2b8', '#28a745'],
                        borderColor: ['#c82333', '#e8690b', '#e0a800', '#117a8b', '#1e7e34'],
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                padding: 20,
                                usePointStyle: true
                            }
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = ((context.raw / total) * 100).toFixed(1);
                                    return context.label + ': ' + context.raw.toLocaleString() + ' (' + percentage + '%)';
                                }
                            }
                        }
                    }
                }
            });
        }
        
        function sortTable(columnIndex) {
            const table = document.getElementById('riskTable');
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            
            // Get the header element
            const header = table.querySelectorAll('th')[columnIndex];
            
            // Determine current sort state based on CSS classes
            const isCurrentlyDescending = header.classList.contains('sort-desc');
            const isCurrentlyAscending = header.classList.contains('sort-asc');
            
            // Determine new sort direction
            let newDescending;
            if (isCurrentlyDescending) {
                newDescending = false; // Switch to ascending
            } else if (isCurrentlyAscending) {
                newDescending = true; // Switch to descending
            } else {
                // No current sort indicator - use default for column
                newDescending = (columnIndex === 0) ? true : false; // Risk score defaults to descending, others to ascending
            }
            
            // Reset all headers - remove sort classes
            table.querySelectorAll('th').forEach(th => {
                th.classList.remove('sort-asc', 'sort-desc');
            });
            
            // Sort rows
            rows.sort((a, b) => {
                let aVal = a.cells[columnIndex].textContent.trim();
                let bVal = b.cells[columnIndex].textContent.trim();
                
                // Handle numeric columns
                if (columnIndex === 0 || columnIndex === 5 || columnIndex === 6 || columnIndex === 7 || columnIndex === 8) {
                    aVal = parseFloat(aVal) || 0;
                    bVal = parseFloat(bVal) || 0;
                }
                
                if (aVal < bVal) return newDescending ? 1 : -1;
                if (aVal > bVal) return newDescending ? -1 : 1;
                return 0;
            });
            
            // Update header with sort indicator using CSS class
            header.classList.add(newDescending ? 'sort-desc' : 'sort-asc');
            
            // Re-append sorted rows
            rows.forEach(row => tbody.appendChild(row));
        }
        
        function filterTable() {
            const searchTerm = document.getElementById('searchBox').value.toLowerCase();
            const riskFilter = document.getElementById('riskFilter').value;
            const tbody = document.querySelector('#riskTable tbody');
            
            // Clear current rows
            tbody.innerHTML = '';
            
            // Filter and display matching rows
            originalData.forEach(item => {
                const matchesSearch = !searchTerm || item.text.includes(searchTerm);
                const matchesRisk = !riskFilter || item.element.getAttribute('data-risk') === riskFilter;
                
                if (matchesSearch && matchesRisk) {
                    tbody.appendChild(item.element.cloneNode(true));
                }
            });
        }
        
        function exportToJSON() {
            const table = document.getElementById('riskTable');
            const rows = table.querySelectorAll('tbody tr');
            const headers = Array.from(table.querySelectorAll('th')).map(th => 
                th.textContent.replace(/[▲▼]/, '').trim()
            );
            
            const data = Array.from(rows).map(row => {
                const cells = row.querySelectorAll('td');
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header] = cells[index] ? cells[index].textContent.trim() : '';
                });
                return obj;
            });
            
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'sharepoint-risk-analysis.json';
            a.click();
            window.URL.revokeObjectURL(url);
        }
        
        function exportToCSV() {
            const table = document.getElementById('riskTable');
            const rows = table.querySelectorAll('tbody tr');
            const headers = Array.from(table.querySelectorAll('th')).map(th => 
                th.textContent.replace(/[↑↓▲▼]/g, '').trim()
            );
            
            // Create CSV content
            let csvContent = headers.join(',') + '\n';
            
            Array.from(rows).forEach(row => {
                const cells = row.querySelectorAll('td');
                const rowData = Array.from(cells).map(cell => {
                    let text = cell.textContent.trim();
                    // Escape quotes and wrap in quotes if contains comma
                    if (text.includes(',') || text.includes('"') || text.includes('\n')) {
                        text = '"' + text.replace(/"/g, '""') + '"';
                    }
                    return text;
                });
                csvContent += rowData.join(',') + '\n';
            });
            
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'sharepoint-risk-analysis.csv';
            a.click();
            window.URL.revokeObjectURL(url);
        }
        
        // Initial sort indicator (data is already sorted by PowerShell)
        document.addEventListener('DOMContentLoaded', function() {
            setTimeout(() => {
                // Set initial sort indicator for Risk Score column
                const riskScoreHeader = document.getElementById('riskScoreHeader');
                if (riskScoreHeader) {
                    riskScoreHeader.classList.add('sort-desc'); // Down arrow to show it's sorted descending
                }
            }, 100);
        });
    </script>
</body>
</html>
"@
    
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
}

# Helper function for HTML encoding
function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    
    return $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;'
}

# Main execution
try {
    Write-Host "SharePoint Risk Analysis Tool" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan
    
    # Validate input file
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }
    
    # Set output path if not specified
    if (-not $OutputPath) {
        $OutputPath = [System.IO.Path]::ChangeExtension($CsvPath, 'html')
    }
    
    # Get scoring configuration
    $useCustomScoring = Get-ScoringPreference
    $scoringConfig = if ($useCustomScoring) { Get-CustomScoring } else { $DefaultScoring }
    
    Write-Host "`nLoading CSV data..." -ForegroundColor Yellow
    $sites = Import-Csv -Path $CsvPath
    Write-Host "Loaded $($sites.Count) sites" -ForegroundColor Green
    
    Write-Host "`nAnalyzing risk scores..." -ForegroundColor Yellow
    $analyzedSites = @()
    
    foreach ($site in $sites) {
        $riskResult = Calculate-RiskScore -Site $site -ScoringConfig $scoringConfig
        $analyzedSite = $site | Select-Object *, @{n='RiskScore';e={$riskResult.Score}}, @{n='RiskReasons';e={$riskResult.Reasons}}
        $analyzedSites += $analyzedSite
    }
    
    # Sort by risk score (highest first)
    $analyzedSites = $analyzedSites | Sort-Object RiskScore -Descending
    
    # Calculate statistics
    $statistics = @{
        TotalSites = $analyzedSites.Count
        HighRiskSites = ($analyzedSites | Where-Object { $_.RiskScore -ge 7 }).Count
        PublicSites = ($analyzedSites | Where-Object { $_.'Site privacy' -eq 'Public' }).Count
        SitesWithAnyoneLinks = ($analyzedSites | Where-Object { $_.'Anyone link count' -gt 0 }).Count
    }
    
    Write-Host "`nGenerating HTML report..." -ForegroundColor Yellow
    Generate-HtmlReport -AnalyzedSites $analyzedSites -ScoringConfig $scoringConfig -Statistics $statistics -OutputPath $OutputPath
    
    Write-Host "`n=== Analysis Complete ===" -ForegroundColor Green
    Write-Host "Report generated: $OutputPath" -ForegroundColor Green
    Write-Host "Total sites analyzed: $($statistics.TotalSites)" -ForegroundColor White
    Write-Host "High risk sites (score 7+): $($statistics.HighRiskSites)" -ForegroundColor Red
    Write-Host "Public sites: $($statistics.PublicSites)" -ForegroundColor Yellow
    Write-Host "Sites with anyone links: $($statistics.SitesWithAnyoneLinks)" -ForegroundColor Yellow
    
    # Show top 5 highest risk sites
    Write-Host "`nTop 5 Highest Risk Sites:" -ForegroundColor Cyan
    $topRisk = $analyzedSites | Select-Object -First 5
    foreach ($site in $topRisk) {
        $siteName = if ([string]::IsNullOrWhiteSpace($site.'Site name')) { "Unnamed Site" } else { $site.'Site name' }
        Write-Host "  $($site.RiskScore) - $siteName" -ForegroundColor White
    }
    
    # Open report in default browser
    Write-Host "`nOpening report in default browser..." -ForegroundColor Yellow
    Start-Process $OutputPath
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    exit 1
}