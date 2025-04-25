# Risk-Analyzer

Risk Analyzer is a data-driven decision-making tool designed to help Small and Medium-sized Businesses (SMBs) evaluate whether to use Original Equipment Manufacturer (OEM) services or develop in-house manufacturing capabilities. The tool analyzes multiple risk factors including cost, time, quality, scalability, and market risks to provide actionable recommendations.

Key Features: 
Comprehensive Risk Analysis - Evaluates 8 key risk factors with weighted scoring
Cost Comparison - Calculates total costs for in-house manufacturing vs OEM
Financial Risk Assessment - Incorporates beta calculations from market data
Automated Recommendations - Provides clear guidance based on analysis results
User-friendly Interface: Simple data input through Google Sheets sidebar

Technical Implementation:
- Built using Google Apps Script (JavaScript)
- Integrates with Google Sheets for data input and analysis
- Uses Google Finance API for market data
- Modular design with separate HTML forms for different input types

Files included:
├── RiskAnalyzer.gs.js                           # Main Apps Script code
├── CostSidebar.html                             # Cost input form
├── FinancialSidebar.html                        # Financial data input form
├── OEMCostSidebar.html                          # OEM cost input form

How to Use
1. Set up the script in Google Sheets by pasting the code into the Apps Script editor
2. Input your data using the custom menu options
3. Cost attributes for in-house manufacturing
4. OEM costs
5. Company financial data for risk assessment
6. Run the analysis to get recommendations

Key Functions
1. setupSheet(): Prepares the data input template
2. calculateRiskScores(): Computes weighted risk scores
3. calculateRisks(): Performs comprehensive risk analysis
4. fetchData(): Retrieves market data for beta calculation

Team Members:
- Kin Sheau Xuan
- Lim Shu Ye
- Lim Ying Ying
- Yong Xuan Lyn

