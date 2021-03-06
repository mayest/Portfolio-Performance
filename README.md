# Risk-Adjusted Portfolio Performance Measures
Excel Add-in<br>Created in C# with [Excel-DNA](https://github.com/Excel-DNA/ExcelDna)<br>Author: Timothy R. Mayes, Ph.D.<br>Version: 0.2<br>Date: 1 February 2020

# Purpose
This Excel add-in (an .xll file) contains functions that calculate common risk-adjusted performance measures. Required arguments typically include a series of asset/portfolio returns, market/benchmark portfolio (e.g., S&P 500) returns, and risk-free asset (e.g., U.S. Treasury security) returns.

## Functions Included:
Purpose|Function Name and Arguments
-------|---------------------------
Sharpe Ratio|SharpeRatio(Asset Returns, Risk-Free Returns, Data Frequency)
Revised Sharpe Ratio|RevisedSharpeRatio(Asset Returns, Risk-Free Returns, Data Frequency)
Adjusted Sharpe Ratio|AdjustedSharpeRatio(Asset Returns, Risk-Free Returns, Data Frequency)
M-Squared (i.e., the Modigliani & Modigliani measure)|MSquared(Asset Returns, Market Returns, Risk-Free Returns, Data Frequency)
Roy's "Safety First" Ratio|RoyRatio(Asset Returns, Risk-Free Returns, Data Frequency)
Information Ratio|InformationRatio(Asset Returns, Benchmark Returns, Data Frequency)
Treynor Index|TreynorIndex(Asset Returns, Risk-Free Returns, Asset Beta, Data Frequency)
Arithmetic Tracking Error|TrackingErrorArithmetic(Asset Returns, Benchmark Returns, Data Frequency)
Geometric Tracking Error|TrackingErrorGeometric(Asset Returns, Benchmark Returns, Data Frequency)
Beta|Beta(Asset Returns, Market Returns)
Adjusted Beta|AdjustedBeta(Asset Returns, Market Returns)
Bull Beta (beta in up markets)|BullBeta(Asset Returns, Market Returns)
Bear Beta (beta in down markets)|BearBeta(Asset Returns, Market Returns)
Beta Timing Ratio (ratio of bull beta to bear beta)|BetaTimingRatio(Asset Returns, Market Returns)
Jensen's Alpha|JensensAlpha(Asset Returns, Market Returns, Risk-Free Returns, Data Frequency)
Appraisal Ratio|AppraisalRatio(Asset Returns, Market Returns, Risk-Free Returns, Data Frequency)
Fama's Decomposition|FamaDecomposition(Asset Returns, Market Returns, Risk-Free Returns, Target Beta, Data Frequency)
Up Capture Ratio|UpCaptureRatio(Asset Returns, Benchmark Returns)
Down Capture Ratio|DownCaptureRatio(Asset Returns, Benchmark Returns)
Up Percentage Ratio|UpPercentageRatio(Asset Returns, Benchmark Returns)
Down Percentage Ratio|DownPercentageRatio(Asset Returns, Benchmark Returns)
Percentage Gain Ratio|PercentageGainRatio(Asset Returns, Benchmark Returns)
Percentage Loss Ratio|PercentageLossRatio(Asset Returns, Benchmark Returns)
Hurst Exponent|HurstExponent(Asset Returns)
Bias Ratio|BiasRatio(Asset Returns, Standard Deviations)
Market Risk|MarketRisk(Asset Returns, Benchmark Returns, Data Frequency)
Unique (Diversifiable)|Risk	UniqueRisk(Asset Returns, Benchmark Returns, Data Frequency)
Lower Partial Moment|LowerPartialMoment(Asset Returns, Target Return, Degree, Data Frequency)
Upper Partial Moment|UpperPartialMoment(Asset Returns, Target Return, Degree, Data Frequency)
Semi-Variance|SemiVariance(Asset Returns, Target Return, Data Frequency)
Semi-Deviation|SemiDeviation(Asset Returns, Target Return, Data Frequency)
Omega Ratio|OmegaRatio(Asset Returns, Target Return, Degree, Data Frequency)
Jarque-Bera Test|JarqueBeraTest(Asset Returns)
K Ratio|KRatio(Asset Returns)
Total Return Index|TotalReturnIndex(Asset Returns, Start Value)
Maximum Drawdown|MaxDrawDown(Asset Returns)
Maximum Drawdown by Year|MaxDrawDownByYear(Asset Returns, Dates)
Average Annual Max Drawdown|AverageMaxDrawDown(Asset Returns, Dates)
Average Drawdown|AverageDrawDown(Asset Returns, Count)
Maximum Drawdown Duration|MaxDrawDownDuration(Asset Returns)
Calmar Ratio|CalmarRatio(Asset Returns, Risk-Free Returns, Data Frequency)
Sterling Ratio|SterlingRatio(Asset Returns, Risk-Free Returns, Count, Data Frequency)
Ulcer Index|UlcerIndex(Asset Returns)
Ulcer Performance Index|UlcerPerformanceIndex(Asset Returns, Risk-Free Returns, Data Frequency)
Parametric Value at Risk|ParametricVaR(Asset Returns, Confidence Level)
Modified Parametric VaR|ModifiedParametricVaR(Asset Returns, Confidence Level)
Historical Simulation VaR|HistoricalSimulationVaR(Asset Returns, Confidence Level)
Holding Period Return|HoldingPeriodReturn(Prices, Cash Flows)
HPR with Reinvestment|HPRWithReinvestment (Prices, Cash Flows)
Sub-Period Returns|SubPeriodReturns(Prices, Cash Flows)
Log Sub-Period Returns|LogSubPeriodReturns(Prices, Cash Flows)

# Installation
The add-in .xll file (most should use PortfolioPerformance.xll, but if you are using 64-bit Excel then use PortfolioPerformance64.xll) can be installed in Microsoft Excel on the Windows platform (not available for Mac) in the usual way. Go to File -> Options -> Add-ins and then click the Go button at the bottom of the dialog box. Click the Browse button and select the .xll file from the directory where you saved it. From this point on, the add-in will be loaded automatically every time that you start Excel.

# Removal
To remove the Add-in from Excel, simply repeat the installation instructions, but remove the check mark next to the add-in name. This will cause it to be unloaded.
To permanently remove the add-in, simply delete the .xll file from the directory in which it is stored. Note that the next time you return to the add-ins dialog box, Excel will inform you that it cannot find the file and ask if you would like it to be removed from the list.

# Usage
The functions are available in the Insert Function dialog box in the Portfolio Performance category. Each function and argument has a short description. 

Of course, the functions can be entered manually as well, and the add-in includes Intellisense to help if you type the function directly. For example, typing =SharpeRatio(A1:A15, B1:B15) will calculate the Sharpe Ratio, assuming that the asset returns are in A1:A15 and the risk-free asset returns are in B1:B15.

Documentation is available at: https://github.com/mayest/Portfolio-Performance/blob/master/PortfolioPerformance/Performance%20Analysis%20Add-in%20Documentation.pdf.

# Updates
This free and open-source project is hosted on GitHub, and you can always download the latest version at https://github.com/mayest/. To update it, make sure that Excel is not open and then download the appropriate version of the add-in to the same folder where you had originally saved it. The next time that you start Excel, the updated version of the add-in will be used.

# License

Copyright (c) 2019-2020 Timothy R. Mayes. All rights reserved.

Licensed under the MIT License.
