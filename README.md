# Risk-Adjusted Portfolio Performance Measures
Excel Add-in<br>Created in C# with [Excel-DNA](https://github.com/Excel-DNA/ExcelDna)<br>Author: Timothy R. Mayes, Ph.D.<br>Version: 0.1<br>Date: 25 December 2019

# Purpose
This Excel add-in (an .xll file) contains functions that calculate common risk-adjusted performance measures. Required arguments typically include a series of asset/portfolio returns, market/benchmark portfolio (e.g., S&P 500) returns, and risk-free asset (e.g., U.S. Treasury security) returns.

## Functions Included:
Purpose|Function Name and Arguments
-------|---------------------------
Sharpe Ratio|SharpeRatio(Asset Returns, Risk-Free Returns)
Revised Sharpe Ratio|RevisedSharpeRatio(Asset Returns, Risk-Free Returns)
M-Squared (i.e., the Modigliani & Modigliani measure)|MSquared(Asset Returns, Market Returns, Risk-Free Returns)
Information Ratio|InformationRatio(Asset Returns, Benchmark Returns)
Treynor Index|TreynorIndex(Asset Returns, Risk-Free Returns, Asset Beta)
Tracking Error|TrackingError(Asset Returns, Benchmark Returns)
Beta|Beta(Asset Returns, Market Returns)
Bull Beta (beta in up markets)|BullBeta(Asset Returns, Market Returns)
Bear Beta (beta in down markets)|BearBeta(Asset Returns, Market Returns)
Beta Timing Ratio (ratio of bull beta to bear beta)|BetaTimingRatio(Asset Returns, Market Returns)
Jensen's Alpha|JensensAlpha(Asset Returns, Market Returns, Risk-Free Returns)
Fama's Decomposition|FamaDecomposition(Asset Returns, Market Returns, Risk-Free Returns, Target Beta)

# Installation
The add-in .xll file can be installed in Microsoft Excel on the Windows platform (not available for Mac) in the usual way. Go to File -> Options -> Add-ins and then click the Go button at the bottom of the dialog box. Click the Browse button and select the .xll file from the directory where you saved it. From this point on, the add-in will be loaded automatically every time that you start Excel.

# Removal
To remove the Add-in from Excel, simply repeat the installation instructions, but remove the check mark next to the add-in name. This will cause it to be unloaded.
To permanently remove the add-in, simply delete the .xll file from the directory in which it is stored. Note that the next time you return to the add-ins dialog box, Excel will inform you that it cannot find the file and ask if you would like it to be removed from the list.

# Usage
The functions are available in the Insert Function dialog box in the Portfolio Performance category. Each function and argument has a short description.

Of course, the functions can be entered manually as well. For example, typing =SharpeRatio(A1:A15, B1:B15) will calculate the Sharpe Ratio, assuming that the asset returns are in A1:A15 and the risk-free asset returns are in B1:B15.


# License

Copyright (c) Timothy R. Mayes. All rights reserved.

Licensed under the MIT License.
