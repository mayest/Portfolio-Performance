using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;


namespace PortfolioPerformance
{
    /// <summary>
    /// Purpose: This class contains various asset/portfolio performance (mostly risk-adjusted) measures.
    /// Author: Timothy R. Mayes, Ph.D.
    /// Date: 16 January 2020
    /// </summary>
    public class Measures
    {

        [ExcelFunction(Name = "RoyRatio", Description = "Calculates Roy's Safety First ratio for a set of asset returns", Category = "Portfolio Performance")]
        public static object RoyRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Minimum Target Return", Description = "(Optional) Range of minimum target returns", AllowReference = false)] object[] minTargetReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    return SharpeRatio(assetReturns, minTargetReturns, frequency);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "SharpeRatio", Description = "Calculates the Sharpe Ratio for a set of asset returns", Category = "Portfolio Performance")]
        //Calculates the Sharpe Ratio, which is: (average asset return less the average risk-free return)/Std Dev of asset return
        public static object SharpeRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double[] rf = Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double assetSd = Helpers.StdDev_P(assetReturns) * Math.Sqrt(freq);

                    if (Math.Abs(assetSd) > 0.0d)
                    {
                        double assetAnnRet = Helpers.AnnualizedReturn(assetReturns, freq);
                        double riskfreeAnnRet = Helpers.AnnualizedReturn(rf, freq);

                        return (assetAnnRet - riskfreeAnnRet) / assetSd;
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }

                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "RevisedSharpeRatio", Description = "Calculates the Revised Sharpe Ratio for a set of asset returns", Category = "Portfolio Performance")]
        //Calculates the Revised Sharpe Ratio, which is: (average asset return less the average risk-free return)/Std Dev of (asset returns - risk free returns)
        public static object RevisedSharpeRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) 
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double[] rf = Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double diffSd = Helpers.StdDev_P(Helpers.ArrayDiff(assetReturns, rf)) * Math.Sqrt(freq);

                    if (Math.Abs(diffSd) > 0.0d)
                    {
                        double assetAnnRet = Helpers.AnnualizedReturn(assetReturns, freq);
                        double riskfreeAnnRet = Helpers.AnnualizedReturn(rf, freq);

                        return (assetAnnRet - riskfreeAnnRet) / diffSd;
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }

                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }


        [ExcelFunction(Name = "AdjustedSharpeRatio", Description = "Sharpe Ratio adjusted for non-normal return distributions (negative skewness and leptokurtosis)", Category = "Portfolio Performance")]
        public static object AdjustedSharpeRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double sRatio = (double) SharpeRatio(assetReturns, riskFreeReturns, freq);
                    double skew = Helpers.Skewness_P(assetReturns);
                    double kurt = Helpers.Kurtosis_P(assetReturns);

                    return sRatio + skew / 6 * Math.Pow(sRatio, 2) - (kurt - 3) / 24 * Math.Pow(sRatio, 3);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }
        
        [ExcelFunction(Name = "MSquared", Description = "Calculates the Modigliani & Modigliani (M-squared) risk-adjusted performance measure for a set of asset returns", Category = "Portfolio Performance")]
        //This calculates the Modigliani & Modigliani (M-squared) risk-adjusted performance measure.
        //This is the asset return levered up or down so that it has the same standard deviation as the market portfolio.
        //In other words, we add or subtract the risk-free asset until the resulting portfolio has the same risk as the market.
        public static object MSquared(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && mktReturns.Length != assetReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double rf = Helpers.AnnualizedReturn(Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length)), freq);
                    double mktStdDev = Helpers.StdDev_P(mktReturns) * Math.Sqrt(freq);
                    return (double)SharpeRatio(assetReturns, riskFreeReturns, frequency) * mktStdDev + rf;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }
        
        [ExcelFunction(Name = "InformationRatioArithmetic", Description = "Calculates the information ratio of an asset", Category = "Portfolio Performance")]
        public static object InformationRatioArithmetic(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns.Length != benchReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double assetAnnualReturn = Helpers.AnnualizedReturn(assetReturns, freq);
                    double benchAnnualReturn = Helpers.AnnualizedReturn(benchReturns, freq);
                    double trackingError = (double)TrackingErrorArithmetic(assetReturns, benchReturns, frequency);
                    if (Math.Abs(trackingError) > 0.0d)
                    {
                        return (assetAnnualReturn - benchAnnualReturn) / trackingError;
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "InformationRatioGeometric", Description = "Calculates the information ratio of an asset", Category = "Portfolio Performance")]
        public static object InformationRatioGeometric(
    [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
    [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns,
    [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns.Length != benchReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double assetAnnualReturn = Helpers.AnnualizedReturn(assetReturns, freq);
                    double benchAnnualReturn = Helpers.AnnualizedReturn(benchReturns, freq);
                    double trackingError = (double)TrackingErrorGeometric(assetReturns, benchReturns, frequency);
                    if (Math.Abs(trackingError) > 0.0d)
                    {
                        return ((1 + assetAnnualReturn) / (1 + benchAnnualReturn) - 1) / trackingError;
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }
        
        [ExcelFunction(Name = "TrackingErrorArithmetic", Description = "Calculates the tracking error of an asset vs its benchmark", Category = "Portfolio Performance")]
        public static object TrackingErrorArithmetic(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns.Length != benchReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double[] diffArray = Helpers.ArrayDiff(assetReturns, benchReturns);
                    return Helpers.StdDev_P(diffArray) * Math.Sqrt(freq);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "TrackingErrorGeometric", Description = "Calculates the tracking error of an asset vs its benchmark", Category = "Portfolio Performance")]
        public static object TrackingErrorGeometric(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns.Length != benchReturns.Length)
                //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
                //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double[] diffArray = Helpers.ArrayDiffGeom(assetReturns, benchReturns);
                    return Helpers.StdDev_P(diffArray) * Math.Sqrt(freq);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "TreynorIndex", Description = "Calculates the Treynor Index for a set of asset returns", Category = "Portfolio Performance")]
        //Calculates the Treynor Index, which is: (average asset return less the average risk-free return)/beta of the asset
        public static object TreynorIndex(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Asset Beta", Description = "The beta of the asset", AllowReference = false)] object assetBeta, 
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetBeta is ExcelMissing)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    if (assetBeta is ExcelMissing)
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                        double[] rf = Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                        double assetAnnualReturn = Helpers.AnnualizedReturn(assetReturns, freq);
                        double rfAnnualReturn = Helpers.AnnualizedReturn(rf, freq);
                        if (Math.Abs((double)assetBeta) > 0.0d)
                        {
                            return (assetAnnualReturn - rfAnnualReturn) / (double)assetBeta;
                        }
                        else
                        {
                            return ExcelError.ExcelErrorDiv0;
                        }

                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }
        
        [ExcelFunction(Name = "BetaTimingRatio", Description = "Calculates the ratio of the bull beta to the bear beta", Category = "Portfolio Performance")]
        //This is the ratio of the bull beta to the bear beta and provides a measure of timing ability.
        public static object BetaTimingRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && mktReturns.Length != assetReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    double bullB = (double)RiskMeasures.BullBeta(assetReturns, mktReturns);
                    double bearB = (double)RiskMeasures.BearBeta(assetReturns, mktReturns);
                    if (Math.Abs(bearB) > 0.0d)
                    {
                        return bullB / bearB;
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }
                }
                catch
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }


        [ExcelFunction(Name = "UpCaptureRatio", Description = "Returns the ratio of the average asset returns in up markets to the average benchmark return in up markets", Category = "Portfolio Performance")]
        public static object UpCaptureRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && benchReturns.Length != assetReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    List<double> benchUpReturns = new List<double>();
                    List<double> assetUpReturns = new List<double>();

                    for (int i = 0; i < benchReturns.Length; i++)
                    {
                        if (benchReturns[i]>0)
                        {
                            benchUpReturns.Add(benchReturns[i]);
                            assetUpReturns.Add(assetReturns[i]);
                        }

                    }

                    if (benchUpReturns.Count > 0)
                    {
                        return assetUpReturns.Average() / benchUpReturns.Average();
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "DownCaptureRatio", Description = "Returns the ratio of the average asset returns in down markets to the average benchmark return in down markets", Category = "Portfolio Performance")]
        public static object DownCaptureRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && benchReturns.Length != assetReturns.Length)
                //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
                //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    List<double> benchDownReturns = new List<double>();
                    List<double> assetDownReturns = new List<double>();

                    for (int i = 0; i < benchReturns.Length; i++)
                    {
                        if (benchReturns[i] < 0)
                        {
                            benchDownReturns.Add(benchReturns[i]);
                            assetDownReturns.Add(assetReturns[i]);
                        }

                    }

                    if (benchDownReturns.Count > 0)
                    {
                        return assetDownReturns.Average() / benchDownReturns.Average();
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }
                    
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "UpPercentageRatio", Description = "Returns the percentage of the time that the portfolio outperforms the benchmark return in up markets", Category = "Portfolio Performance")]
        public static object UpPercentageRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && benchReturns.Length != assetReturns.Length)
                //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
                //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    List<double> benchUpReturns = new List<double>();
                    int outPerformCount = 0;

                    //Get returns in positive markets
                    for (int i = 0; i < benchReturns.Length; i++)
                    {
                        if (benchReturns[i] > 0)
                        {
                            benchUpReturns.Add(benchReturns[i]);
                            if (assetReturns[i] > benchReturns[i]) outPerformCount++;
                        }

                    }
                    if (benchUpReturns.Count > 0)
                    {
                        return (double)outPerformCount / (double)benchUpReturns.Count;
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "DownPercentageRatio", Description = "Returns the percentage of the time that the portfolio outperforms the benchmark return in down markets", Category = "Portfolio Performance")]
        public static object DownPercentageRatio(
    [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
    [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && benchReturns.Length != assetReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    List<double> benchDownReturns = new List<double>();
                    int outPerformCount = 0;

                    //Get returns in negative markets
                    for (int i = 0; i < benchReturns.Length; i++)
                    {
                        if (benchReturns[i] < 0)
                        {
                            benchDownReturns.Add(benchReturns[i]);
                            if (assetReturns[i] > benchReturns[i]) outPerformCount++;
                        }

                    }

                    if (benchDownReturns.Count > 0)
                    {
                        return (double)outPerformCount / (double)benchDownReturns.Count;
                    }
                    else
                    {
                        return ExcelError.ExcelErrorDiv0;
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }


        [ExcelFunction(Name = "PercentageGainRatio", Description = "Compares the number of positive asset returns to the number of positive benchmark returns.", Category = "Portfolio Performance")]
        public static object PercentageGainRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && benchReturns.Length != assetReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double benchPos = 0;
                    double assetPos = 0;
                    for (int i = 0; i < assetReturns.Length; i++)
                    {
                        if (benchReturns[i] > 0) benchPos++;
                        if (assetReturns[i] > 0) assetPos++;
                    }

                    return assetPos / benchPos;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "PercentageLossRatio", Description = "Compares the number of negative asset returns to the number of negative benchmark returns.", Category = "Portfolio Performance")]
        public static object PercentageLossRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && benchReturns.Length != assetReturns.Length)
                //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
                //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double benchNeg = 0;
                    double assetNeg = 0;
                    for (int i = 0; i < assetReturns.Length; i++)
                    {
                        if (benchReturns[i] < 0) benchNeg++;
                        if (assetReturns[i] < 0) assetNeg++;
                    }

                    return assetNeg / benchNeg;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "HurstExponent", Description = "Calculates the Hurst Exponent of a return series.", Category = "Portfolio Performance")]
        public static object HurstExponent(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double assetMean = assetReturns.Average();
                    double assetSD = Helpers.StdDev_P(assetReturns);
                    int n = assetReturns.Length;
                    List<double> cumulativeDeviations = new List<double>();
                    double prevDev = 0; //previous return
                    for (int i = 0; i < assetReturns.Length; i++)
                    {
                        cumulativeDeviations.Add(prevDev+(assetReturns[i] - assetMean));
                        prevDev += assetReturns[i] - assetMean;
                    }
                    double maxCD = cumulativeDeviations.Max();
                    double minCD = cumulativeDeviations.Min();
                    return Math.Log((maxCD - minCD) / assetSD, Math.E) / Math.Log(n, Math.E);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }
        
        [ExcelFunction(Name = "BiasRatio", Description = "Calculates the ratio of returns between 0 and +1 standard deviation to those between -1 standard deviation and 0.", Category = "Portfolio Performance")]
        public static object BiasRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Standard Deviations", Description = "(Optional) Number of Standard Deviations Away from 0%, default is 1", AllowReference = false)] object stdDevs)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    int countAbove = 0;
                    int countBelow = 0;

                    double numStdDevs = (stdDevs is ExcelMissing) ? 1d : (double)stdDevs;
                    double assetSD = Helpers.StdDev_P(assetReturns);
                    double topRange = numStdDevs * assetSD;
                    double bottomRange = -numStdDevs * assetSD;
                    foreach (var ret in assetReturns)
                    {
                        if (ret >= 0 && ret <= topRange) countAbove++;
                        if (ret < 0 && ret >= bottomRange) countBelow++;
                    }
                    
                    return (double)countAbove / (1 + (double)countBelow);

                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }
        
        [ExcelFunction(Name = "JensensAlpha", Description = "Calculates Jensen's alpha for an asset", Category = "Portfolio Performance")]
        public static object JensensAlpha(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && (mktReturns.Length != assetReturns.Length || mktReturns.Length != riskFreeReturns.Length || assetReturns.Length != riskFreeReturns.Length))
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double[] rf = Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double assetAnnualReturn = Helpers.AnnualizedReturn(assetReturns, freq);
                    double rfAnnualReturn = Helpers.AnnualizedReturn(rf, freq);
                    double mktAnnualReturn = Helpers.AnnualizedReturn(mktReturns, freq); 
                    double assetBeta = (double)RiskMeasures.Beta(assetReturns, mktReturns); 
                    return (assetAnnualReturn - rfAnnualReturn) - assetBeta * (mktAnnualReturn - rfAnnualReturn);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }
        
        [ExcelFunction(Name = "AppraisalRatio", Description = "Calculates the Treynor & Black appraisal ratio for a set of asset/portfolio returns", Category = "Portfolio Performance")]
        public static object AppraisalRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns.Length != mktReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double jensenAlpha = (double)JensensAlpha(assetReturns, mktReturns, riskFreeReturns, frequency);
                    double uniqueRisk = (double) RiskMeasures.UniqueRisk(assetReturns, mktReturns, frequency);
                    return jensenAlpha/Math.Sqrt(uniqueRisk);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }
        
        [ExcelFunction(Name = "FamaDecomposition", Description = "Returns an array with Fama's decomposition of the excess return", Category = "Portfolio Performance")]
        public static object FamaDecomposition(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns, 
            [ExcelArgument(Name = "Target Beta", Description = "The target beta for the asset", AllowReference = false)] object targetBeta,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns.Length != mktReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    int nOutputs = (targetBeta is ExcelMissing) ? 8 : 10; //Set number of output cells
                    object[,] outputArray = new object[nOutputs, 2]; //Create an array to hold the outputs
                    double[] rf = Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double assetAnnualReturn = Helpers.AnnualizedReturn(assetReturns, freq);
                    double assetStdDev = Helpers.StdDev_P(assetReturns) * Math.Sqrt(freq);
                    double mktAnnualReturn = Helpers.AnnualizedReturn(mktReturns, freq);
                    double mktStdDev = Helpers.StdDev_P(mktReturns) * Math.Sqrt(freq);
                    double rfAnnualReturn = Helpers.AnnualizedReturn(rf, freq);
                    double beta = (double)RiskMeasures.Beta(assetReturns, mktReturns);
                    double hypBeta = assetStdDev / mktStdDev;//Hypothetical beta (i.e., beta if portfolio was perfectly diversified and therefore has perfect correlation with market)
                    double hypReturn = rfAnnualReturn + hypBeta * (mktAnnualReturn - rfAnnualReturn);//Expected return based on hypothetical beta
                    double hypRiskPremium = hypReturn - rfAnnualReturn;

                    double totalRiskPremium = assetAnnualReturn - rfAnnualReturn;
                    double rpDueToRisk = beta * (mktAnnualReturn - rfAnnualReturn);
                    double rpDueToSelectivity = assetAnnualReturn - rfAnnualReturn - rpDueToRisk;
                    double diversification = (mktStdDev / assetStdDev - beta) * (mktAnnualReturn - rfAnnualReturn);
                    double netSelectivity = rpDueToSelectivity - diversification;

                    if (nOutputs == 8)
                    {
                        outputArray[0, 0] = "Risk Premium"; outputArray[0, 1] = totalRiskPremium;
                        outputArray[1, 0] = "Due to Risk"; outputArray[1, 1] = rpDueToRisk;
                        outputArray[2, 0] = "Due to Selectivity"; outputArray[2, 1] = rpDueToSelectivity;
                        outputArray[3, 0] = "Diversification"; outputArray[3, 1] = diversification;
                        outputArray[4, 0] = "Net Selectivity"; outputArray[4, 1] = netSelectivity;
                        outputArray[5, 0] = "Hypothetical Beta"; outputArray[5, 1] = hypBeta;
                        outputArray[6, 0] = "Hypothetical Exp Return"; outputArray[6, 1] = hypReturn;
                        outputArray[7, 0] = "Hypothetical Risk Premium"; outputArray[7, 1] = hypRiskPremium;
                        return outputArray;
                    }
                    else
                    {
                        // Here we have a target beta, so we can decompose risk premium due to risk
                        // Populating outputArray separately because we are reordering the output
                        double invRisk = (double)targetBeta * (mktAnnualReturn - rfAnnualReturn);
                        double mgrRisk = (beta - (double)targetBeta) * (mktAnnualReturn - rfAnnualReturn);

                        outputArray[0, 0] = "Risk Premium"; outputArray[0, 1] = totalRiskPremium;
                        outputArray[1, 0] = "Due to Risk"; outputArray[1, 1] = rpDueToRisk;
                        outputArray[2, 0] = "Due to Investor's Risk"; outputArray[2, 1] = invRisk;
                        outputArray[3, 0] = "Due to Manager's Risk"; outputArray[3, 1] = mgrRisk;
                        outputArray[4, 0] = "Due to Selectivity"; outputArray[4, 1] = rpDueToSelectivity;
                        outputArray[5, 0] = "Diversification"; outputArray[5, 1] = diversification;
                        outputArray[6, 0] = "Net Selectivity"; outputArray[6, 1] = netSelectivity;
                        outputArray[7, 0] = "Hypothetical Beta"; outputArray[7, 1] = hypBeta;
                        outputArray[8, 0] = "Hypothetical Exp Return"; outputArray[8, 1] = hypReturn;
                        outputArray[9, 0] = "Hypothetical Risk Premium"; outputArray[9, 1] = hypRiskPremium;
                        
                        return outputArray;
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "KRatio", Description = "Calculates Kestner's K Ratio for a series of returns. Higher values suggest more consistency of return.", Category = "Portfolio Performance")]
        public static object KRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    Int32 n = assetReturns.Length + 1;
                    double meanPeriod = 0.5 * (n + 1); //Average period number
                    double periodCountVar = 0;
                    double meanCumRet = 0;
                    double[] cumRet = new double[n]; 
                    double cumCov = 0;
                    for (int i = 0; i < n; i++) //Build cumulative return series
                    {
                        cumRet[i] = (i == 0) ? 0 : (1 + cumRet[i - 1]) * (1 + assetReturns[i - 1]) - 1; //assetReturns[i]
                    }

                    meanCumRet = cumRet.Average();
                    for (int i = 0; i < n; i++)
                    {
                        cumCov += ((cumRet[i] - meanCumRet) * (i + 1 - meanPeriod)) / n;
                        periodCountVar += Math.Pow(i + 1 - meanPeriod, 2) / n;
                    }
                    double kBeta = cumCov / periodCountVar;
                    double kIntercept = meanCumRet - kBeta * meanPeriod;
                    double[] errSquared = new double[n];

                    for (int i = 0; i < n; i++)
                    {
                        errSquared[i] = Math.Pow(cumRet[i] - (kIntercept + kBeta * (i + 1)), 2);
                    }

                    double meanSquareError = errSquared.Sum()/(n - 2);
                    double stdErrorBeta = Math.Pow(meanSquareError / (periodCountVar * n), 0.5);
                    return kBeta / stdErrorBeta / Math.Pow(n, 0.5);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }


        [ExcelFunction(Name = "TotalReturnIndex", Description = "Creates a total return index from a series of returns and a starting value (returns an array of values)", Category = "Portfolio Performance")]
        public static object TotalReturnIndex(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Start Value", Description = "(Optional) Initial investment in the asset/portfolio (default is 1)", AllowReference = false)] object startValue)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double sValue = (startValue is ExcelMissing) ? 1d : (double)startValue; //Set the starting value
                    return Helpers.ConvertToColumnArray(Helpers.GetTotalReturnIndex(assetReturns, sValue));
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }


        [ExcelFunction(Name = "MaxDrawDown", Description = "Calculates the maximum drawdown from a set of returns", Category = "Portfolio Performance")]
        public static object MaxDrawDown(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    return Helpers.GetDrawDowns(assetReturns).Min();
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "AverageDrawDown", Description = "Returns the average of the largest drawdowns", Category = "Portfolio Performance")]
        public static object AverageDrawDown(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Count", Description = "(Optional) Number of drawdowns to include in the average. Default is all of them.", AllowReference = false)] object nDrawDowns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double[] continuousDrawDowns = Helpers.GetContinuousDrawDowns(assetReturns);
                    Int32 n = (nDrawDowns is ExcelMissing) ? continuousDrawDowns.Length : (Int32)Math.Truncate((double)nDrawDowns); //Set the number of drawdowns to use in the average
                    var nLargest = continuousDrawDowns.Where(x => (double)x < 0).OrderBy(x => (double)x).Take(n);
                    return nLargest.Average(); //Return average of the n smallest numbers
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "MaxDrawDownDuration", Description = "Calculates the longest amount of time between peaks in asset performance", Category = "Portfolio Performance")]
        public static object MaxDrawDownDuration(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
                //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double[] totReturnIndex = Helpers.GetTotalReturnIndex(assetReturns, 1);
                    double[] peaks = Helpers.GetPeaks(assetReturns);
                    List<int> peakPeriods = new List<int>();
                    List<int> peakDurations = new List<int>();
                    peakPeriods.Add(0);
                    for (int i = 1; i < totReturnIndex.Length - 1; i++)
                    {
                        if (peaks[i] > peaks[i - 1])
                        {
                            peakPeriods.Add(i + 1);//Get the period number of each peak
                        }
                    }
                    if (peakPeriods.Count > 1)
                    {
                        double prevMaxPeak = 1d;
                        int prevMaxPeakPeriod = 0;
                        for (int i = 1; i < peakPeriods.Count; i++)
                        {
                            if (totReturnIndex[peakPeriods[i]] > prevMaxPeak)
                            {
                                prevMaxPeak = totReturnIndex[peakPeriods[i]];//Set the prevMaxPeak to this value
                                peakDurations.Add(peakPeriods[i] - peakPeriods[prevMaxPeakPeriod]);//Calculate periods since last peak
                                prevMaxPeakPeriod = i;//Set the last peak period to this one
                            }
                        }

                        return peakDurations.Max();
                    }
                    else
                    {
                        return 0d;
                    }

                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "UlcerIndex", Description = "Calculates Peter G. Martin's Ulcer Index from a set of asset/portfolio returns", Category = "Portfolio Performance")]
        public static object UlcerIndex(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double[] drawDowns = Helpers.GetDrawDowns(assetReturns);
                    double sumSqDd = 0;
                    //Calculate sum of squared drawdowns.
                    //Note that this varies slightly from Bacon's result.
                    //Bacon ignores the initial 0 drawdown, while Martin does not. I match Martin's example exactly.
                    //See http://www.tangotools.com/ui/UlcerIndex.xls

                    for (int i = 1; i < drawDowns.Length; i++)
                    {
                        sumSqDd += drawDowns[i] * drawDowns[i];
                    }

                    return Math.Pow(sumSqDd / drawDowns.Length, 0.5);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }


        [ExcelFunction(Name = "UlcerPerformanceIndex", Description = "Similar to the Sharpe ratio, except that it uses the Ulcer Index as the risk measure.", Category = "")]
        public static object UlcerPerformanceIndex(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double[] rf = Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double assetAnnualReturn = Helpers.AnnualizedReturn(assetReturns, freq);
                    double rfAnnualReturn = Helpers.AnnualizedReturn(rf, freq);
                    return (assetAnnualReturn - rfAnnualReturn) / (double)UlcerIndex(assetReturns);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }


        [ExcelFunction(Name = "CalmarRatio", Description = "Calculates the Calmar Ratio for a series of asset returns", Category = "Portfolio Performance")]
        public static object CalmarRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns,
            [ExcelArgument(Name = "Data Frequency", Description = "(Optional) Number of periods per year (annual = 1, monthly = 12, etc)", AllowReference = false)] object frequency)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double freq = (frequency is ExcelMissing) ? 1d : (double)frequency; //Set the frequency
                    double[] rf = Helpers.ObjToDouble(Helpers.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double assetAnnualReturn = Helpers.AnnualizedReturn(assetReturns, freq);
                    double rfAnnualReturn = Helpers.AnnualizedReturn(rf, freq);
                    double maxDd = (double) MaxDrawDown(assetReturns);

                    return (assetAnnualReturn - rfAnnualReturn) / maxDd;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }


    }//End of Class Measures
}//End of Namespace PortfolioPerformance


