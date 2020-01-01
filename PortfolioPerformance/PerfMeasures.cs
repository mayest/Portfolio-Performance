using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;


namespace PortfolioPerformance
{
    public class Measures
    {
        [ExcelFunction(Name = "SharpeRatio", Description = "Calculates the Sharpe Ratio for a set of asset returns", Category = "Portfolio Performance")]
        //Calculates the Sharpe Ratio, which is: (average asset return less the average risk-free return)/Std Dev of asset return
        public static object SharpeRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] object[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns[0] is ExcelMissing)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    double[] rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double[] assetRets = Statistics.ObjToDouble(assetReturns);
                    double assetMean = assetRets.Average();
                    double assetSd = Statistics.StdDev_S(assetRets);
                    double rfMean = rf.Average();
                    if (Math.Abs(assetSd) > 0.0d)
                    {
                        return (assetMean - rfMean) / assetSd;
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
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] object[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns[0] is ExcelMissing)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    double[] rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double[] assetRets = Statistics.ObjToDouble(assetReturns);
                    double assetMean = assetRets.Average();
                    double diffSd = Statistics.StdDev_S(Statistics.ArrayDiff(assetRets, rf));
                    double rfMean = rf.Average();
                    if (Math.Abs(diffSd) > 0.0d)
                    {
                        return (assetMean - rfMean) / diffSd;
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

        [ExcelFunction(Name = "MSquared", Description = "Calculates the Modigliani & Modigliani (M-squared) risk-adjusted performance measure for a set of asset returns", Category = "Portfolio Performance")]
        //This calculates the Modigliani & Modigliani (M-squared) risk-adjusted performance measure.
        //This is the asset return levered up or down so that it has the same standard deviation as the market portfolio.
        //In other words, we add or subtract the risk-free asset until the resulting portfolio has the same risk as the market.
        public static object MSquared(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns)
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
                    double assetMean = assetReturns.Average();
                    double[] rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double rfMean = rf.Average();
                    double mktStdDev = Statistics.StdDev_S(mktReturns);
                    double assetStdDev = Statistics.StdDev_S(assetReturns);
                    if (Math.Abs(assetStdDev) > 0.0d)
                    {
                        return (mktStdDev / assetStdDev) * (assetMean - rfMean) + rfMean;
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


        [ExcelFunction(Name = "InformationRatio", Description = "Calculates the information ratio of an asset", Category = "Portfolio Performance")]
        public static object InformationRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
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
                    double[] diffArray = Statistics.ArrayDiff(assetReturns, benchReturns);
                    double trackingError = Statistics.StdDev_S(diffArray);
                    if (Math.Abs(trackingError) > 0.0d)
                    {
                        return diffArray.Average() / trackingError;
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

        [ExcelFunction(Name = "TrackingError", Description = "Calculates the tracking error of an asset vs its benchmark", Category = "Portfolio Performance")]
        public static object TrackingError(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)] double[] benchReturns)
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
                    double[] diffArray = Statistics.ArrayDiff(assetReturns, benchReturns);
                    return Statistics.StdDev_S(diffArray);
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
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns)
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
                        double[] rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                        double assetMean = assetReturns.Average();
                        double rfMean = rf.Average();
                        if (Math.Abs((double)assetBeta) > 0.0d)
                        {
                            return (assetMean - rfMean) / (double)assetBeta;
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

        [ExcelFunction(Name = "Beta", Description = "Calculates the Beta (systematic risk) of an Asset", Category = "Portfolio Performance")]
        public static object Beta(
                [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
                [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns)
        //Calculate the beta as cov(assetReturns-riskFreeReturns, mktReturns-riskFreeReturns)/var(mktReturns-riskFreeReturns)
        //The user can either supply just the two sets of returns, or all three, or a constant for the risk-FreeReturns.
        //If riskFreeReturns is a constant or omitted, then it will be extended.
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
                    double cov = Statistics.Covariance_S(assetReturns, mktReturns);
                    double mktVar = Statistics.Variance_S(mktReturns);
                    if (mktVar > 0)
                    {
                        return cov / mktVar;
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


        [ExcelFunction(Name = "AdjustedBeta", Description = "Calculates beta using Blume's Adjustment for the tendency to revert towards 1.00", Category = "Portfolio Performance")]
        public static object AdjustedBeta(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && mktReturns.Length != assetReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    return (double)Beta(assetReturns, mktReturns) * 2.0d / 3.0d + 1.0d / 3.0d;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "BullBeta", Description = "Calculates the Bull Beta of an Asset (uses only returns when the market is up)", Category = "Portfolio Performance")]
        public static object BullBeta(
                [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
                [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns)
        //This is the same as beta, except that it only looks at those periods where the market portfolio has a positive return.
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
                    List<double> mktUpReturns = new List<double>();
                    List<double> assetUpReturns = new List<double>();

                    for (int i = 0; i < mktReturns.Length; i++)
                    {
                        if (mktReturns[i] > 0)
                        {
                            mktUpReturns.Add(mktReturns[i]);
                            assetUpReturns.Add(assetReturns[i]);
                        }
                    }

                    if (mktUpReturns.Count == 0) //No positive returns
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        //Get beta using returns only when market is positive
                        return Beta(assetUpReturns.ToArray(), mktUpReturns.ToArray());
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "BearBeta", Description = "Calculates the Bear Beta of an Asset (uses only returns when the market is down)", Category = "Portfolio Performance")]
        public static object BearBeta(
                [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
                [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns)
        //This is the same as beta, except that it only looks at those periods where the market portfolio has a negative return.
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
                    List<double> mktDownReturns = new List<double>();
                    List<double> assetDownReturns = new List<double>();

                    for (int i = 0; i < mktReturns.Length; i++)
                    {
                        if (mktReturns[i] < 0)
                        {
                            mktDownReturns.Add(mktReturns[i]);
                            assetDownReturns.Add(assetReturns[i]);
                        }
                    }

                    if (mktDownReturns.Count == 0) //No negative returns
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        //Get beta using returns only when market is negative
                        return Beta(assetDownReturns.ToArray(), mktDownReturns.ToArray());
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
                    double bullB = (double)BullBeta(assetReturns, mktReturns);
                    double bearB = (double)BearBeta(assetReturns, mktReturns);
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
                    return assetUpReturns.Average()/benchUpReturns.Average();

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
                    return assetDownReturns.Average() / benchDownReturns.Average();

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

                    return (double)outPerformCount / (double)benchUpReturns.Count;

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

                    return (double)outPerformCount / (double)benchDownReturns.Count;

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
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "(Optional) Range of risk-free asset returns", AllowReference = false)] object[] riskFreeReturns)
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
                    double[] rf = new double[assetReturns.Length];
                    rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double assetMean = assetReturns.Average();
                    double rfMean = rf.Average();
                    double mktMean = mktReturns.Average();
                    double assetBeta = (double)Beta(assetReturns, mktReturns); //Calculate beta without subtracting the risk-free rate
                    return (assetMean - rfMean) - assetBeta * (mktMean - rfMean);
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
            [ExcelArgument(Name = "Target Beta", Description = "The target beta for the asset", AllowReference = false)] object targetBeta)
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
                    short nOutputs = new short();
                    if (targetBeta is ExcelMissing)
                    {
                        nOutputs = 8;
                    }
                    else
                    {
                        nOutputs = 10;
                    }
                    object[,] outputArray = new object[nOutputs, 2];
                    double[] rf = new double[assetReturns.Length];
                    rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    double assetMean = assetReturns.Average();
                    double assetSD = Statistics.StdDev_S(assetReturns);
                    double mktMean = mktReturns.Average();
                    double mktSD = Statistics.StdDev_S(mktReturns);
                    double rfMean = rf.Average();
                    double beta = (double)Beta(assetReturns, mktReturns);
                    double hypBeta = assetSD / mktSD;//Hypothetical beta (i.e., beta if portfolio was perfectly diversified and therefore has perfect correlation with market)
                    double hypReturn = rfMean + hypBeta * (mktMean - rfMean);//Expected return based on hypothetical beta
                    double hypRiskPremium = hypReturn - rfMean;

                    double totalRiskPremium = assetMean - rfMean;
                    double rpDueToRisk = beta * (mktMean - rfMean);
                    double rpDueToSelectivity = assetMean - rfMean - rpDueToRisk;
                    double diversification = (mktSD / assetSD - beta) * (mktMean - rfMean);
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
                        double invRisk = (double)targetBeta * (mktMean - rfMean);
                        double mgrRisk = (beta - (double)targetBeta) * (mktMean - rfMean);

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

        [ExcelFunction(Name = "JarqueBeraTest", Description = "Returns a test statistic that tests for normality", Category = "Portfolio Performance")]
        public static object JarqueBeraTest([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns)
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
                    double n = assetReturns.Length;
                    if (n > 2)
                    {
                        double skew = Statistics.Skewness_P(assetReturns);
                        double kurtExcess = Statistics.Kurtosis_P_Excess(assetReturns);
                        return n / 6 * (Math.Pow(skew, 2) + Math.Pow(kurtExcess, 2) / 4);
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



    }//End of Class Measures
}//End of Namespace PortfolioPerformance


