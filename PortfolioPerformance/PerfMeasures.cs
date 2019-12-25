using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;
using ExcelDna.IntelliSense;

namespace PortfolioPerformance
{
    public class Measures
    {

        [ExcelFunction(Name = "SharpeRatio", Description = "Calculates the Sharpe Ratio for a set of asset returns",
            Category = "Portfolio Performance")]
        //Calculates the Sharpe Ratio, which is: (average asset return less the average risk-free return)/Std Dev of asset return
        public static object SharpeRatio([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "Range of Risk-free Returns", AllowReference = false)]
            double[] riskFreeReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && assetReturns.Length != riskFreeReturns.Length)
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
                    double assetSd = Statistics.StdDev_S(assetReturns);
                    if (riskFreeReturns is null) //If no risk-free rates, then set rfMean to 0
                    {
                        double rfMean = 0;
                        return (assetMean - rfMean) / assetSd;
                    }
                    else
                    {
                        double rfMean = riskFreeReturns.Average();
                        return (assetMean - rfMean) / assetSd;
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "MSquared", Description = "Calculates the Modigliani & Modigliani (M-squared) risk-adjusted performance measure\nfor a set of asset returns",
            Category = "Portfolio Performance")]
        //This calculates the Modigliani & Modigliani (M-squared) risk-adjusted performance measure.
        //This is the asset return levered up or down so that it has the same standard deviation as the market portfolio.
        //In other words, we add or subtract the risk-free asset until the resulting portfolio has the same risk as the market.
        public static object MSquared([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)]
            double[] mktReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "Range of Risk-free Returns", AllowReference = false)]
            double[] riskFreeReturns)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && (mktReturns.Length != assetReturns.Length || mktReturns.Length != riskFreeReturns.Length || assetReturns.Length != riskFreeReturns.Length))
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
                    double rfMean = riskFreeReturns.Average();
                    double mktStdDev = Statistics.StdDev_S(mktReturns);
                    double assetStdDev = Statistics.StdDev_S(assetReturns);
                    return (mktStdDev / assetStdDev) * (assetMean - rfMean) + rfMean;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }


        [ExcelFunction(Name = "InformationRatio", Description = "Calculates the information ratio of an asset", Category = "Portfolio Performance")]
        public static object InformationRatio([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
        [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)]
            double[] benchReturns)
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
                    double assetMean = assetReturns.Average();
                    double benchMean = benchReturns.Average();
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
        public static object TrackingError([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
            [ExcelArgument(Name = "Benchmark Returns", Description = "Range of Benchmark Returns", AllowReference = false)]
            double[] benchReturns)
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


        [ExcelFunction(Name = "TreynorIndex", Description = "Calculates the Treynor Index for a set of asset returns",
            Category = "Portfolio Performance")]
        //Calculates the Treynor Index, which is: (average asset return less the average risk-free return)/beta of the asset
        public static object TreynorIndex([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
            [ExcelArgument(Name = "Asset Beta", Description = "The beta of the asset", AllowReference = false)] object assetBeta, [ExcelArgument(Name = "Risk-free Asset Returns", Description = "Range of Risk-free Returns", AllowReference = false)]
            object[] riskFreeReturns)
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
        public static object Beta([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
    [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)]
            double[] mktReturns,
    [ExcelArgument(Name = "Risk-free Returns", Description = "Range of Risk-free asset Returns", AllowReference = false)]
            object[] riskFreeReturns)
        //Calculate the beta as cov(assetReturns-riskFreeReturns, mktReturns-riskFreeReturns)/var(mktReturns-riskFreeReturns)
        //The user can either supply just the two sets of returns, or all three, or a constant for the risk-FreeReturns.
        //If riskFreeReturns is a constant or omitted, then it will be extended.
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && mktReturns.Length < assetReturns.Length)
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
                    double cov = Statistics.Covariance_S(Statistics.ArrayDiff(assetReturns, rf), Statistics.ArrayDiff(mktReturns, rf));
                    double mktVar = Statistics.Variance_S(Statistics.ArrayDiff(mktReturns, rf));
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


        [ExcelFunction(Name = "BullBeta", Description = "Calculates the Bull Beta of an Asset (uses only returns when the market is up)", Category = "Portfolio Performance")]
        public static object BullBeta([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
                double[] assetReturns,
                [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)]
                double[] mktReturns, [ExcelArgument(Name = "Risk-free Returns", Description = "Range of Risk-free asset Returns", AllowReference = false)]
                object[] riskFreeReturns)
        //This is the same as beta, except that it only looks at those periods where the market portfolio has a positive return.
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && mktReturns.Length < assetReturns.Length)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return 0.00d; //Return a placeholder value until both ranges are fully entered
            }
            else
            {
                try
                {
                    List<double> mktUpReturns = new List<double>();
                    List<double> assetUpReturns = new List<double>();
                    List<double> rfUpReturns = new List<double>();
                    double[] rf = new double[assetReturns.Length];

                    rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));

                    for (int i = 0; i < mktReturns.Length; i++)
                    {
                        if (mktReturns[i] > 0)
                        {
                            mktUpReturns.Add(mktReturns[i]);
                            assetUpReturns.Add(assetReturns[i]);
                            rfUpReturns.Add(rf[i]);
                        }
                    }

                    if (mktUpReturns.Count == 0)
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        //Get beta using returns only when market is positive
                        return Beta(assetUpReturns.ToArray(), mktUpReturns.ToArray(),
                            Statistics.DoubleToObject(rfUpReturns.ToArray()));
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "BearBeta", Description = "Calculates the Bear Beta of an Asset (uses only returns when the market is down)", Category = "Portfolio Performance")]
        public static object BearBeta([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
                double[] assetReturns,
        [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)]
                double[] mktReturns, [ExcelArgument(Name = "Risk-free Returns", Description = "Range of Risk-free asset Returns", AllowReference = false)]
                object[] riskFreeReturns)
        //This is the same as beta, except that it only looks at those periods where the market portfolio has a negative return.
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && mktReturns.Length < assetReturns.Length)
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
                    List<double> rfDownReturns = new List<double>();
                    double[] rf = new double[assetReturns.Length];

                    rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns, assetReturns.Length));
                    for (int i = 0; i < mktReturns.Length; i++)
                    {
                        if (mktReturns[i] < 0)
                        {
                            mktDownReturns.Add(mktReturns[i]);
                            assetDownReturns.Add(assetReturns[i]);
                            rfDownReturns.Add(rf[i]);
                        }
                    }

                    if (mktDownReturns.Count == 0)
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        //Get beta using returns only when market is negative
                        return Beta(assetDownReturns.ToArray(), mktDownReturns.ToArray(),
                            Statistics.DoubleToObject(rfDownReturns.ToArray()));
                    }
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "BetaTimingRatio", Description = "Calculates the ratio of the bull beta to the bear beta",
            Category = "Portfolio Performance")]
        //This is the ratio of the bull beta to the bear beta and provides a measure of timing ability.
        public static object BetaTimingRatio(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)]
            double[] mktReturns, [ExcelArgument(Name = "Risk-free Returns", Description = "Range of Risk-free asset Returns", AllowReference = false)]
            object[] riskFreeReturns)
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
                    double bullB = (double)BullBeta(assetReturns, mktReturns, riskFreeReturns);
                    double bearB = (double)BearBeta(assetReturns, mktReturns, riskFreeReturns);
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

        [ExcelFunction(Name = "JensensAlpha", Description = "Calculates Jensen's alpha for an asset", Category = "Portfolio Performance")]
        public static object JensensAlpha([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)]
            double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)]
            double[] mktReturns,
            [ExcelArgument(Name = "Risk-free Asset Returns", Description = "Range of Risk-free Returns", AllowReference = false)]
            object[] riskFreeReturns)
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
                    rf = Statistics.ObjToDouble(Statistics.ExtendRiskFreeRateArray(riskFreeReturns,assetReturns.Length));
                    double assetMean = assetReturns.Average();
                    double rfMean = rf.Average();
                    double mktMean = mktReturns.Average();
                    double assetBeta = (double)Beta(assetReturns, mktReturns, Statistics.DoubleToObject(rf));
                    return (assetMean - rfMean) - assetBeta * (mktMean - rfMean);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        //[ExcelFunction(Name = "ReturnArray", Description = "Test Function for returning an array", Category = "Portfolio Performance")]
        //public static object[,] ReturnArray()
        //{
        //    object[,] retarray = new object[2,1];
        //    retarray[0, 0] = 10;
        //    retarray[1, 0] = 15;
        //    return retarray;

        //}

    }
}


