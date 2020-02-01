using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace PortfolioPerformance
{
    /// <summary>
    /// Purpose: This class contains various measures of asset/portfolio risk.
    /// Author: Timothy R. Mayes, Ph.D.
    /// Date: 16 January 2020
    /// </summary>
    public class RiskMeasures
    {
        [ExcelFunction(Name = "Beta", Description = "Calculates the Beta (index of systematic risk) of an Asset/Portfolio", Category = "Portfolio Performance")]
        public static object Beta(
        [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
        [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns)
        //Calculate the beta as cov(assetReturns, mktReturns)/var(mktReturns)
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
                    double cov = Helpers.Covariance_P(assetReturns, mktReturns);
                    double mktVar = Helpers.Variance_P(mktReturns);
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

        [ExcelFunction(Name = "MarketRisk", Description = "Calculates the market (systematic) risk of a set of asset returns", Category = "Portfolio Performance")]
        public static object MarketRisk(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns,
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
                    double beta = (double)Beta(assetReturns, mktReturns);
                    double mktVariance = Helpers.Variance_P(mktReturns) * freq;
                    return Math.Pow(beta, 2) * mktVariance;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "UniqueRisk", Description = "Calculates the unique (diversifiable) risk of a set of asset returns", Category = "Portfolio Performance")]
        public static object UniqueRisk(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Market Returns", Description = "Range of Market Returns", AllowReference = false)] double[] mktReturns,
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
                    double assetVariance = Helpers.Variance_P(assetReturns) * freq;
                    double mktRisk = (double) MarketRisk(assetReturns, mktReturns, (object)freq);
                    return assetVariance - mktRisk;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }
        }

        [ExcelFunction(Name = "LowerPartialMoment", Description = "Calculates a lower partial moment of a set of returns", Category = "Portfolio Performance")]
        public static object LowerPartialMoment(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Target Return", Description = "(Optional) The target return. Only returns less than the target are used in the calculation. " +
                                                                 "If omitted, the mean is used as the target.", AllowReference = false)] object target,
            [ExcelArgument(Name = "Degree", Description = "(Optional)The degree of the LPM (default is 2 for semi-variance below the target). Must be greater than or equal to 0.", AllowReference = false)] object degree,
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
                    double tgt = (target is ExcelMissing) ? assetReturns.Average() : (double)target; //Set the target return, by default mean
                    double deg = (degree is ExcelMissing) ? 2.0d : (double)degree; //Set the target degree, by default 2 for semi-variance
                    if (deg < 0)
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                    return Helpers.LowerPartialMoment_P(assetReturns, tgt, deg) * freq;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "UpperPartialMoment", Description = "Calculates a lower partial moment of a set of returns", Category = "Portfolio Performance")]
        public static object UpperPartialMoment(
    [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
    [ExcelArgument(Name = "Target Return", Description = "(Optional) The target return. Only returns greater than the target are used in the calculation. " +
                                                                 "If omitted, the mean is used as the target.", AllowReference = false)] object target,
    [ExcelArgument(Name = "Degree", Description = "(Optional)The degree of the UPM (default is 2 for semi-variance above the target). Must be greater than or equal to 0.", AllowReference = false)] object degree,
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
                    double tgt = (target is ExcelMissing) ? assetReturns.Average() : (double)target; //Set the target return, by default mean
                    double deg = (degree is ExcelMissing) ? 2.0d : (double)degree; //Set the target degree, by default 2 for semi-variance
                    if (deg < 0)
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                    return Helpers.UpperPartialMoment_P(assetReturns, tgt, deg) * freq;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }


        [ExcelFunction(Name = "SemiVariance", Description = "Calculates the semi-variance of a set of returns", Category = "Portfolio Performance")]
        public static object SemiVariance(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Target Return", Description = "(Optional) The target return. Only returns less than the target are used in the calculation. " +
                                                                 "If omitted, the mean is used as the target.", AllowReference = false)] object target,
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
                    return LowerPartialMoment(assetReturns, target, 2.0d, frequency);
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "SemiDeviation", Description = "Calculates the semi-variance of a set of returns", Category = "Portfolio Performance")]
        public static object SemiDeviation(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Target Return", Description = "(Optional) The target return. Only returns less than the target are used in the calculation. " +
                                                                 "If omitted, the mean is used as the target.", AllowReference = false)] object target,
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
                    return Math.Pow((double) SemiVariance(assetReturns, target, frequency), 0.5);//Square root of semi-variance
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "ParametricVaR", Description = "Calculates Value at Risk assuming normally distributed returns", Category = "Portfolio Performance")]
        public static object ParametricVaR([ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Confidence Level", Description = "(Optional) The confidence level (e.g., 0.95 for 95%). If omitted, 0.95 is used.", AllowReference = false)] object confLevel)
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
                    double alpha = (confLevel is ExcelMissing) ? 0.05d : (1 - (double)confLevel); //Set the frequency
                    return assetReturns.Average() + Helpers.NormalCdfInverse(alpha, 0, 1) * Helpers.StdDev_P(assetReturns); //We add because the inverse cdf will be negative
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "ModifiedParametricVaR", Description = "Calculates parametric Value at Risk adjusted for skewness and kurtosis", Category = "Portfolio Performance")]
        public static object ModifiedParametricVaR(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Confidence Level", Description = "(Optional) The confidence level (e.g., 0.95 for 95%). If omitted, 0.95 is used.", AllowReference = false)] object confLevel)
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
                    double alpha = (confLevel is ExcelMissing) ? 0.05d : (1 - (double)confLevel); //Set the alpha
                    double invCDF = Helpers.NormalCdfInverse(alpha, 0, 1);
                    double adjFactor = invCDF + ((Math.Pow(invCDF, 2) - 1) / 6) * Helpers.Skewness_P(assetReturns) +
                                       ((Math.Pow(invCDF, 3) - 3 * invCDF) / 24) * Helpers.Kurtosis_P_Excess(assetReturns) -
                                       (2 * (Math.Pow(invCDF, 3) - 5 * invCDF) / 36) * Math.Pow(Helpers.Skewness_P(assetReturns), 2);

                    //Note that we use excess kurtosis here:
                    return assetReturns.Average() + adjFactor * Helpers.StdDev_P(assetReturns); //We add because the inverse cdf will be negative
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "HistoricalSimulationVaR", Description = "Calculates Value at Risk using the historical simulation method", Category = "Portfolio Performance")]
        public static object HistoricalSimulationVaR(
            [ExcelArgument(Name = "Asset Returns", Description = "Range of Asset Returns", AllowReference = false)] double[] assetReturns,
            [ExcelArgument(Name = "Confidence Level", Description = "(Optional) The confidence level (e.g., 0.95 for 95%). If omitted, 0.95 is used.", AllowReference = false)] object confLevel)
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
                    double alpha = (confLevel is ExcelMissing) ? 0.05d : (1 - (double)confLevel); //Set the alpha
                    Array.Sort(assetReturns);
                    double targetPos = alpha * (assetReturns.Length + 1) - 1; //Position that we want
                    if (targetPos < 0 || targetPos > (assetReturns.Length - 1))
                    {
                        //Can't interpolate if the position is less than 0 or greater than n - 1.
                        //Return #NUM! consistent with Percentile.Exc().
                        return ExcelError.ExcelErrorNum;
                    }
                    Int32 nextPos = (Int32)Math.Floor(targetPos); //Nearest integer position less than targetPos
                    return assetReturns[nextPos] + (assetReturns[nextPos + 1] - assetReturns[nextPos]) * (targetPos - nextPos); //Interpolated result

                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

    }
}
