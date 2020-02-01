using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace PortfolioPerformance
{
    /// <summary>
    /// Purpose: This class contains various measures of asset/portfolio return.
    /// Author: Timothy R. Mayes, Ph.D.
    /// Date: 25 January 2020
    /// </summary>
    public class ReturnMeasures
    { 
        [ExcelFunction(Name = "HoldingPeriodReturn", Description = "Calculates the overall holding period return given asset prices and cash flows", Category = "Portfolio Performance")]
        public static object HoldingPeriodReturn(
        [ExcelArgument(Name = "Prices", Description = "The range of prices for the asset/portfolio.", AllowReference = false)] double[] prices,
        [ExcelArgument(Name = "Cash Flows", Description = "(Optional) The range of cash flows for the asset/portfolio. " +
                                                          "Note: a cash flow at period 0 is ignored.", AllowReference = false)] object[] cashFlows)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && prices.Length == 0)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    if (cashFlows.Length > prices.Length) return ExcelError.ExcelErrorValue; //Too many cash flows, return error

                    if (cashFlows.Length == 1 && cashFlows[0] is ExcelMissing) //No cash flows given
                        return prices.Last() / prices[0] - 1;
                    double[] cf = Helpers.ObjToDouble(cashFlows);
                    if (cf.Length == prices.Length)
                        return (prices.Last() + cf.Sum() - cf[0]) / prices.First() - 1; //Don't include any cash flow at period 0
                    return (prices.Last() + cf.Sum()) / prices.First() - 1;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "HPRWithReinvestment", Description = "Calculates the overall holding period return with reinvestment given asset prices and cash flows", Category = "Portfolio Performance")]
        public static object HPRWithReinvestment(
            [ExcelArgument(Name = "Prices", Description = "The range of prices for the asset/portfolio.", AllowReference = false)] double[] prices,
            [ExcelArgument(Name = "Cash Flows", Description = "(Optional) The range of cash flows for the asset/portfolio. " +
                                                              "Note: a cash flow at period 0 is ignored, and cash flows must align with prices.", AllowReference = false)] object[] cashFlows)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && prices.Length == 0)
                //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
                //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    if (cashFlows.Length > prices.Length) return ExcelError.ExcelErrorValue; //Too many cash flows, return error

                    double cumRet = 1d;
                    //No cash flows given
                    if (cashFlows.Length == 1 && cashFlows[0] is ExcelMissing)
                    {
                        for (int i = 1; i < prices.Length; i++)
                        {
                            cumRet *= prices[i] / prices[i - 1];
                        }

                        return cumRet - 1;
                    }
                    //Cash flows given and same length as prices
                    double[] cf = Helpers.ObjToDouble(cashFlows);
                    if (cf.Length == prices.Length)
                    {
                        for (int i = 1; i < prices.Length; i++)
                        {
                            cumRet *= (prices[i] + cf[i]) / prices[i - 1];
                        }

                        return cumRet - 1;
                    }
                    //Cash flows given, but fewer of them than prices.
                    //So, we assume that cf[0] goes with prices[1], cf[1] with prices[2], and so on.
                    //Note that we are ignoring any cash flow that might go with prices[0] (the beginning price).
                    else if (cf.Length < prices.Length)
                    {
                        for (int i = 1; i < prices.Length; i++)
                        {
                            if (i-1 < cf.Length)
                                cumRet *= (prices[i] + cf[i-1]) / prices[i - 1];
                            else cumRet *= prices[i] / prices[i - 1];
                        }

                        return cumRet - 1;

                    }
                    return ExcelError.ExcelErrorValue;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "SubPeriodReturns", Description = "Returns an array of returns for each period given asset prices and cash flows", Category = "Portfolio Performance")]
        public static object SubPeriodReturns(
            [ExcelArgument(Name = "Prices", Description = "The range of prices for the asset/portfolio.", AllowReference = false)] double[] prices,
            [ExcelArgument(Name = "Cash Flows", Description = "(Optional) The range of cash flows for the asset/portfolio. " +
                                                                      "Note: a cash flow at period 0 is ignored.", AllowReference = false)] object[] cashFlows)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && prices.Length == 0)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    if (cashFlows.Length > prices.Length) return ExcelError.ExcelErrorValue; //Too many cash flows, return error

                    double[,] returns = new double[prices.Length - 1, 1]; //The array that is returned - a 2D array so that we can return a column
                    
                    //No cash flows given 
                    if (cashFlows.Length == 1 && cashFlows[0] is ExcelMissing) 
                    {
                        for (int i = 1; i < prices.Length; i++)
                        {
                            returns[i - 1, 0] = prices[i] / prices[i - 1] - 1;
                        }
                        return returns;
                    }                    
                    
                    double[] cf = Helpers.ObjToDouble(cashFlows);
                    //Cash flows same length as prices
                    if (cf.Length == prices.Length)
                    {
                        for (int i = 1; i < prices.Length; i++)
                        {
                            returns[i - 1, 0] = (prices[i] + cf[i]) / prices[i - 1] - 1;
                        }

                    }
                    //Fewer cash flows than prices, so assume that cf[0] goes with prices[1], and so on.
                    else
                    {
                        for (int i = 1; i < prices.Length; i++)
                        {
                            returns[i - 1, 0] = (prices[i] + cf[i - 1]) / prices[i - 1] - 1;
                        }
                    }

                    return returns;

                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

        [ExcelFunction(Name = "LogSubPeriodReturns", Description = "Returns an array of log price relatives for each period given asset prices and cash flows", Category = "Portfolio Performance")]
        public static object LogSubPeriodReturns(
            [ExcelArgument(Name = "Prices", Description = "The range of prices for the asset/portfolio.", AllowReference = false)] double[] prices,
            [ExcelArgument(Name = "Cash Flows", Description = "(Optional) The range of cash flows for the asset/portfolio. " +
                                                                              "Note: a cash flow at period 0 is ignored.", AllowReference = false)] object[] cashFlows)
        {
            if (ExcelDnaUtil.IsInFunctionWizard() && prices.Length == 0)
            //This is required because Function Wizard repeatedly calls the function and will cause an error on partial range entry for second var
            //The check on lengths means that the Function Wizard will show a correct result when the lengths are equal
            {
                return ExcelError.ExcelErrorValue; //Return a placeholder value until both ranges are fully entered
            }
            else //Try the calculation
            {
                try
                {
                    double[,] logReturns = new double[prices.Length - 1, 1]; //The array that is returned - a 2D array so that we can return a column
                    double[,] tempReturns = (double[,])SubPeriodReturns(prices, cashFlows);
                    for (int i = 0; i < logReturns.Length; i++)
                    {
                        logReturns[i, 0] = Math.Log(1+tempReturns[i,0], Math.E);
                    }
                    return logReturns;
                }
                catch (Exception)
                {
                    return ExcelError.ExcelErrorValue;
                }
            }

        }

    }

}
