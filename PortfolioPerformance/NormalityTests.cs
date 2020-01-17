using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;


namespace PortfolioPerformance
{
    /// <summary>
    /// Purpose: This class contains various tests to check if data follows a normal distribution.
    /// Author: Timothy R. Mayes, Ph.D.
    /// Date: 16 January 2020
    /// </summary>

    public class NormalityTests
    {
        [ExcelFunction(Name = "JarqueBeraTest", Description = "Returns a test statistic that tests for normality", Category = "Portfolio Performance")]
        public static object JarqueBeraTest(
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
                    double n = assetReturns.Length;
                    if (n > 2)
                    {
                        double skew = Helpers.Skewness_P(assetReturns);
                        double kurtExcess = Helpers.Kurtosis_P_Excess(assetReturns);
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

    }
}
