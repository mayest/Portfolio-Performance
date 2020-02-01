using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PortfolioPerformance
{
    /// <summary>
    /// Purpose: A class of helper functions, mostly statistics but some converters and other helpers.
    ///          These functions are meant for internal use only, so they are marked as hidden from Excel.
    ///          For now, they are public for testing purposes.
    /// Author: Timothy R. Mayes, Ph.D.
    /// Date: 16 January 2020
    /// </summary>
    public class Helpers
    {
        [ExcelFunction(IsHidden = true)]
        public static double Covariance_P(double[] data1, double[] data2) 
            //Population Covariance
        {
            try
            {
                double n = data1.Length;
                double accumulator = 0;
                double mean1 = data1.Average();
                double mean2 = data2.Average();
                for (int i = 0; i <= n - 1; i++)
                {
                    accumulator += (data1[i] - mean1) * (data2[i] - mean2);
                }

                return accumulator / n;
            }
            catch (Exception)
            {
                throw;
            }
        }

        [ExcelFunction(IsHidden = true)]
        public static double Covariance_S(double[] data1, double[] data2) 
            //Sample Covariance
        {
            double n = data1.Length;
            return Covariance_P(data1, data2) * n / (n - 1); //adjust population covariance to sample covariance
        }

        [ExcelFunction(IsHidden = true)]
        public static double Variance_P(double[] data) //Population Variance
        {
            double ssq = 0; // sum of squares
            double dataMean = data.Average();
            foreach (var d in data)
            {
                ssq += Math.Pow(d - dataMean, 2);
            }

            return ssq / data.Length;
        }

        [ExcelFunction(IsHidden = true)]
        public static double Variance_S(double[] data) 
            //Sample Variance
        {
            double n = data.Length;
            return Variance_P(data) * n / (n - 1); //adjust population variance to sample variance
        }

        [ExcelFunction(IsHidden = true)]
        public static double StdDev_P(double[] data) 
            //Population Standard Deviation
        {
            return Math.Pow(Variance_P(data), 0.5);
        }

        [ExcelFunction(IsHidden = true)]
        public static double StdDev_S(double[] data) 
            //Sample Standard Deviation
        {
            double n = data.Length;
            return
                StdDev_P(data) *
                Math.Pow(n / (n - 1), 0.5); //adjust population standard deviation to sample standard deviation
        }

        [ExcelFunction(IsHidden = true)]
        public static double LowerPartialMoment_P(double[] data, double targetReturn, double degree)
        {
            double x = 0;
            foreach (var d in data)
            {
                if (d < targetReturn) //Ignore anything >= target return
                {
                    x += Math.Pow(Math.Max(0, targetReturn - d), degree);
                }

            }

            return x / data.Length;
        }

        [ExcelFunction(IsHidden = true)]
        public static double UpperPartialMoment_P(double[] data, double targetReturn, double degree)
        {
            double x = 0;
            foreach (var d in data)
            {
                if (d > targetReturn) //Ignore anything <= target return
                {
                    x += Math.Pow(Math.Max(0, d - targetReturn), degree);
                }

            }

            return x / data.Length;
        }



        [ExcelFunction(IsHidden = true)]
        public static double SemiVariance_P(double[] data, double targetReturn)
        {
            List<double> belowMean = new List<double>();
            foreach (var d in data)
            {
                if (d < targetReturn)
                {
                    belowMean.Add(d);
                }
            }

            return Variance_P(belowMean.ToArray());
        }

        [ExcelFunction(IsHidden = true)]
        public static double SemiDeviation_P(double[] data, double targetReturn)
        {
            return Math.Sqrt(SemiVariance_P(data, targetReturn));
        }

        [ExcelFunction(IsHidden = true)]
        public static double Skewness_P(double[] data)
            //Calculate population skewness
            //See https://www.itl.nist.gov/div898/handbook/eda/section3/eda35b.htm
        {
            double dataMean = data.Average();
            double dataSD = StdDev_P(data);
            double sumCubes = 0;
            foreach (var d in data)
            {
                sumCubes += Math.Pow((d - dataMean) / dataSD, 3);
            }

            return sumCubes / data.Length;
        }

        [ExcelFunction(IsHidden = true)]
        public static double Skewness_S(double[] data)
        {
            long length = data.Length;
            try
            {
                return Skewness_P(data) * Math.Sqrt(length * (length - 1)) / (length - 2);
            }
            catch (Exception)
            {

                throw;
            }
        }

        [ExcelFunction(IsHidden = true)]
        public static double Kurtosis_P(double[] data)
        {
            double dataMean = data.Average();
            double dataSD = StdDev_P(data);
            double sumQuads = 0;
            foreach (var d in data)
            {
                sumQuads += Math.Pow((d - dataMean) / dataSD, 4);
            }

            return sumQuads / data.Length;
        }

        [ExcelFunction(IsHidden = true)]
        public static double Kurtosis_P_Excess(double[] data)
        {
            return Kurtosis_P(data) - 3;
        }

        [ExcelFunction(IsHidden = true)]
        public static double Kurtosis_S(double[] data)
        {
            double dataMean = data.Average();
            double dataSD = StdDev_S(data);
            double n = data.Length;
            double sumQuads = 0;
            foreach (var d in data)
            {
                sumQuads += Math.Pow((d - dataMean) / dataSD, 4);
            }

            return sumQuads * (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)));
        }

        [ExcelFunction(IsHidden = true)]
        public static double Kurtosis_S_Excess(double[] data)
            //This returns the same result as Excel's Kurt function
        {
            double n = data.Length;
            return Kurtosis_S(data) - (3 * Math.Pow(n - 1, 2) / ((n - 2) * (n - 3)));
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] ArrayDiff(double[] arr1, double[] arr2) 
            //Calculate element-wise arithmetic difference between two equal-length arrays
        {
            double[] diff = new double[arr1.Length];
            for (int i = 0; i < arr1.Length; i++)
            {
                diff[i] = arr1[i] - arr2[i];
            }

            return diff;
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] ArrayDiffGeom(double[] arr1, double[] arr2) 
            //Calculate element-wise geometric difference between two equal-length arrays
        {
            double[] diff = new double[arr1.Length];
            for (int i = 0; i < arr1.Length; i++)
            {
                diff[i] = (1 + arr1[i]) / (1 + arr2[i]) - 1;
            }

            return diff;
        }


        [ExcelFunction(IsHidden = true)]
        internal static double[] ObjToDouble(object[] arr) //Convert an object[] to double[]
        {
            int n = arr.Length;
            double[] retArray = new double[n];
            for (int i = 0; i < n; i++)
            {
                if (arr[i] is ExcelEmpty)
                {
                    arr[i] = 0d; //If it is empty/missing, set it to 0
                }
                retArray[i] = (double)arr[i];
            }

            return retArray;
        }

        [ExcelFunction(IsHidden = true)]
        internal static object[] DoubleToObject(double[] arr) //Convert an double[] to object[]
        {
            int n = arr.Length;
            object[] retArray = new object[n];
            for (int i = 0; i < n; i++)
            {
                retArray[i] = arr[i];
            }

            return retArray;
        }

        [ExcelFunction(IsHidden = true)]
        internal static double[,] ConvertToColumnArray(double[] arr)
            //By default, ExcelDNA object arrays will be row-oriented. Most often, I want it to be in a column instead of a row.
            //This will effectively transpose the result be returning a 2D array with nothing in the second dimension.
            //Excel will NOT fill in the second column. The result will only appear as a single column.
        {
            double[,] retArray = new double[arr.Length, 1];
            for (int i = 0; i < arr.Length; i++) //This changes the array so that it will be a column array in Excel
            {
                retArray[i, 0] = arr[i];
            }

            return retArray;
        }

        [ExcelFunction(IsHidden = true)]
        public static object[] ExtendRiskFreeRateArray(object[] rfArray, int length)
            //Extend the risk-free rate array to the given length
        {
            object[] newRf = new object[length]; //New temporary array

            // Check if riskFreeReturns is given or not. If not, set it to 0 for each period.
            // If it is only a single number, then set it to that number for each period.
            // If there is more than one number, but less than length, then set it to the first number.
            if (rfArray.Length != length)
            {
                double tempRf;
                if (rfArray[0] is ExcelMissing)
                {
                    tempRf = 0.0d; //Set to 0 if it is missing
                }
                else
                {
                    tempRf = (double) rfArray[0]; //Otherwise, set it to the first number
                }

                for (int i = 0; i < newRf.Length; i++)
                {
                    newRf[i] = tempRf;
                }

                return newRf;
            }
            else
            {
                return rfArray; //If it matches length, then just return it unchanged
            }
        }

        [ExcelFunction(IsHidden = true)]
        public static double AnnualizedReturn(double[] returns, double frequency)
        {
            double cumRet = 1;
            foreach (var ret in returns)
            {
                cumRet *= (1 + ret);
            }

            return Math.Pow(cumRet, frequency / returns.Length) - 1;
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] GetDrawDowns(double[] returns) 
            //Returns an array of all drawdowns in the returns series
        {
            Int32 n = returns.Length;
            double[] cumRets = new double[n + 1];
            double[] drawDowns = new double[n + 1];
            for (int i = 0; i <= n; i++) //Build total return index
            {
                double maxDd = (i == 0) ? 1 : cumRets.Max();
                cumRets[i] = (i == 0) ? 1 : cumRets[i - 1] * (1 + returns[i - 1]);
                if (i >= 1)
                {
                    drawDowns[i] = cumRets[i] / maxDd - 1;
                    drawDowns[i] = (drawDowns[i] > 0) ? 0d : drawDowns[i]; //Set to 0 if greater than 0
                }
                else
                {
                    drawDowns[i] = 0;
                }
            }

            return drawDowns;
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] GetContinuousDrawDowns(double[] returns)
            //Returns an array containing the maximum drawdowns for each drawdown period in the returns series
        {
            //Get the product of each sequence of drawdowns
            double product = 0d;
            List<double> continuousDrawDowns = new List<double>();
            for (int i = 0; i < returns.Length; i++)
            {
                while (i < returns.Length && returns[i] < 0)
                {
                    product = (1 + product) * (1 + returns[i]) - 1;
                    i++; 
                }

                if (product < 0) continuousDrawDowns.Add(product);
                product = 0; //Reset the product
            }

            return continuousDrawDowns.ToArray();
        }

        [ExcelFunction(IsHidden = true)]
        public static object[,] SplitToYears(double[] returns, int frequency)
        //Takes in an array of returns and divides them into years based on the return frequency.
        //Note that this assumes that the first return is at the start of the year.
        //The resize method will result in padding the array with zeros.
        {
            Int32 remainder = returns.Length % frequency;
            if (remainder > 0)
            {
                Array.Resize(ref returns, returns.Length + frequency - remainder); //Resize to an even multiple of frequency
            }

            Int32 numSeries = (int)(returns.Length / frequency);
            object[,] retArray = new object[numSeries, frequency];
            for (int i = 0; i < numSeries; i++)
            {
                for (int j = i * frequency; j < (i * frequency + frequency); j++)
                {
                    retArray[i, j - i * frequency] = returns[j];
                }
            }
            return retArray;
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] GetUniqueYears(double[] dates)
        //Takes in an array of Excel dates and returns an array of unique years
        {
            double[] years = new double[dates.Length];
            for (int i = 0; i < dates.Length; i++)
            {
                years[i] = DateTime.FromOADate(dates[i]).Year;
            }

            double[] uniqueYears = years.ToHashSet().ToArray();//Convert to HashSet to get unique years
            return uniqueYears;

        }

        [ExcelFunction(IsHidden = true)]
        public static double[] GetPeaks(double[] returns) //Returns an array containing a list of the peaks in a return series
        {
            List<double> totalReturnIndex = new List<double>();
            List<double> peaks = new List<double>();
            totalReturnIndex.Add(1d);
            for (int i = 1; i < returns.Length+1; i++)
            {
                totalReturnIndex.Add((totalReturnIndex[i - 1])*(1 + returns[i - 1]));
            }
            for (int i = 1; i < totalReturnIndex.Count; i++)
            {
                if (i == 1)
                {
                    if (totalReturnIndex[i] > totalReturnIndex[i+1])
                    {
                        peaks.Add(totalReturnIndex[i]);
                    }
                    else
                    {
                        peaks.Add((double)0d);
                    }
                }
                else if (i < totalReturnIndex.Count-2 && totalReturnIndex[i] > totalReturnIndex[i-1] && totalReturnIndex[i] > totalReturnIndex[i+1])
                {
                    peaks.Add(totalReturnIndex[i]);
                }
                else if (i == totalReturnIndex.Count-1 && totalReturnIndex[i] > totalReturnIndex[i-1])
                {
                    peaks.Add(totalReturnIndex[i]);
                }
                else
                {
                    peaks.Add((double)0d);
                }
            }
            return peaks.ToArray();
        }
        
        [ExcelFunction(IsHidden = true)]
        public static double[] GetTroughs(double[] returns) //Returns an array containing a list of the troughs in a return series
        {
            List<double> totalReturnIndex = new List<double>();
            List<double> troughs = new List<double>();
            totalReturnIndex.Add(1d);
            for (int i = 1; i < returns.Length + 1; i++)
            {
                totalReturnIndex.Add((totalReturnIndex[i - 1]) * (1 + returns[i - 1]));
            }
            for (int i = 1; i < totalReturnIndex.Count; i++)
            {
                if (i == 1)
                {
                    if (totalReturnIndex[i] < totalReturnIndex[i + 1])
                    {
                        troughs.Add(totalReturnIndex[i]);
                    }
                    else
                    {
                        troughs.Add((double)0d);
                    }
                }
                else if (i < totalReturnIndex.Count - 2 && totalReturnIndex[i] < totalReturnIndex[i - 1] && totalReturnIndex[i] < totalReturnIndex[i + 1])
                {
                    troughs.Add(totalReturnIndex[i]);
                }
                else if (i == totalReturnIndex.Count - 1 && totalReturnIndex[i] < totalReturnIndex[i - 1])
                {
                    troughs.Add(totalReturnIndex[i]);
                }
                else
                {
                    troughs.Add((double)0d);
                }
            }
            return troughs.ToArray();
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] GetTotalReturnIndex(double[] returns, double startValue) //Returns an array containing a list of the troughs in a return series
        {
            Int32 n = returns.Length;
            double[] totalReturnIdx = new double[n + 1];
            for (int i = 0; i <= n; i++) //Build total return index
            {
                totalReturnIdx[i] = (i == 0) ? startValue : totalReturnIdx[i - 1] * (1 + returns[i - 1]);
            }

            return totalReturnIdx;
        }

        [ExcelFunction(IsHidden = true)]
        public static double NormalCdfInverse(double p, double mu, double sigma)
        {
            // Calculates the inverse of the normal CDF given a probability, mean, and standard deviation.
            // This code agrees with Excel to at least 1e-14.
            // There is no closed-form solution to the inverse CDF for the normal
            // distribution, so we use a rational approximation instead:
            // Wichura, M.J. (1988). "Algorithm AS241: The Percentage Points of the
            // Normal Distribution".  Applied Statistics. Blackwell Publishing. 37(3): 477–484. doi:10.2307/2347330. JSTOR 2347330.
            // Code translated from Python source at https://github.com/python/cpython/blob/master/Lib/statistics.py
            double q = p - 0.5;
            if (Math.Abs(q) <= 0.425)
            {
                double r = 0.180625 - q * q;
                double num = (((((((2.50908_09287_30122_6727e+3 * r + 3.34305_75583_58812_8105e+4) * r +
                                   6.72657_70927_00870_0853e+4) * r + 4.59219_53931_54987_1457e+4) * r +
                                 1.37316_93765_50946_1125e+4) * r + 1.97159_09503_06551_4427e+3) * r +
                               1.33141_66789_17843_7745e+2) * r + 3.38713_28727_96366_6080e+0) * q;
                double den = (((((((5.22649_52788_52854_5610e+3 * r + 2.87290_85735_72194_2674e+4) * r +
                                   3.93078_95800_09271_0610e+4) * r + 2.12137_94301_58659_5867e+4) * r +
                                 5.39419_60214_24751_1077e+3) * r + 6.87187_00749_20579_0830e+2) * r +
                               4.23133_30701_60091_1252e+1) * r + 1.0);
                double x = num / den;
                return mu + (x * sigma);
            }

            double y = (q <= 0) ? Math.Sqrt(-Math.Log(p)) : Math.Sqrt(-Math.Log(1 - p));
            if (y <= 5.0)
            {
                y -= 1.6;
                double num = (((((((7.74545_01427_83414_07640e-4 * y + 2.27238_44989_26918_45833e-2) * y +
                                   2.41780_72517_74506_11770e-1) * y + 1.27045_82524_52368_38258e+0) * y +
                                 3.64784_83247_63204_60504e+0) * y + 5.76949_72214_60691_40550e+0) * y +
                               4.63033_78461_56545_29590e+0) * y + 1.42343_71107_49683_57734e+0);
                double den = (((((((1.05075_00716_44416_84324e-9 * y + 5.47593_80849_95344_94600e-4) * y +
                                   1.51986_66563_61645_71966e-2) * y + 1.48103_97642_74800_74590e-1) * y +
                                 6.89767_33498_51000_04550e-1) * y + 1.67638_48301_83803_84940e+0) * y +
                               2.05319_16266_37758_82187e+0) * y + 1.0);
                double x = num / den;
                if (q < 0)
                {
                    x = -x;
                }

                return mu + (x * sigma);

            }
            else
            {
                y -= 5.0;
                double num = (((((((2.01033_43992_92288_13265e-7 * y + 2.71155_55687_43487_57815e-5) * y +
                                   1.24266_09473_88078_43860e-3) * y + 2.65321_89526_57612_30930e-2) * y +
                                 2.96560_57182_85048_91230e-1) * y + 1.78482_65399_17291_33580e+0) * y +
                               5.46378_49111_64114_36990e+0) * y + 6.65790_46435_01103_77720e+0);
                double den = (((((((2.04426_31033_89939_78564e-15 * y + 1.42151_17583_16445_88870e-7) * y +
                                   1.84631_83175_10054_68180e-5) * y + 7.86869_13114_56132_59100e-4) * y +
                                 1.48753_61290_85061_48525e-2) * y + 1.36929_88092_27358_05310e-1) * y +
                               5.99832_20655_58879_37690e-1) * y + 1.0);
                double x = num / den;
                if (q < 0)
                {
                    x = -x;
                }

                return mu + (x * sigma);
            }

        }


    }

}
