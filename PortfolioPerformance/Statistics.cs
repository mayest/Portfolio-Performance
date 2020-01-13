using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PortfolioPerformance
{
    /// <summary>
    /// Purpose: A class of helper functions, mostly statistics but some converters.
    /// Author: Timothy R. Mayes, Ph.D.
    /// Date: 25 December 2019
    /// </summary>
    public class Statistics
    {
        [ExcelFunction(IsHidden = true)]
        public static double Covariance_P(double[] data1, double[] data2) //Population Covariance
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
        public static double Covariance_S(double[] data1, double[] data2) //Sample Covariance
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
        public static double Variance_S(double[] data) //Sample Variance
        {
            double n = data.Length;
            return Variance_P(data) * n / (n - 1); //adjust population variance to sample variance
        }

        [ExcelFunction(IsHidden = true)]
        public static double StdDev_P(double[] data) //Population Standard Deviation
        {
            return Math.Pow(Variance_P(data), 0.5);
        }

        [ExcelFunction(IsHidden = true)]
        public static double StdDev_S(double[] data) //Sample Standard Deviation
        {
            double n = data.Length;
            return
                StdDev_P(data) * Math.Pow(n / (n - 1), 0.5); //adjust population standard deviation to sample standard deviation
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
                return Skewness_P(data) * Math.Sqrt(length*(length-1))/(length-2);
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

            return sumQuads * (n*(n+1)/((n-1)*(n-2)*(n-3)));
        }

        [ExcelFunction(IsHidden = true)]
        public static double Kurtosis_S_Excess(double[] data)
        //This is what Excel will return
        {
            double n = data.Length;
            return Kurtosis_S(data) - (3 * Math.Pow(n - 1, 2)/((n - 2) * (n - 3)));
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] ArrayDiff(double[] arr1, double[] arr2) //Calculate element-wise arithmetic difference between two equal-length arrays
        {
            double[] diff = new double[arr1.Length];
            for (int i = 0; i < arr1.Length; i++)
            {
                diff[i] = arr1[i] - arr2[i];
            }

            return diff;
        }

        [ExcelFunction(IsHidden = true)]
        public static double[] ArrayDiffGeom(double[] arr1, double[] arr2) //Calculate element-wise geometric difference between two equal-length arrays
        {
            double[] diff = new double[arr1.Length];
            for (int i = 0; i < arr1.Length; i++)
            {
                diff[i] = (1 + arr1[i]) / (1 + arr2[i]) - 1;
            }

            return diff;
        }


        [ExcelFunction(IsHidden = true)]
        public static double[] ObjToDouble(object[] arr) //Convert an object[] to double[]
        {
            int n = arr.Length;
            double[] retArray = new double[n];
            for (int i = 0; i < n; i++)
            {
                retArray[i] = (double)arr[i];
            }

            return retArray;
        }

        [ExcelFunction(IsHidden = true)]
        public static object[] DoubleToObject(double[] arr) //Convert an double[] to object[]
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
                    tempRf = (double)rfArray[0]; //Otherwise, set it to the first number
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

            return Math.Pow(cumRet, frequency/returns.Length) - 1;
        }
    }
}
