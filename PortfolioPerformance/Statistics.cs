using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace PortfolioPerformance
{
    class Statistics
    {
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
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static double Covariance_S(double[] data1, double[] data2) //Sample Covariance
        {
            double n = data1.Length;
            return Covariance_P(data1, data2) * n / (n - 1); //adjust population covariance to sample covariance
        }

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

        public static double Variance_S(double[] data) //Sample Variance
        {
            double n = data.Length;
            return Variance_P(data) * n / (n - 1); //adjust population variance to sample variance
        }

        public static double StdDev_P(double[] data) //Population Standard Deviation
        {
            return Math.Pow(Variance_P(data), 0.5);
        }

        public static double StdDev_S(double[] data) //Sample Standard Deviation
        {
            double n = data.Length;
            return
                StdDev_P(data) *
                Math.Pow(n / (n - 1), 0.5); //adjust population standard deviation to sample standard deviation
        }

        public static double[]
            ArrayDiff(double[] arr1, double[] arr2) //Calculate element-wise difference between two equal-length arrays
        {
            double[] diff = new double[arr1.Length];
            for (int i = 0; i < arr1.Length; i++)
            {
                diff[i] = arr1[i] - arr2[i];
            }

            return diff;
        }

        public static double[] ObjToDouble(object[] arr) //Convert an object[] to double[]
        {
            int n = arr.Length;
            double[] retArray = new double[n];
            for (int i = 0; i < n; i++)
            {
                retArray[i] = (double) arr[i];
            }

            return retArray;
        }

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
    }
}
