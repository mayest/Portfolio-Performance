using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortfolioPerformance
{
    class Statistics
    {
        public static double Covar_P(double[] data1, double[] data2) //Population Covariance
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
            return Covar_P(data1, data2)*n/(n-1); //adjust population covariance to sample covariance
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
            return StdDev_P(data)*Math.Pow(n/(n-1),0.5); //adjust population standard deviation to sample standard deviation
        }

        public static double[] ArrayDiff(double[] arr1, double[] arr2) //Calculate element-wise difference between two equal-length arrays
        {
            double[] diff = new double[arr1.Length];
            for (int i = 0; i < arr1.Length; i++)
            {
                diff[i] = arr1[i] - arr2[i];
            }
            return diff;
        }

    }
}
