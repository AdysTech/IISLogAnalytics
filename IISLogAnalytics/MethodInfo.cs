using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
    
namespace AdysTech.IISLogAnalytics
{
    class MethodInfo
    {
        public string Url { get; private set; }
        public DateTime Timestamp { get; set; }
        public long Hits { get; private set; }
        public double MinResponseTime { get { return responseTimes.Min() / 1000.0; } }
        public double MaxResponseTime { get { return responseTimes.Max() / 1000.0; } }
        public double AvgResponseTime { get { return (TotalTime / Hits) / 1000.0; } }

        public double NinetiethPercentile
        {
            get
            {
                responseTimes.Sort();
                int nth = (90 * responseTimes.Count) / 100;
                return responseTimes[nth] / 1000.0;
            }
        }
        public double Median
        {
            get
            {
                responseTimes.Sort();
                return responseTimes[responseTimes.Count / 2] / 1000.0;
            }
        }
        public double StandardDeviation
        {
            get
            {
                var avg = responseTimes.Average();               
                //Perform the Sum of (value-avg)^2
                double sum = responseTimes.Sum(d => (d - avg) * (d - avg));
                //Put it all together
                return Math.Sqrt(sum / responseTimes.Count) / 1000.0;
            }
        }

        private long TotalTime { get; set; }
        private List<int> responseTimes;
        public void Hit(int responseTime)
        {
            Hits++;
            TotalTime += responseTime;
            responseTimes.Add(responseTime);
        }
        public MethodInfo(string Url, int responseTime)
        {
            this.Url = Url;
            responseTimes = new List<int>();
            Hit(responseTime);
        }
        public static long TotalHits(IEnumerable<MethodInfo> list)
        {
            return list.Sum(p => p.Hits);
        }
    }
}
