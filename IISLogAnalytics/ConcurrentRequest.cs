using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IISLogAnalytics
{
    class ConcurrentRequest
    {
        public DateTime TimeStamp { get; set; }
        public long Transactions { get; set; }
        public double AverageResponseTime { get; set; }

        public long BytesSent { get; set; }
        public double Tps { get { return Transactions / (concurrencyWindow * 60.0); } }
        public double NetworkSpeed { get { return (BytesSent * 8) / (concurrencyWindow * 60.0) / (1000 * 1000); } }
        public long ConcurrentUsers { get { var u = (long)Math.Round(Tps * (AverageResponseTime / 1000), 0); return u == 0 ? 1 : u; } }

        float concurrencyWindow;
        public ConcurrentRequest(float ConcurrencyWindow)
        {
            concurrencyWindow = ConcurrencyWindow;
        }
    }
}
