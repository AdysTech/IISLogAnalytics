# IISLogAnalytics
A small console application to parse IIS logs for a period and generate a MS Excel spreadsheet with analytics like page view trend, peak hour trend etc.

This console application parses multiple IIS log files (many days, hours), and analysis them, and generates analytics. Currently implemented are:

1. Concurrent Users, found using Little's Law. To get there tps (transactions per second) and average time for each request is also calculated.
The concurrency window is configurable.

2. Page visit Summary, it tracks all pages visited, and their 
    Total Hits, response time metrics(Min(sec), Avg(sec), Max(sec), 90th-%(sec), Median(sec), Std. Deviation(sec)), and graphs.
<img src="IISLogAnalytics/Sample Graphs/Top Pages.png"/>

3. Daily summary, Top Pages - Visits Trend, Response Time(90%tile) Trend
<img src="IISLogAnalytics/Sample Graphs/Daily graphs.png"/>

4. Hourly Analysis, Top Pages (so you can see which pages overlap over time), peak hour pages.

5. URL parameters, if pages have a specifiec query parameter, that needs to be analyzed.

6. Generic statistics, like total requests by various HTTP codes etc.

Pre-Req: needs office 2010+, reference: Microsoft.Office.Interop.Excel 14.0.0.0, office 14.0.0.0
