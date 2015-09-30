//Copyright: Adarsha @ AdysTech.com

using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AdysTech.IISLogAnalytics
{
    class Program
    {
        const string IgnoreSwitch = "-ignore";
        const string FilterSwitch = "-include";
        const string ExportFileSwitch = "-export";
        const string ConcurrencySwitch = "-concurrency";
        const string TopPagesSwitch = "-toppages";
        const string URLParamsSwitch = "-param";
        const string PeakHoursSwitch = "-peaks";
        const string FolderSwitch = "-folder";


        static Application excelApp = null;
        static Workbook reportSpreadsheet = null;
        static Worksheet reportSheet = null;


        static int Main(string[] args)
        {
            #region Command Line argument processing
            if ( args.Contains ("--help") )
            {
                Console.WriteLine ("This tool depends on Microsoft Office 2010+");
                Console.WriteLine ("Valid switches are");
                Console.WriteLine ("-ignore <comma separated list of file extn>    : Ignore line with pattern");
                Console.WriteLine ("-include <comma separated list of file extn>   : Filter for pattern");
                Console.WriteLine ("-concurrency <No of minutes>                   : Concurrency Window in minutes");
                Console.WriteLine ("-toppages <No of pages>                        : No of Top Pages/day");
                Console.WriteLine ("-peaks <No of peaks>                           : No of Peak Hours to consider");
                Console.WriteLine ("-param <comma seperated list of patterns>      : Summarize specific URL parameters");
                Console.WriteLine ("-export <export filename>                      : Excel file report name, default will be with time stamp");
                Console.WriteLine ("-folder <log file folder path>                 : Current folder will be defaulted. All .log files in this folder will be processed.");
                Console.WriteLine ("Add a space after the pattern if you want extension mapping (e.g. .aspx ,.jpg)");
                return 0;
            }

            if ( args.Length % 2 != 0 )
            {
                throw new ArgumentException ("Command line arguments not valid, try --help to see valid ones!");
            }

            Dictionary<string, string> cmdArgs = new Dictionary<string, string> ();
            for ( int i = 0; i < args.Length; i += 2 )
            {
                cmdArgs.Add (args[i].ToLower (), args[i + 1]);
            }


            List<string> ignoredTypes = new List<string> (), filterTypes = new List<string> (), hitsPerURLParams = new List<string> ();
            if ( cmdArgs.ContainsKey (IgnoreSwitch) )
            {
                ignoredTypes = cmdArgs[IgnoreSwitch].ToLower ().Split (',').ToList ();
            }

            if ( cmdArgs.ContainsKey (FilterSwitch) )
            {
                filterTypes = cmdArgs[FilterSwitch].ToLower ().Split (',').ToList ();
            }

            if ( cmdArgs.ContainsKey (URLParamsSwitch) )
            {
                hitsPerURLParams = cmdArgs[URLParamsSwitch].ToLower ().Split (',').ToList ();
            }


            float concurrencyWindow = 5;
            if ( cmdArgs.ContainsKey (ConcurrencySwitch) )
            {
                concurrencyWindow = float.Parse (cmdArgs[ConcurrencySwitch]);
            }
            else
                cmdArgs.Add (ConcurrencySwitch, concurrencyWindow.ToString ());

            int topPagesPerDay = 10;
            if ( cmdArgs.ContainsKey (TopPagesSwitch) )
            {
                topPagesPerDay = int.Parse (cmdArgs[TopPagesSwitch]);
            }
            else
                cmdArgs.Add (TopPagesSwitch, topPagesPerDay.ToString ());

            int peakHoursCount = 3;
            if ( cmdArgs.ContainsKey (PeakHoursSwitch) )
            {
                peakHoursCount = int.Parse (cmdArgs[PeakHoursSwitch]);
            }
            else
                cmdArgs.Add (PeakHoursSwitch, peakHoursCount.ToString ());


            string exportFileName = null;
            if ( cmdArgs.ContainsKey (ExportFileSwitch) )
            {
                try
                {
                    exportFileName = Path.GetFullPath (cmdArgs[ExportFileSwitch]);
                }
                catch ( Exception e )
                {
                    Console.WriteLine ("Error creating report file:{0},{1}", e.GetType ().Name, e.Message);
                }
            }
            if ( exportFileName == null )
            {
                exportFileName = Path.GetFullPath ("Processing results_" + DateTime.Now.ToString ("dd_hh_mm") + ".xlsx");
                Console.WriteLine ("Writing output to {0}", exportFileName);
            }

            string curerntPath;
            if ( cmdArgs.ContainsKey (FolderSwitch) )
            {
                try
                {
                    curerntPath = Path.GetFullPath (cmdArgs[FolderSwitch]);
                }
                catch ( Exception e )
                {
                    Console.WriteLine ("Error accessing folder {0}:{1},{2}", cmdArgs[FolderSwitch], e.GetType ().Name, e.Message);
                    return 1;
                }
            }
            else
            {
                curerntPath = Directory.GetCurrentDirectory ();
                Console.WriteLine ("Working on IIS logs from current folder {0}", curerntPath);
            }
            #endregion

            Stopwatch stopWatch = new Stopwatch ();
            stopWatch.Start ();



            //var files = Directory.GetFiles(curerntPath, "*.log").ToList();
            var files = new DirectoryInfo (curerntPath)
                        .GetFiles ("*.log")
                        .OrderBy (f => f.LastWriteTime)
                        .Select (f => f.FullName)
                        .ToArray ();
            var totalFile = files.Count ();

            if ( totalFile == 0 )
            {
                Console.WriteLine ("No log files found!!");
                return 0;
            }

            Console.WriteLine ("Found {0} log files", totalFile);

            var tmpFile = System.IO.Path.GetTempFileName ();
            int fileCount = 0;
            int headerRows = 4;
            int entryCount = 0;



            List<IISLogEntry> processingList = new List<IISLogEntry> ();
            DateTime nextTime = DateTime.MinValue;

            long TotalHits = 0, ServedRequests = 0;
            List<ConcurrentRequest> requests = new List<ConcurrentRequest> ();
            HashSet<string> uniqueIPs = new HashSet<string> ();
            Dictionary<int, int> httpStatus = new Dictionary<int, int> ();
            Dictionary<string, MethodInfo> pageViewsForPeriod = new Dictionary<string, MethodInfo> ();

            int totalDays = 0, totalHours = 0;

            Dictionary<string, MethodInfo> pageViewsDaily = new Dictionary<string, MethodInfo> ();
            HashSet<MethodInfo> dailyPages = new HashSet<MethodInfo> ();

            Dictionary<string, MethodInfo> pageViewsHourly = new Dictionary<string, MethodInfo> ();
            HashSet<MethodInfo> hourlyPages = new HashSet<MethodInfo> ();

            //hits for key URL parameters
            Dictionary<string, MethodInfo> urlParamHits = new Dictionary<string, MethodInfo> ();
            DateTime firstEntry = DateTime.MinValue, lastEntry = DateTime.MinValue;

            //placeholder
            HashSet<MethodInfo> filteredEntries = new HashSet<MethodInfo> ();
            int startRow = 1, startCol = 1;
            int reportRow = 2, reportCol = 1;


            Console.WriteLine ("Preparing to Process..");


            foreach ( var f in files )
            {
                try
                {

                    ++fileCount;
                    var progress = fileCount * 100 / totalFile;

                    IEnumerable<string> matchedEntries = null;

                    var contents = File.ReadLines (f);


                    Dictionary<string, int> fieldIndex = new Dictionary<string, int> ();

                    #region Content filter


                    if ( filterTypes.Any () && ignoredTypes.Any () )
                        matchedEntries = contents.Where (s => s.StartsWith ("#") ||
                            ( filterTypes.Any (x => s.ToLower ().Contains (x)) &&
                            !ignoredTypes.Any (x => s.ToLower ().Contains (x)) ));

                    else if ( filterTypes.Any () )
                        matchedEntries = contents.Where (s => s.StartsWith ("#") || filterTypes.Any (x => s.ToLower ().Contains (x)));

                    else if ( ignoredTypes.Any () )
                        matchedEntries = contents.Where (s => s.StartsWith ("#") || !ignoredTypes.Any (x => s.ToLower ().Contains (x)));
                    else
                        matchedEntries = contents;


                    foreach ( var rawLogEntry in matchedEntries )
                    {

                        IISLogEntry logEntry;
                        if ( rawLogEntry.StartsWith ("#") )
                        {
                            if ( rawLogEntry.StartsWith ("#Fields:") )
                                fieldIndex = ParseHeaderFields (rawLogEntry);
                        }
                        else
                        {
                            Console.Write ("\r{0} File {1} of {2} files ({3}%), processing {4}      ", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"), fileCount, totalFile, progress, ++TotalHits);

                            var columns = rawLogEntry.Split (' ');
                            logEntry = new IISLogEntry ()
                            {
                                TimeStamp = DateTime.Parse (columns[0] + " " + columns[1]),
                                ClientIPAddress = fieldIndex.ContainsKey (IISLogEntry.propClientIPAddress) ? columns[fieldIndex[IISLogEntry.propClientIPAddress]] : String.Empty,
                                UserName = fieldIndex.ContainsKey (IISLogEntry.propUserName) ? columns[fieldIndex[IISLogEntry.propUserName]] : String.Empty,
                                ServiceNameandInstanceNumber = fieldIndex.ContainsKey (IISLogEntry.propServiceNameandInstanceNumber) ? columns[fieldIndex[IISLogEntry.propServiceNameandInstanceNumber]] : String.Empty,
                                ServerName = fieldIndex.ContainsKey (IISLogEntry.propServerName) ? columns[fieldIndex[IISLogEntry.propServerName]] : String.Empty,
                                ServerIPAddress = fieldIndex.ContainsKey (IISLogEntry.propServerIPAddress) ? columns[fieldIndex[IISLogEntry.propServerIPAddress]] : String.Empty,
                                ServerPort = fieldIndex.ContainsKey (IISLogEntry.propClientIPAddress) ? Int32.Parse (columns[fieldIndex[IISLogEntry.propServerPort]]) : 0,
                                Method = fieldIndex.ContainsKey (IISLogEntry.propMethod) ? columns[fieldIndex[IISLogEntry.propMethod]] : String.Empty,
                                URIStem = fieldIndex.ContainsKey (IISLogEntry.propURIStem) ? columns[fieldIndex[IISLogEntry.propURIStem]] : String.Empty,
                                URIQuery = fieldIndex.ContainsKey (IISLogEntry.propURIQuery) ? columns[fieldIndex[IISLogEntry.propURIQuery]] : String.Empty,
                                HTTPStatus = fieldIndex.ContainsKey (IISLogEntry.propHTTPStatus) ? Int32.Parse (columns[fieldIndex[IISLogEntry.propHTTPStatus]]) : 0,
                                //Win32Status = fieldIndex.ContainsKey(IISLogEntry.propWin32Status) ? Int32.Parse(row[fieldIndex[IISLogEntry.propWin32Status]]) : 0,
                                BytesSent = fieldIndex.ContainsKey (IISLogEntry.propBytesSent) ? Int32.Parse (columns[fieldIndex[IISLogEntry.propBytesSent]]) : 0,
                                BytesReceived = fieldIndex.ContainsKey (IISLogEntry.propBytesReceived) ? Int32.Parse (columns[fieldIndex[IISLogEntry.propBytesReceived]]) : 0,
                                TimeTaken = fieldIndex.ContainsKey (IISLogEntry.propTimeTaken) ? Int32.Parse (columns[fieldIndex[IISLogEntry.propTimeTaken]]) : 0,
                                ProtocolVersion = fieldIndex.ContainsKey (IISLogEntry.propProtocolVersion) ? columns[fieldIndex[IISLogEntry.propProtocolVersion]] : String.Empty,
                                Host = fieldIndex.ContainsKey (IISLogEntry.propHost) ? columns[fieldIndex[IISLogEntry.propHost]] : String.Empty,
                                UserAgent = fieldIndex.ContainsKey (IISLogEntry.propUserAgent) ? columns[fieldIndex[IISLogEntry.propUserAgent]] : String.Empty,
                                Cookie = fieldIndex.ContainsKey (IISLogEntry.propCookie) ? columns[fieldIndex[IISLogEntry.propCookie]] : String.Empty,
                                Referrer = fieldIndex.ContainsKey (IISLogEntry.propReferrer) ? columns[fieldIndex[IISLogEntry.propReferrer]] : String.Empty,
                                ProtocolSubstatus = fieldIndex.ContainsKey (IISLogEntry.propProtocolSubstatus) ? columns[fieldIndex[IISLogEntry.propProtocolSubstatus]] : String.Empty
                            };

                    #endregion

                            #region entry processing

                            var url = logEntry.URIStem.ToLower ();

                            #region HTTP status codes & IP
                            if ( httpStatus.ContainsKey (logEntry.HTTPStatus) )
                                httpStatus[logEntry.HTTPStatus]++;
                            else
                                httpStatus.Add (logEntry.HTTPStatus, 1);

                            if ( !uniqueIPs.Contains (logEntry.ClientIPAddress) )
                                uniqueIPs.Add (logEntry.ClientIPAddress);
                            #endregion

                            if ( nextTime == DateTime.MinValue )
                            {
                                firstEntry = logEntry.TimeStamp;
                                lastEntry = logEntry.TimeStamp;
                                nextTime = logEntry.TimeStamp.Date.
                                            AddHours (logEntry.TimeStamp.Hour).
                                            AddMinutes (logEntry.TimeStamp.Minute).
                                            AddMinutes (concurrencyWindow);
                            }

                            if ( logEntry.TimeStamp > nextTime )
                            {
                                if ( processingList.Any () )
                                {
                                    requests.Add (new ConcurrentRequest (concurrencyWindow)
                                    {
                                        TimeStamp = nextTime,
                                        Transactions = processingList.Count,
                                        AverageResponseTime = processingList.Average (p => p.TimeTaken),
                                        BytesSent = processingList.Sum (t => t.BytesSent)
                                    });
                                    processingList.Clear ();
                                }
                                else
                                {
                                    requests.Add (new ConcurrentRequest (concurrencyWindow)
                                    {
                                        TimeStamp = nextTime,
                                        Transactions = 0,
                                        AverageResponseTime = 0,
                                        BytesSent = 0
                                    });
                                }
                                nextTime = nextTime.AddMinutes (concurrencyWindow);
                            }

                            if ( lastEntry.Hour != logEntry.TimeStamp.Hour )
                            {
                                totalHours++;
                                AddHourlyPages (pageViewsHourly, hourlyPages, lastEntry);
                            }

                            if ( lastEntry.Date != logEntry.TimeStamp.Date )
                            {
                                totalDays++;
                                AddDailyPages (pageViewsDaily, dailyPages, lastEntry);
                            }

                            //add the current one to future processing, otherwise one in teh borderlien will be missing
                            if ( logEntry.HTTPStatus == 200 )
                            {
                                processingList.Add (logEntry);
                                ServedRequests++;

                                if ( pageViewsForPeriod.ContainsKey (url) )
                                    pageViewsForPeriod[url].Hit (logEntry.TimeTaken);
                                else
                                    pageViewsForPeriod.Add (url, new MethodInfo (logEntry.URIStem, logEntry.TimeTaken));

                                if ( lastEntry.Hour == logEntry.TimeStamp.Hour )
                                {
                                    if ( pageViewsHourly.ContainsKey (url) )
                                        pageViewsHourly[url].Hit (logEntry.TimeTaken);
                                    else
                                        pageViewsHourly.Add (url, new MethodInfo (logEntry.URIStem, logEntry.TimeTaken));
                                }

                                if ( lastEntry.Date == logEntry.TimeStamp.Date )
                                {
                                    if ( pageViewsDaily.ContainsKey (url) )
                                        pageViewsDaily[url].Hit (logEntry.TimeTaken);
                                    else
                                        pageViewsDaily.Add (url, new MethodInfo (logEntry.URIStem, logEntry.TimeTaken));
                                }

                                if ( hitsPerURLParams.Any () )
                                {
                                    var urlParam = hitsPerURLParams.Where (p => logEntry.URIQuery.Contains (p)).FirstOrDefault ();
                                    if ( urlParam != null && urlParam != String.Empty )
                                    {
                                        if ( urlParamHits.ContainsKey (url) )
                                            urlParamHits[url].Hit (logEntry.TimeTaken);
                                        else
                                            urlParamHits.Add (url, new MethodInfo (urlParam, logEntry.TimeTaken));
                                    }
                                }
                            }

                            lastEntry = logEntry.TimeStamp;
                        }
                    }

                    if ( processingList.Any () )
                    {
                        requests.Add (new ConcurrentRequest (concurrencyWindow)
                        {
                            TimeStamp = nextTime,
                            Transactions = processingList.Count,
                            AverageResponseTime = processingList.Average (p => p.TimeTaken),
                            BytesSent = processingList.Sum (t => t.BytesSent)
                        });
                        processingList.Clear ();
                    }
                    AddHourlyPages (pageViewsHourly, hourlyPages, lastEntry);
                    AddDailyPages (pageViewsDaily, dailyPages, lastEntry);

                            #endregion
                }


                catch ( Exception e )
                {
                    Console.WriteLine ("Error!! {0}:{1} - {2}", e.GetType ().Name, e.Message, e.StackTrace);
                    Debug.WriteLine ("Error!! {0}:{1}", e.GetType ().Name, e.Message);
                }
            }
            Console.WriteLine ("\nGenerating Statistics");


            #region resultprocessing
            IEnumerable<MethodInfo> topPages;
            IEnumerable<IGrouping<DateTime, MethodInfo>> hourlyHits = null;
            long peakHits;
            IEnumerable<IGrouping<DateTime, MethodInfo>> peakHourPages = null;

            try
            {
                excelApp = new Application ();
                excelApp.Visible = false;
                reportSpreadsheet = excelApp.Workbooks.Add ();
                excelApp.Calculation = XlCalculation.xlCalculationManual;
                reportSheet = reportSpreadsheet.ActiveSheet;
                #region Concurrent Users
                if ( requests.Any () )
                {
                    Console.WriteLine ("{0} Calculating Concurrent User Count", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"));

                    reportSheet.Name = "Concurrent Users";
                    reportSheet.Cells[reportRow, reportCol++] = "Timestamp";
                    reportSheet.Cells[reportRow, reportCol++] = "Requests";
                    reportSheet.Cells[reportRow, reportCol++] = "TPS";
                    reportSheet.Cells[reportRow, reportCol++] = "Average Response Time";
                    reportSheet.Cells[reportRow, reportCol++] = "Concurrent Users (based on Little's Law)";
                    reportSheet.Cells[reportRow, reportCol++] = "Bytes Sent";
                    reportSheet.Cells[reportRow, reportCol++] = "Network Speed (Mbps)";


                    foreach ( var p in requests )
                    {
                        reportCol = 1; reportRow++;
                        reportSheet.Cells[reportRow, reportCol++] = p.TimeStamp;
                        reportSheet.Cells[reportRow, reportCol++] = p.Transactions;
                        reportSheet.Cells[reportRow, reportCol++] = p.Tps;
                        reportSheet.Cells[reportRow, reportCol++] = p.AverageResponseTime;
                        reportSheet.Cells[reportRow, reportCol++] = p.ConcurrentUsers;
                        reportSheet.Cells[reportRow, reportCol++] = p.BytesSent;
                        reportSheet.Cells[reportRow, reportCol++] = p.NetworkSpeed;
                    }
                }
                #endregion

                reportSpreadsheet.Application.DisplayAlerts = false;
                reportSpreadsheet.SaveAs (exportFileName, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);

                #region Page visit Summary
                if ( pageViewsForPeriod.Any () )
                {
                    Console.WriteLine ("{0} Genrating Page visit Summary", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"));
                    reportSheet = reportSpreadsheet.Worksheets.Add (Type.Missing, reportSheet, 1);
                    reportSheet.Name = "Page visit Summary";


                    startRow = startCol = 1;

                    startRow = CollectionToTable (pageViewsForPeriod.Values, startRow, startCol, "Page visit Summary (for the period)");


                    reportSheet.Shapes.AddChart (XlChartType.xlLine).Select ();
                    excelApp.ActiveChart.SetSourceData (Source: reportSheet.get_Range ("A1:B" + startRow));

                    reportSheet.Shapes.AddChart (XlChartType.xlPie).Select ();
                    excelApp.ActiveChart.SetSourceData (Source: reportSheet.get_Range ("A1:B" + startRow));
                    excelApp.ActiveChart.ClearToMatchStyle ();
                    try
                    {
                        excelApp.ActiveChart.ChartStyle = 256;
                    }
                    catch ( Exception e )
                    { }

                    excelApp.ActiveChart.SetElement (Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                    excelApp.ActiveChart.ChartTitle.Text = "Page visit Summary (for the period) Most Visited Pages";

                    reportSheet.Shapes.AddChart (XlChartType.xlBarClustered).Select ();
                    excelApp.ActiveChart.SetSourceData (Source: reportSheet.get_Range ("A1:D" + startRow));
                    excelApp.ActiveChart.ClearToMatchStyle ();
                    try
                    {
                        excelApp.ActiveChart.ChartStyle = 222;
                    }
                    catch ( Exception e )
                    { }
                    excelApp.ActiveChart.SetElement (Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                    excelApp.ActiveChart.ChartTitle.Text = "Page visit Summary (for the period) Average Response Time";
                    SpreadCharts (reportSheet);

                }
                #endregion

                #region Daily Analysis
                if ( dailyPages.Any () )
                {
                    Console.WriteLine ("{0} Genrating Daily Statistics", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"));
                    reportSheet = reportSpreadsheet.Worksheets.Add (Type.Missing, reportSheet, 1);
                    reportSheet.Name = "Daily Analysis";

                    foreach ( var d in dailyPages.Select (p => p.Timestamp).Distinct () )
                    {
                        filteredEntries.UnionWith (dailyPages.Where (p => p.Timestamp == d.Date)
                                                                    .OrderByDescending (p => p.Hits).Take (topPagesPerDay));
                        //Debug.WriteLine("Date: {0} - {1}", date, MethodInfo.TotalHits(dailyPages.Where(p => p.Timestamp == d.Date)));
                    }

                    topPages = filteredEntries.Where (p => filteredEntries.Count (q => q.Url == p.Url) > totalDays / 2);
                    startRow = startCol = 1;
                    AddChartFromSeries (startRow, startCol, "Daily Top Pages - Visits Trend", topPages, p => p.Hits, d => d.ToString (DateTimeFormatInfo.CurrentInfo.ShortDatePattern));

                    startRow = reportRow + 10;
                    startCol = 1;
                    AddChartFromSeries (startRow, startCol, "Daily Top Pages - Response Time(Average) Trend", topPages, p => p.AvgResponseTime, d => d.ToString (DateTimeFormatInfo.CurrentInfo.ShortDatePattern));


                    startRow = reportRow + 10;
                    startCol = 1;
                    AddChartFromSeries (startRow, startCol, "Daily Top Pages - Response Time(90%tile) Trend", topPages, p => p.NinetiethPercentile, d => d.ToString (DateTimeFormatInfo.CurrentInfo.ShortDatePattern));

                    startRow = 1;
                    startCol = 30;
                    filteredEntries.Clear ();

                    //reportSheet.Cells[reportRow, reportCol] = "Date";
                    foreach ( var d in dailyPages.Select (p => p.Timestamp).Distinct () )
                    {
                        filteredEntries.UnionWith (dailyPages.Where (p => p.Timestamp == d.Date)
                                               .OrderByDescending (p => p.NinetiethPercentile).Take (topPagesPerDay));
                    }
                    topPages = filteredEntries.Where (p => filteredEntries.Count (q => q.Url == p.Url) > totalDays / 2);
                    AddChartFromSeries (startRow, startCol, "Daily Slow Pages - Response Time(90%tile) Trend", topPages, p => p.NinetiethPercentile, d => d.ToString (DateTimeFormatInfo.CurrentInfo.ShortDatePattern));


                    startRow = reportRow + 10;
                    startCol = 30;
                    filteredEntries.Clear ();

                    //reportSheet.Cells[reportRow, reportCol] = "Date";
                    foreach ( var d in dailyPages.Select (p => p.Timestamp).Distinct () )
                    {
                        filteredEntries.UnionWith (dailyPages.Where (p => p.Timestamp == d.Date)
                                               .OrderByDescending (p => p.AvgResponseTime).Take (topPagesPerDay));
                        //Debug.WriteLine("Date: {0} - {1}", date, MethodInfo.TotalHits(dailyPages.Where(p => p.Timestamp == d.Date)));
                    }
                    topPages = filteredEntries.Where (p => filteredEntries.Count (q => q.Url == p.Url) > totalDays / 2);
                    AddChartFromSeries (startRow, startCol, "Daily Slow Pages - Response Time(Average) Trend", topPages, p => p.AvgResponseTime, d => d.ToString (DateTimeFormatInfo.CurrentInfo.ShortDatePattern));

                    SpreadCharts (reportSheet);
                }

                #endregion

                #region Hourly analysis
                if ( hourlyPages.Any () )
                {
                    Console.WriteLine ("{0} Genrating Hourly Statistics", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"));
                    reportSheet = reportSpreadsheet.Worksheets.Add (Type.Missing, reportSheet, 1);
                    reportSheet.Name = "Hourly Analysis";

                    startRow = 1;
                    startCol = 1;
                    filteredEntries.Clear ();

                    foreach ( var d in hourlyPages.Select (p => p.Timestamp).Distinct () )
                    {
                        filteredEntries.UnionWith (hourlyPages.Where (p => p.Timestamp == d.Date.AddHours (d.Hour))
                                            .OrderByDescending (p => p.Hits).Take (topPagesPerDay));
                        //Debug.WriteLine("Date: {0} - {1}", date, MethodInfo.TotalHits(dailyPages.Where(p => p.Timestamp == d.Date)));
                    }
                    var totalHits = hourlyPages.Sum (p => p.Hits);
                    //filter out top pages which are there for 10% of time or 2% traffic
                    topPages = filteredEntries.Where (p => filteredEntries.Count (q => q.Url == p.Url) > totalHours / 10 || p.Hits > totalHits * 2 / 100);
                    startRow += AddChartFromSeries (startRow, startCol, "Hourly Top Pages Summary (By Hits)", topPages, p => p.Hits, d => d.ToString ());
                    excelApp.ActiveChart.Axes (XlAxisType.xlCategory).CategoryType = XlCategoryType.xlCategoryScale;

                    hourlyHits = hourlyPages.GroupBy (p => p.Timestamp, q => q);
                    peakHits = hourlyHits.Select (p => p.Sum (q => q.Hits)).OrderByDescending (p => p).Take (peakHoursCount).Min ();
                    peakHourPages = hourlyHits.Where (p => p.Sum (q => q.Hits) >= peakHits);

                    startRow += 10; startCol = 1;
                    startRow += AddChartFromSeries (startRow, startCol, "Peak Hour Top Pages Summary (By Hits)", peakHourPages.SelectMany (g => g.Where (p => p.Hits > peakHits * 2 / 100)), p => p.Hits, d => d.ToString ());
                    excelApp.ActiveChart.Axes (XlAxisType.xlCategory).CategoryType = XlCategoryType.xlCategoryScale;

                    CollectionToTable (peakHourPages.SelectMany (g => g), startRow + 10, 1, "Peak Hour Pages", true);

                    SpreadCharts (reportSheet);
                }
                #endregion

                #region URL Param Hits Summary
                if ( hitsPerURLParams.Any () )
                {
                    Console.WriteLine ("{0} Genrating URL parameter statistics", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"));

                    reportSheet = reportSpreadsheet.Worksheets.Add (Type.Missing, reportSheet, 1);
                    startRow = startCol = 1;
                    reportSheet.Name = "URL Parameters";
                    CollectionToTable (urlParamHits.Values, startRow, startCol, "URL Parameters Summary (for the period)");
                }
                #endregion

                #region Summary
                Console.WriteLine ("{0} Genrating Summary", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"));
                reportSheet = reportSpreadsheet.Worksheets.Add (reportSheet, Type.Missing, 1);
                reportRow = reportCol = 1;
                reportSheet.Name = "Summary";
                reportSheet.Cells[reportRow, 1] = "Running From";
                reportSheet.Cells[reportRow++, 2] = curerntPath;

                reportSheet.Cells[reportRow, 1] = "Commandline Argument";
                reportSheet.Cells[reportRow++, 2] = string.Join (";", cmdArgs.Select (x => x.Key + "=" + x.Value));

                reportSheet.Cells[reportRow, 1] = "Files Processed";
                reportSheet.Cells[reportRow++, 2] = fileCount;

                reportSheet.Cells[reportRow, 1] = "From";
                reportSheet.Cells[reportRow++, 2] = firstEntry;

                reportSheet.Cells[reportRow, 1] = "To";
                reportSheet.Cells[reportRow++, 2] = lastEntry;

                reportSheet.Cells[reportRow, 1] = "TotalHits";
                reportSheet.Cells[reportRow++, 2] = TotalHits;

                reportSheet.Cells[reportRow, 1] = "Average Transactions/Sec";
                reportSheet.Cells[reportRow++, 2] = requests.Average (p => p.Tps);

                if ( hourlyHits!=null )
                {
                    reportSheet.Cells[reportRow, 1] = "Average Transactions/Hour";
                    reportSheet.Cells[reportRow++, 2] = hourlyHits.Average (p => p.Sum (q => q.Hits));
                }

                if ( peakHourPages!=null )
                {
                    reportSheet.Cells[reportRow, 1] = "Peak Hour Transactions/Hour";
                    reportSheet.Cells[reportRow++, 2] = peakHourPages.Average (p => p.Sum (q => q.Hits));

                    reportSheet.Cells[reportRow, 1] = "Peak Hour Transactions/Sec";
                    reportSheet.Cells[reportRow++, 2] = peakHourPages.Average (p => p.Sum (q => q.Hits) / 3600);
                }

                reportSheet.Cells[reportRow, 1] = "UniqueIPs";
                reportSheet.Cells[reportRow++, 2] = uniqueIPs.Count;

                reportSheet.Cells[reportRow, 1] = "ServedRequests";
                reportSheet.Cells[reportRow++, 2] = ServedRequests;


                reportRow += 10;
                reportSheet.Cells[reportRow++, 1] = "Http Status code summary";

                reportSheet.Cells[reportRow, 1] = "HTTP Code";
                reportSheet.Cells[reportRow++, 2] = "Count";

                foreach ( var i in httpStatus )
                {
                    reportSheet.Cells[reportRow, reportCol++] = i.Key;
                    reportSheet.Cells[reportRow++, reportCol--] = ( i.Value );
                }
                #endregion

            }
            catch ( Exception e )
            {
                Console.WriteLine ("Error!! {0}:{1} - {2}", e.GetType ().Name, e.Message, e.StackTrace);
                Debug.WriteLine ("Error!! {0}:{1}", e.GetType ().Name, e.Message);
            }
            finally
            {
                if ( excelApp != null )
                {
                    excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
                    if ( reportSpreadsheet != null )
                    {
                        reportSpreadsheet.Save ();
                        reportSpreadsheet.Close ();
                        excelApp.Quit ();
                    }
                }
                File.Delete (tmpFile);
                stopWatch.Stop ();
                Console.WriteLine ("Done, Final time : {0}", stopWatch.Elapsed.ToString (@"hh\:mm\:ss"));
            }
            #endregion

            return 0;
        }

        private static Dictionary<string, int> ParseHeaderFields(string header)
        {
            var fields = header.Split (' ');
            Dictionary<string, int> fieldIndex = new Dictionary<string, int> ();

            for ( int i = 1; i < fields.Count (); i++ )
            {
                fieldIndex.Add (fields[i], i - 1);
            }
            return fieldIndex;
        }

        private static int CollectionToTable(IEnumerable<MethodInfo> collection, int reportRow, int reportCol, string Title, bool WithDateTime = false)
        {
            reportSheet.Cells[reportRow++, 1] = Title;
            if ( WithDateTime )
                reportSheet.Cells[reportRow, reportCol++] = "Date-Time";

            reportSheet.Cells[reportRow, reportCol++] = "Page Name";
            reportSheet.Cells[reportRow, reportCol++] = "Hits";
            reportSheet.Cells[reportRow, reportCol++] = "Min(sec)";
            reportSheet.Cells[reportRow, reportCol++] = "Avg(sec)";
            reportSheet.Cells[reportRow, reportCol++] = "Max(sec)";
            reportSheet.Cells[reportRow, reportCol++] = "90th-%(sec)";
            reportSheet.Cells[reportRow, reportCol++] = "Median(sec)";
            reportSheet.Cells[reportRow, reportCol++] = "Std. Deviation(sec)";
            foreach ( var i in collection )
            {
                reportRow++; reportCol = 1;
                if ( WithDateTime )
                    reportSheet.Cells[reportRow, reportCol++] = i.Timestamp;

                reportSheet.Cells[reportRow, reportCol++] = i.Url;
                reportSheet.Cells[reportRow, reportCol++] = i.Hits;
                reportSheet.Cells[reportRow, reportCol++] = i.MinResponseTime;
                reportSheet.Cells[reportRow, reportCol++] = i.AvgResponseTime;
                reportSheet.Cells[reportRow, reportCol++] = i.MaxResponseTime;
                reportSheet.Cells[reportRow, reportCol++] = i.NinetiethPercentile;
                reportSheet.Cells[reportRow, reportCol++] = i.Median;
                reportSheet.Cells[reportRow, reportCol++] = i.StandardDeviation;
            }

            return reportRow;
        }

        private static int AddChartFromSeries(int startRow, int startCol, string Title, IEnumerable<MethodInfo> TopPages, Func<MethodInfo, object> selector, Func<DateTime, string> dateFormat)
        {
            Dictionary<string, int> lookupRowCol = new Dictionary<string, int> ();
            int reportRow, reportCol;

            reportSheet.Cells[startRow, startCol] = Title;
            startRow += 2;
            reportRow = startRow;
            reportCol = startCol;

            if ( !TopPages.Any () ) return reportRow - startRow;

            foreach ( var page in TopPages )
            {
                var date = dateFormat (page.Timestamp);
                //Debug.WriteLine("Date: {0} Url{1} - daily {2} : top {3}", date, page.Url,
                //                MethodInfo.TotalHits(dailyPages.Where(p => p.Url == page.Url)),
                //                MethodInfo.TotalHits(topPages.Where(p => p.Url == page.Url)));

                if ( !lookupRowCol.ContainsKey (page.Url) )
                {
                    lookupRowCol.Add (page.Url, ++reportCol);
                    reportSheet.Cells[startRow, reportCol] = page.Url;
                }
                if ( !lookupRowCol.ContainsKey (date) )
                {
                    lookupRowCol.Add (date, ++reportRow);
                    reportSheet.Cells[reportRow, startCol] = date;
                }
                reportSheet.Cells[lookupRowCol[date], lookupRowCol[page.Url]] = selector.Invoke (page);
            }
            reportSheet.Shapes.AddChart (XlChartType.xlLine).Select ();
            excelApp.ActiveChart.SetSourceData (Source: reportSheet.Cells[startRow, startCol].CurrentRegion);
            excelApp.ActiveChart.SetElement (Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
            excelApp.ActiveChart.ChartTitle.Text = Title;
            return reportRow - startRow;
        }

        private static void AddDailyPages(Dictionary<string, MethodInfo> pageViewsDaily, HashSet<MethodInfo> dailyPages, DateTime lastEntry)
        {
            if ( pageViewsDaily.Any () )
            {
                foreach ( var page in pageViewsDaily.Values )
                {
                    page.Timestamp = lastEntry.Date;
                    dailyPages.Add (page);
                }
                pageViewsDaily.Clear ();
            }
        }

        private static void AddHourlyPages(Dictionary<string, MethodInfo> pageViewsHourly, HashSet<MethodInfo> hourlyPages, DateTime lastEntry)
        {
            if ( pageViewsHourly.Any () )
            {
                foreach ( var page in pageViewsHourly.Values )
                {
                    page.Timestamp = lastEntry.Date.AddHours (lastEntry.Hour);
                    hourlyPages.Add (page);
                }
                pageViewsHourly.Clear ();
            }
        }


        private static void SpreadCharts(Worksheet reportSheet)
        {
            Shape lastChart = null;
            foreach ( Shape chart in reportSheet.Shapes )
            {
                if ( lastChart != null )
                    chart.Left = lastChart.Left + lastChart.Width + 20;
                lastChart = chart;
            }
        }
    }
}
