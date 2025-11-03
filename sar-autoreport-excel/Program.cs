using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace sar_autoreport_excel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string paramDST = "";
            string paramSRC = "";
            string paramMON = "";
            string paramYRR = "";
            string paramWTH = "";
            string paramHGT = "";
            int intparamMonth = 1;
            int intparamYear = 2015;
            int intparamWidth = 1000;
            int intparamHeight = 300;
            bool switchDST = false;
            bool switchSRC = false;
            bool switchMON = false;
            bool switchYRR = false;
            bool switchWTH = false;
            bool switchHGT = false;
            bool switchNEW = false;
            bool switchYearReport = false;
            bool switchCSV = false;
            bool switchExcel = false;
            foreach (string arg in args)
            {
                try
                {

                    switch (arg.Substring(0, 4).ToUpper())
                    {
                        case "/DST":
                            switchDST = true;
                            paramDST = arg.Substring(5);
                            break;
                        case "/SRC":
                            switchSRC = true;
                            paramSRC = arg.Substring(5);
                            break;
                        case "/MON":
                            switchMON = true;
                            paramMON = arg.Substring(5);
                            intparamMonth = Convert.ToInt32(paramMON);
                            if (intparamMonth < 1 || intparamMonth > 12)
                            {
                                Console.WriteLine("The /MON Month Parameter must be between 1 and 12");
                                Console.WriteLine("");
                                showHelp();
                                return;
                            }
                            break;
                        case "/YRR":
                            switchYRR = true;
                            paramYRR = arg.Substring(5);
                            intparamYear = Convert.ToInt32(paramYRR);
                            if (intparamYear < 2010 || intparamYear > 2100)
                            {
                                Console.WriteLine("The /YRR Parameter must be between 2010 and 2099");
                                Console.WriteLine("");
                                showHelp();
                                return;
                            }
                            break;
                        case "/WTH":
                            switchWTH = true;
                            paramWTH = arg.Substring(5);
                            intparamWidth = Convert.ToInt32(paramWTH);
                            break;
                        case "/HGT":
                            switchHGT = true;
                            paramHGT = arg.Substring(5);
                            intparamHeight = Convert.ToInt32(paramHGT);
                            break;
                        case "/NEW":
                            switchNEW = true;
                            break;
                        case "/YRT":
                            switchYearReport = true;
                            break;
                        case "/CSV":
                            switchCSV = true;
                            break;
                        case "/XLS":
                            switchExcel = true;
                            break;
                        case "/HLP":
                            showHelp();
                            return;
                        default:
                            // do other stuff...
                            break;

                    }
                }
                catch
                {
                    Console.WriteLine("Error Reading Command Line Parameters");
                    Console.WriteLine("");
                    showHelp();
                    return;
                }

            }

            string sarcsv_directory;
            string destination;
            int graphwidth;
            int graphheight;
            bool generatecsv;
            bool generatexls;

            try
            {
                sarcsv_directory = Properties.Settings.Default.SarCsvSource;
                destination = Properties.Settings.Default.Destination;
                graphwidth = Properties.Settings.Default.GraphWidth;
                graphheight = Properties.Settings.Default.GraphHeight;
                generatecsv = Properties.Settings.Default.GenerateCSV;
                generatexls = Properties.Settings.Default.GenerateXLS;
            }
            catch
            {
                Console.WriteLine("There was an error reading the Applicaiton Config File");
                return;
            }

            // Get the default start date
            DateTime currentTime = DateTime.Now;
            DateTime MonthAgo = currentTime.AddMonths(-1);
            bool doyearlyreport = false;

            int startMonth = MonthAgo.Month;
            string startMonthstr = startMonth.ToString();

            int startYear = MonthAgo.Year;
            string startYearstr = startYear.ToString();

            if (switchDST == true)
            {
                Console.WriteLine("The /DST Switch was used " + paramDST);
            }
            if (switchSRC == true)
            {
                Console.WriteLine("The /SRC Switch was used");
            }

            if (switchMON == true)
            {
                startMonth = intparamMonth;
                startMonthstr = startMonth.ToString();
            }
            if (switchYRR == true)
            {
                startYear = intparamYear;
                startYearstr = startYear.ToString();
            }
            if (switchWTH == true)
            {
                Console.WriteLine("Setting the Excel Graphs Width to: " + paramWTH);
                graphwidth = intparamWidth;
            }
            if (switchHGT == true)
            {
                Console.WriteLine("Setting the Excel Graphs Height to: " + paramHGT);
                graphheight = intparamHeight;
            }
            if (switchYearReport == true)
            {
                Console.WriteLine("Generating Yearly Reports instead of Monthly Reports");
                doyearlyreport = true;
                startMonth = 1;
                startMonthstr = startMonth.ToString();
            }
            if (switchCSV == true)
            {
                generatecsv = true;
            }
            if (switchExcel == true)
            {
                generatexls = true;
            }

            // Calculate the End Dates
            string StartDateStr = startMonthstr + "/1/" + startYearstr;
            DateTime StartDate = Convert.ToDateTime(StartDateStr);
            DateTime tmpEndDate = StartDate.AddMonths(1);
            tmpEndDate = tmpEndDate.AddDays(-1);
            string tmpEndDateString = tmpEndDate.ToString();
            DateTime EndDate = Convert.ToDateTime(tmpEndDateString);
            if (doyearlyreport == true)
            {
                string YearEndDateStr = "12/31/" + startYearstr;
                DateTime YearEndDateTime = Convert.ToDateTime(YearEndDateStr);
                EndDate = YearEndDateTime;
            }

            // Create the Temp Directory and Check the Output Dir
            string tempPath = System.IO.Path.GetTempPath();
            tempPath = tempPath + "sar-reportautogen\\";
            if (Directory.Exists(tempPath))
            {
                Directory.Delete(tempPath, true);
                System.Threading.Thread.Sleep(2000);
            }
            Directory.CreateDirectory(tempPath);

            destination = destination + "\\";
            if (!(Directory.Exists(destination)))
            {
                try
                {
                    Directory.CreateDirectory(destination);
                }
                catch
                {
                    Console.WriteLine("Unable to Create Directory: " + destination);
                    return;
                }
            }
            string tmpDestination = Path.Combine(destination, "temp");
            try
            {
                Directory.CreateDirectory(tmpDestination);
            }
            catch
            {
                Console.WriteLine("Unable to write to the Destination: " + tmpDestination);
                return;
            }
            Directory.Delete(tmpDestination);

            // Get the Server List
            sarcsv_directory = sarcsv_directory + "\\";
            string basefolder = sarcsv_directory;

            List<string> dirs = new List<string>(Directory.EnumerateDirectories(sarcsv_directory));
            List<string> computers = new List<string>();

            foreach (var dir in dirs)
            {
                string test_dir = Path.Combine(dir, startYearstr);

                if (Directory.Exists(test_dir))
                {
                    string cpudir = dir.Substring(dir.LastIndexOf("\\") + 1);
                    cpudir = cpudir.ToUpper();
                    computers.Add(cpudir);
                }
            }

            int numcpus = computers.Count();
            if (numcpus < 1)
            {
                Console.WriteLine("There are no Performance Stats for any system located at " + sarcsv_directory);
                return;
            }
            Console.WriteLine("Start Generating Reports for " + numcpus + " Systems");

            foreach (var computer in computers)
            {
                string servername = computer;
                string strStartDate = StartDate.ToString("yyyyMMdd");
                string strEndDate = EndDate.ToString("yyyyMMdd");

                Console.WriteLine("Starting to Generate a Report for " + servername);

                // Initialize Output File Paths
                string outputFile;
                outputFile = servername + "-" + strStartDate + "-" + strEndDate;
                string csvoutputFile = outputFile + ".csv";
                string exceloutputFile = outputFile + ".xlsx";

                string finaldest = Path.Combine(destination, servername);
                if (!Directory.Exists(finaldest))
                {
                    Directory.CreateDirectory(finaldest);
                }
                finaldest = Path.Combine(finaldest, startYearstr);
                if (!Directory.Exists(finaldest))
                {
                    Directory.CreateDirectory(finaldest);
                }
                string destCSVFile = Path.Combine(finaldest, csvoutputFile);
                string destXLSFile = Path.Combine(finaldest, exceloutputFile);
                outputFile = Path.Combine(tempPath, outputFile);
                csvoutputFile = Path.Combine(tempPath, csvoutputFile);
                exceloutputFile = Path.Combine(tempPath, exceloutputFile);

                StreamWriter swOutputCSV = new StreamWriter(csvoutputFile);
                swOutputCSV.WriteLine("servername,timestamp,cpu_system,cpu_user,cpu_steal,cpu_iowait,cpu_nice,cpu_idle,mem_total,mem_active,mem_used,mem_buffers,mem_cached,mem_hugused,mem_hugfree,mem_free,%active,%buffers,%cached,%hugeused,%hugefree,%free,%swapused,swap_free,swap_used,iotps,iortps,iowtps,readmbs,writembs,net_rxpck,net_txpck,net_rxkbs,net_txkbs,rxmbs,txmbs");

                // Confirm if Ex
                if ((switchNEW == true) || ((switchNEW == false) && (!File.Exists(destXLSFile))))
                {
                    for (DateTime date = StartDate; date.Date <= EndDate.Date; date = date.AddDays(1))
                    {

                            DateTime processtime = date;

                            int processyear = processtime.Year;
                            string processyearstr = processyear.ToString();
                            int processmonth = processtime.Month;
                            string processmonthstr = processmonth.ToString();
                            if (processmonth < 10)
                            {
                                processmonthstr = "0" + processmonthstr;
                            }
                            int processday = processtime.Day;
                            string processdaystr = processday.ToString();
                            if (processday < 10)
                            {
                                processdaystr = "0" + processdaystr;
                            }

                            string csvPath = Path.Combine(basefolder, servername, processyearstr, processmonthstr, processdaystr);

                            try
                            {

                            string cpuStats = csvPath + "\\sar_cpu.csv";
                            string memStats = csvPath + "\\sar_mem.csv";
                            string hugepageStats = csvPath + "\\sar_hugepage.csv";
                            string iotransStats = csvPath + "\\sar_iotrans.csv";
                            string networkStats = csvPath + "\\sar_net.csv";
                            string swapStats = csvPath + "\\sar_swap.csv";
                        


                            System.Data.DataTable tmpcpuDatatable = new System.Data.DataTable();
                            try { tmpcpuDatatable = ReadCsvToDataTable(cpuStats); } catch { Console.WriteLine("Error reading " + cpuStats); }
                            System.Data.DataTable tmpmemDatatable = new System.Data.DataTable();
                            try { tmpmemDatatable = ReadCsvToDataTable(memStats); } catch { Console.WriteLine("Error reading " + memStats); }
                            System.Data.DataTable tmphugepageDatatable = new System.Data.DataTable();
                            try { tmphugepageDatatable = ReadCsvToDataTable(hugepageStats); } catch { Console.WriteLine("Error reading " + hugepageStats); }
                            System.Data.DataTable tmpiotransDatatable = new System.Data.DataTable();
                            try { tmpiotransDatatable = ReadCsvToDataTable(iotransStats); } catch { Console.WriteLine("Error reading " + iotransStats); }
                            System.Data.DataTable tmpnetworkDatatable = new System.Data.DataTable();
                            try { tmpnetworkDatatable = ReadCsvToDataTable(networkStats); } catch { Console.WriteLine("Error reading " + networkStats); }
                            System.Data.DataTable tmpswapDatatable = new System.Data.DataTable();
                            try { tmpswapDatatable = ReadCsvToDataTable(swapStats); } catch { Console.WriteLine("Error reading " + swapStats); }


                            // Create New Net Datatable to add the Sums
                            //
                            // Sum Up NetworkStats since it is by Interface
                            string[] columnsToSum = { "rxpck/s", "txpck/s", "rxkB/s", "txkB/s" };
                            System.Data.DataTable tmpnetsumDatatable = new System.Data.DataTable();
                            tmpnetsumDatatable.Columns.Add("timestamp", typeof(string));
                            tmpnetsumDatatable.Columns.Add("rxpck/s", typeof(string));
                            tmpnetsumDatatable.Columns.Add("txpck/s", typeof(string));
                            tmpnetsumDatatable.Columns.Add("rxkB/s", typeof(string));
                            tmpnetsumDatatable.Columns.Add("txkB/s", typeof(string));

                            var grouped = tmpnetworkDatatable.AsEnumerable()
                                .GroupBy(r => r.Field<string>("timestamp"))
                                .Select(g =>
                                {
                                    var row = tmpnetsumDatatable.NewRow();
                                    row["timestamp"] = g.Key;

                                    foreach (string col in columnsToSum)
                                    {
                                        row[col] = g.Sum(r => Convert.ToDecimal(r[col]));
                                    }
                                    return row;

                                });

                            foreach (var r in grouped)
                            {
                                tmpnetsumDatatable.Rows.Add(r);
                            }
                            // End Summing up NetStats

                            System.Data.DataTable cpuDatatable = new System.Data.DataTable();

                            int cpurows = tmpcpuDatatable.Rows.Count;
                            int memrows = tmpmemDatatable.Rows.Count;
                            int hugerows = tmphugepageDatatable.Rows.Count;
                            int iorows = tmpiotransDatatable.Rows.Count;
                            int netrows = tmpnetsumDatatable.Rows.Count;
                            int swaprows = tmpswapDatatable.Rows.Count;

                            if ((cpurows == memrows) && (cpurows == hugerows) && (cpurows == iorows) && (cpurows == netrows) && (cpurows == swaprows))
                            {
                                // Console.WriteLine("Good, All CSV Files at " + csvPath + " have the same number of rows after transformation");

                            }
                            else
                            {
                                Console.WriteLine("Warning: not all CSV files at " + csvPath + " have the same number of rows");
                                Console.WriteLine("CPU: " + cpurows);
                                Console.WriteLine("Mem: " + memrows);
                                Console.WriteLine("Huge: " + hugerows);
                                Console.WriteLine("IO: " + iorows);
                                Console.WriteLine("Net: " + netrows);
                                Console.WriteLine("Swap: " + swaprows);

                            }


                            // Down and dirty way to pull data based on index ;)

                            for (int i = 0; i < cpurows; i++)
                            {
                                // CPU Values from CSV
                                string str_servername = tmpcpuDatatable.Rows[i].Field<string>("hostname");
                                string str_timestamp = tmpcpuDatatable.Rows[i].Field<string>("timestamp");
                                string str_cpuuser = tmpcpuDatatable.Rows[i].Field<string>("%user");
                                string str_cpunice = tmpcpuDatatable.Rows[i].Field<string>("%nice");
                                string str_cpusystem = tmpcpuDatatable.Rows[i].Field<string>("%system");
                                string str_cpuiowait = tmpcpuDatatable.Rows[i].Field<string>("%iowait");
                                string str_cpusteal = tmpcpuDatatable.Rows[i].Field<string>("%steal");
                                string str_cpuidle = tmpcpuDatatable.Rows[i].Field<string>("%idle");
                                // Memory Values from CSV
                                string str_kbmemfree = tmpmemDatatable.Rows[i].Field<string>("kbmemfree");
                                string str_kbmemused = tmpmemDatatable.Rows[i].Field<string>("kbmemused");
                                string str_kbbuffers = tmpmemDatatable.Rows[i].Field<string>("kbbuffers");
                                string str_kbcache = tmpmemDatatable.Rows[i].Field<string>("kbcached");
                                string str_kbhugused = tmphugepageDatatable.Rows[i].Field<string>("kbhugused");
                                string str_kbhugfree = tmphugepageDatatable.Rows[i].Field<string>("kbhugfree");
                                // IO Trans Values from CSV
                                string str_iotps = tmpiotransDatatable.Rows[i].Field<string>("tps");
                                string str_iortps = tmpiotransDatatable.Rows[i].Field<string>("rtps");
                                string str_iowtps = tmpiotransDatatable.Rows[i].Field<string>("wtps");
                                string str_iobread = tmpiotransDatatable.Rows[i].Field<string>("bread/s");
                                string str_iobwrtn = tmpiotransDatatable.Rows[i].Field<string>("bwrtn/s");
                                // Network Summed Values from Converted from CSV 
                                string str_rxpck = tmpnetsumDatatable.Rows[i].Field<string>("rxpck/s");
                                string str_txpck = tmpnetsumDatatable.Rows[i].Field<string>("txpck/s");
                                string str_rxkbs = tmpnetsumDatatable.Rows[i].Field<string>("rxkB/s");
                                string str_txkbs = tmpnetsumDatatable.Rows[i].Field<string>("txkB/s");
                                // Swap Values from CSV
                                string str_perswapused = tmpswapDatatable.Rows[i].Field<string>("%swpused");
                                string str_kbswpfree = tmpswapDatatable.Rows[i].Field<string>("kbswpfree");
                                string str_kbswpused = tmpswapDatatable.Rows[i].Field<string>("kbswpused");

                                // CPU Conversion
                                decimal cpuuser = Convert.ToDecimal(str_cpuuser);
                                decimal cpunice = Convert.ToDecimal(str_cpunice);
                                decimal cpusystem = Convert.ToDecimal(str_cpusystem);
                                decimal cpuiowait = Convert.ToDecimal(str_cpuiowait);
                                decimal cpusteal = Convert.ToDecimal(str_cpusteal);
                                decimal cpuidle = Convert.ToDecimal(str_cpuidle);
                                decimal cputotalpercent = cpuuser + cpunice + cpusystem + cpuiowait + cpusteal + cpuidle;

                                // Memory Conversion
                                decimal kbmemfree = Convert.ToDecimal(str_kbmemfree);
                                decimal kbmemused = Convert.ToDecimal(str_kbmemused);
                                decimal kbbuffers = Convert.ToDecimal(str_kbbuffers);
                                decimal kbcached = Convert.ToDecimal(str_kbcache);
                                decimal kbhugused = Convert.ToDecimal(str_kbhugused);
                                decimal kbhugfree = Convert.ToDecimal(str_kbhugfree);

                                decimal totalmem = kbmemfree + kbmemused;
                                decimal activemem = kbmemused - kbbuffers - kbcached - kbhugused - kbhugfree;
                                // Adjust calculation for SAR v12+ since those calculations were changed, not 100% accurate, but close :(
                                if (activemem < 0)
                                {
                                    totalmem = kbmemfree + kbmemused + kbbuffers + kbcached;
                                    activemem = kbmemused - kbhugused - kbhugfree;
                                }
                                decimal percent_activemem = Math.Round(((activemem / totalmem) * 100), 2);
                                decimal percent_buffers = Math.Round(((kbbuffers / totalmem) * 100), 2);
                                decimal percent_cached = Math.Round(((kbcached / totalmem) * 100), 2);
                                decimal percent_hugused = Math.Round(((kbhugused / totalmem) * 100), 2);
                                decimal percent_hugfree = Math.Round(((kbhugfree / totalmem) * 100), 2);
                                decimal percent_freemem = Math.Round(((kbmemfree / totalmem) * 100), 2);

                                // Blocks Read Per Second to MB/sec
                                // Example Calculation is (iobread/sec x 512) / (1024 x 1024) or iobread/sec x 0.00048828125

                                decimal iobread = Convert.ToDecimal(str_iobread);
                                decimal iobwrtn = Convert.ToDecimal(str_iobwrtn);
                                decimal ioreadmbs = iobread * 512m / (1024m * 1024m);
                                decimal readmbs = Math.Round(ioreadmbs, 2);
                                decimal iowritembs = iobwrtn * 512m / (1024m * 1024m);
                                decimal writembs = Math.Round(iowritembs, 2);

                                decimal rxkbs = Convert.ToDecimal(str_rxkbs);
                                decimal txkbs = Convert.ToDecimal(str_txkbs);
                                decimal rxmbs = Math.Round((rxkbs / 1024 * 8), 2);
                                decimal txmbs = Math.Round((txkbs / 1024 * 8), 2);

                                swOutputCSV.WriteLine(str_servername + "," + str_timestamp + "," + cpusystem + "," + cpuuser + "," + cpusteal + "," + cpuiowait + "," + cpunice + "," + cpuidle + "," + totalmem + "," + activemem + "," + kbmemused + "," + kbbuffers + "," + kbcached + "," + kbhugused + "," + kbhugfree + "," + kbmemfree + "," + percent_activemem + "," + percent_buffers + "," + percent_cached + "," + percent_hugused + "," + percent_hugfree + "," + percent_freemem + "," + str_perswapused + "," + str_kbswpfree + "," + str_kbswpused + "," + str_iotps + "," + str_iortps + "," + str_iowtps + "," + readmbs + "," + writembs + "," + str_rxpck + "," + str_txpck + "," + str_rxkbs + "," + str_txkbs + "," + rxmbs + "," + txmbs);




                            }

                        } catch { Console.WriteLine("Error processing data from " + csvPath);  }

                    }

                    swOutputCSV.Close();
                    Console.WriteLine("Wrote the file: " + csvoutputFile);

                    int csvlineCount = File.ReadAllLines(csvoutputFile).Length;

                    if ((generatexls == true) && (csvlineCount > 10 ))
                    {

                        Excel.Application xlApp = new Excel.Application();

                        if (xlApp == null)
                        {
                            Console.WriteLine("Excel is not properly installed!");
                            return;
                        }

                        xlApp.Visible = false;
                        Excel.Workbook xlWorkBook;
                        Excel.Worksheet xlWorkSheet;
                        object misValue = System.Reflection.Missing.Value;


                        Microsoft.Office.Interop.Excel.WorksheetFunction wsf = xlApp.WorksheetFunction;

                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Name = "PerfData";

                        Excel.Range CellRef = xlWorkSheet.Range["A1"];

                        string TxtConnector = "TEXT;" + csvoutputFile;
                        Excel.QueryTable qtables = xlWorkSheet.QueryTables.Add(TxtConnector, CellRef);
                        qtables.TextFileCommaDelimiter = true;
                        qtables.TextFileParseType = Excel.XlTextParsingType.xlDelimited;
                        qtables.Refresh();
                        qtables.Delete();

                        Excel.Range CellRef2 = xlWorkSheet.Range["A1"];

                        xlWorkSheet.UsedRange.EntireColumn.AutoFit();
                        CellRef2.EntireRow.Font.Bold = true;
                        // xlWorkSheet.Application.ActiveWindow.SplitColumn = 1;
                        xlWorkSheet.Application.ActiveWindow.SplitRow = 1;
                        xlWorkSheet.Application.ActiveWindow.FreezePanes = true;

                        // Set timestamp to an actual timestamp
                        Range datetimecol = xlWorkSheet.get_Range("B:B");
                        datetimecol.NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM";

                        double numgraphs = 0;


                        // Graphs go here
                        /* Create Graphs WorkSheet */

                        numgraphs = 0;
                        Excel.Worksheet GraphsWorkSheet;
                        GraphsWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                        GraphsWorkSheet.Name = "PerfGraphs";

                        Excel.Range CPUStatsRange;
                        Excel.ChartObjects CPUStatsChartObjects = (Excel.ChartObjects)GraphsWorkSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject CPUStatsChartObject = (Excel.ChartObject)CPUStatsChartObjects.Add(0, (graphheight * numgraphs), graphwidth, graphheight);
                        Excel.Chart CPUStatsChart = CPUStatsChartObject.Chart;

                        CPUStatsRange = xlWorkSheet.get_Range("B:B,C:C,D:D,E:E,F:F,G:G,H:H");
                        CPUStatsChart.SetSourceData(CPUStatsRange);
                        CPUStatsChart.ChartType = Excel.XlChartType.xlAreaStacked100;

                        Excel.Axis CPUDateAxis = CPUStatsChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        CPUDateAxis.TickLabels.NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@";


                        CPUStatsChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                        Excel.ChartTitle CPUStatsChartTitle = CPUStatsChart.ChartTitle;
                        CPUStatsChartTitle.Text = "% CPU Usage - " + servername + " - " + startYearstr + "-" + startMonthstr;

                        Excel.Series CPUSystemSeries = (Excel.Series)CPUStatsChart.SeriesCollection(1);
                        CPUSystemSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent2;
                        CPUSystemSeries.Format.Fill.ForeColor.Brightness = -0.5f;

                        Excel.Series CPUUserSeries = (Excel.Series)CPUStatsChart.SeriesCollection(2);
                        CPUUserSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent2;
                        CPUUserSeries.Format.Fill.ForeColor.Brightness = 0.1f;

                        Excel.Series CPUStealSeries = (Excel.Series)CPUStatsChart.SeriesCollection(3);
                        CPUStealSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorText2;
                        CPUStealSeries.Format.Fill.ForeColor.Brightness = 0.25f;

                        Excel.Series CPUIowaitSeries = (Excel.Series)CPUStatsChart.SeriesCollection(4);
                        CPUIowaitSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorText2;
                        CPUIowaitSeries.Format.Fill.ForeColor.Brightness = 0.6f;

                        Excel.Series CPUNiceSeries = (Excel.Series)CPUStatsChart.SeriesCollection(5);
                        CPUNiceSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent6;
                        CPUNiceSeries.Format.Fill.ForeColor.Brightness = 0.1f;

                        Excel.Series CPUIdleSeries = (Excel.Series)CPUStatsChart.SeriesCollection(6);
                        CPUIdleSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent6;
                        CPUIdleSeries.Format.Fill.ForeColor.Brightness = 0.5f;

                        /* Memory Usage Chart */
                        numgraphs++;
                        Excel.Range MemStatsRange;
                        Excel.ChartObjects MemStatsChartObjects = (Excel.ChartObjects)GraphsWorkSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject MemStatsChartObject = (Excel.ChartObject)MemStatsChartObjects.Add(0, (graphheight * numgraphs), graphwidth, graphheight);
                        Excel.Chart MemStatsChart = MemStatsChartObject.Chart;

                        MemStatsRange = xlWorkSheet.get_Range("B:B,Q:Q,R:R,S:S,T:T,U:U,V:V");
                        MemStatsChart.SetSourceData(MemStatsRange);
                        MemStatsChart.ChartType = Excel.XlChartType.xlAreaStacked100;

                        Excel.Axis MemDateAxis = MemStatsChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        MemDateAxis.TickLabels.NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@";

                        Excel.SeriesCollection MemSeries = (Excel.SeriesCollection)MemStatsChart.SeriesCollection();

                        MemStatsChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                        Excel.ChartTitle MemStatsChartTitle = MemStatsChart.ChartTitle;
                        MemStatsChartTitle.Text = "% Mem Usage - " + servername + " - " + startYearstr + "-" + startMonthstr;

                        Excel.Series MemActiveSeries = (Excel.Series)MemStatsChart.SeriesCollection(1);
                        MemActiveSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent2;
                        MemActiveSeries.Format.Fill.ForeColor.Brightness = -0.5f;
                        
                        Excel.Series MemBuffersSeries = (Excel.Series)MemStatsChart.SeriesCollection(2);
                        MemBuffersSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorText2;
                        MemBuffersSeries.Format.Fill.ForeColor.Brightness = 0.2f;

                        Excel.Series MemCachedSeries = (Excel.Series)MemStatsChart.SeriesCollection(3);
                        MemCachedSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent6;
                        MemCachedSeries.Format.Fill.ForeColor.Brightness = -0.1f;

                        Excel.Series MemHugeUsedSeries = (Excel.Series)MemStatsChart.SeriesCollection(4);
                        MemHugeUsedSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent2;
                        MemHugeUsedSeries.Format.Fill.ForeColor.Brightness = -0.1f;
                        
                        Excel.Series MemHugeFreeSeries = (Excel.Series)MemStatsChart.SeriesCollection(5);
                        MemHugeFreeSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
                        MemHugeFreeSeries.Format.Fill.ForeColor.Brightness = 0.5f;
                        
                        Excel.Series MemFreeSeries = (Excel.Series)MemStatsChart.SeriesCollection(6);
                        MemFreeSeries.Format.Fill.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent6;
                        MemFreeSeries.Format.Fill.ForeColor.Brightness = 0.5f;

                        // Adjust Series order to move HugeUsed/HugeFree at the ends
                        MemHugeUsedSeries.PlotOrder = 1;
                        MemHugeFreeSeries.PlotOrder = MemSeries.Count;


                        // % Swap Usage Chart
                        numgraphs++;

                        Excel.Range SwapUsageRange;
                        Excel.ChartObjects SwapUsageChartObjects = (Excel.ChartObjects)GraphsWorkSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject SwapUsageChartObject = (Excel.ChartObject)MemStatsChartObjects.Add(0, (graphheight * numgraphs), graphwidth, graphheight);
                        Excel.Chart SwapUsageChart = SwapUsageChartObject.Chart;

                        SwapUsageRange = xlWorkSheet.get_Range("B:B,W:W");
                        SwapUsageChart.SetSourceData(SwapUsageRange, misValue);
                        SwapUsageChart.ChartType = Excel.XlChartType.xlLine;

                        Excel.Axis SwapUsageAxis = SwapUsageChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        SwapUsageAxis.TickLabels.NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@";

                        Excel.Axis SwapUsageValueAxis = (Excel.Axis)SwapUsageChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                        SwapUsageValueAxis.MaximumScale = 100;

                        SwapUsageChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                        Excel.ChartTitle SwapUsageChartTitle = SwapUsageChart.ChartTitle;
                        SwapUsageChartTitle.Text = "% Swap Usage - " + servername + " - " + startYearstr + "-" + startMonthstr;


                        Excel.Series SwapPercentSeries = (Excel.Series)SwapUsageChart.SeriesCollection(1);


                        SwapPercentSeries.Format.Line.Weight = 1.5f;


                        // Total IOPs all Drives Chart
                        numgraphs++;

                        Excel.Range IopsRange;
                        Excel.ChartObjects IopsChartObjects = (Excel.ChartObjects)GraphsWorkSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject IopsChartObject = (Excel.ChartObject)MemStatsChartObjects.Add(0, (graphheight * numgraphs), graphwidth, graphheight);
                        Excel.Chart IopsChart = IopsChartObject.Chart;

                        IopsRange = xlWorkSheet.get_Range("B:B,Z:Z,AA:AA,AB:AB");
                        IopsChart.SetSourceData(IopsRange, misValue);
                        IopsChart.ChartType = Excel.XlChartType.xlLine;

                        Excel.Axis IopsAxis = IopsChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        IopsAxis.TickLabels.NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@";

                        Excel.Axis IopsValueAxis = (Excel.Axis)IopsChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                        //IopsValueAxis.MaximumScale = 100;

                        IopsChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                        Excel.ChartTitle IopsChartTitle = IopsChart.ChartTitle;
                        IopsChartTitle.Text = "Total IOPs (All Disks) - " + servername + " - " + startYearstr + "-" + startMonthstr;

                        Excel.Series IopsTotSeries = (Excel.Series)IopsChart.SeriesCollection(1);
                        Excel.Series IopsReadSeries = (Excel.Series)IopsChart.SeriesCollection(2);
                        Excel.Series IopsWriteSeries = (Excel.Series)IopsChart.SeriesCollection(3);

                        IopsTotSeries.Format.Line.Weight = 1;
                        IopsReadSeries.Format.Line.Weight = 1;
                        IopsWriteSeries.Format.Line.Weight = 1;

                        // Total Disk Transfer all Drives Chart
                        numgraphs++;

                        Excel.Range DtransRange;
                        Excel.ChartObjects DtransChartObjects = (Excel.ChartObjects)GraphsWorkSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject DtransChartObject = (Excel.ChartObject)MemStatsChartObjects.Add(0, (graphheight * numgraphs), graphwidth, graphheight);
                        Excel.Chart DtransChart = DtransChartObject.Chart;

                        DtransRange = xlWorkSheet.get_Range("B:B,AC:AC,AD:AD");
                        DtransChart.SetSourceData(DtransRange, misValue);
                        DtransChart.ChartType = Excel.XlChartType.xlLine;

                        Excel.Axis DtransAxis = DtransChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        DtransAxis.TickLabels.NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@";

                        Excel.Axis DtransValueAxis = (Excel.Axis)DtransChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                        //DtransValueAxis.MaximumScale = 100;

                        DtransChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                        Excel.ChartTitle DtransChartTitle = DtransChart.ChartTitle;
                        DtransChartTitle.Text = "Disk Bandwidth (All Disks) - " + servername + " - " + startYearstr + "-" + startMonthstr;

                        Excel.Series DtransReadSeries = (Excel.Series)DtransChart.SeriesCollection(1);
                        Excel.Series DtransWriteSeries = (Excel.Series)DtransChart.SeriesCollection(2);
   

                        DtransReadSeries.Format.Line.Weight = 1;
                        DtransWriteSeries.Format.Line.Weight = 1;


                        // Total Disk Transfer all Drives Chart
                        numgraphs++;

                        Excel.Range NetBandwithRange;
                        Excel.ChartObjects NetBandwithChartObjects = (Excel.ChartObjects)GraphsWorkSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject NetBandwithChartObject = (Excel.ChartObject)MemStatsChartObjects.Add(0, (graphheight * numgraphs), graphwidth, graphheight);
                        Excel.Chart NetBandwithChart = NetBandwithChartObject.Chart;

                        NetBandwithRange = xlWorkSheet.get_Range("B:B,AI:AI,AJ:AJ");
                        NetBandwithChart.SetSourceData(NetBandwithRange, misValue);
                        NetBandwithChart.ChartType = Excel.XlChartType.xlLine;

                        Excel.Axis NetBandwithAxis = NetBandwithChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        NetBandwithAxis.TickLabels.NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@";

                        Excel.Axis NetBandwithValueAxis = (Excel.Axis)NetBandwithChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                        //NetBandwithValueAxis.MaximumScale = 100;

                        NetBandwithChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart);
                        Excel.ChartTitle NetBandwithChartTitle = NetBandwithChart.ChartTitle;
                        NetBandwithChartTitle.Text = "Network Bandwidth (All NICs) - " + servername + " - " + startYearstr + "-" + startMonthstr;

                        Excel.Series NetBandwithRxSeries = (Excel.Series)NetBandwithChart.SeriesCollection(1);
                        Excel.Series NetBandwithTxSeries = (Excel.Series)NetBandwithChart.SeriesCollection(2);


                        NetBandwithRxSeries.Format.Line.Weight = 1;
                        NetBandwithTxSeries.Format.Line.Weight = 1;


                        // Save file and release app

                        xlWorkBook.SaveAs(exceloutputFile, Excel.XlFileFormat.xlOpenXMLWorkbook);

                        System.Threading.Thread.Sleep(5000);

                        xlWorkBook.Close();

                        System.Threading.Thread.Sleep(1000);
                        xlApp.Quit();


                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkBook);
                        releaseObject(xlApp);

                        try
                        {
                            File.Copy(exceloutputFile, destXLSFile, true);
                            System.Threading.Thread.Sleep(2000);
                            if (File.Exists(destXLSFile))
                            {
                                File.Delete(exceloutputFile);
                            }
                        }
                        catch { Console.WriteLine("Unable to save file to " + destXLSFile); }

                    } else { Console.WriteLine("CSV File has no Perf Data - " + csvoutputFile + " - Skipping"); }

                    if (generatecsv == true)
                    {
                        try
                        {
                            File.Copy(csvoutputFile, destCSVFile, true);
                            System.Threading.Thread.Sleep(2000);
                            if (File.Exists(destXLSFile))
                            {
                                File.Delete(csvoutputFile);
                            }
                        }
                        catch { Console.WriteLine("Unable to save file to " + destXLSFile); }
                    }


                }
                else { Console.WriteLine("Excel File Exists: " + destXLSFile + " - skipping"); }

            }
            Console.WriteLine("Finished Generating Reports");


        }

        static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }



        static void showHelp()
        {
            Console.WriteLine("Example: sar-autoreport-excel /D:\"path to report destination\" /M:5 /W:2000 /CSV");
            Console.WriteLine("");
            Console.WriteLine("Options:");
            Console.WriteLine("/DST:\"...\" = Destination Location to write reports to");
            Console.WriteLine("/SRC:\"...\" = Performance Data Source Location");
            Console.WriteLine("/MON:XX = Month to generate report for, defaults to last month");
            Console.WriteLine("/YRR:XXXX = Year to generate report for, defaults to year of last month");
            Console.WriteLine("/WTH:XXXX = Width of Graph in pixels, defaults to 1000");
            Console.WriteLine("/HGT:XXXX = Height of Graph in pixels, defaults to 300");
            Console.WriteLine("/YRT - generate a Year report instead of a Monthly Report");
            Console.WriteLine("/HRY - force collection of Hourly stats for Year Reports");
            Console.WriteLine("/CSV - include a CSV file");
            Console.WriteLine("/XLS - include a XLS file");
            Console.WriteLine("/NEW - Overwrite Existing XLS file");
            Console.WriteLine("");
            Console.WriteLine("Note setting these options will disable what is set in the apps config file");
            Console.WriteLine("");


        }

        static System.Data.DataTable ReadCsvToDataTable(string filePath)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            bool firstLine = true;

            using (var reader = new StreamReader(filePath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line == null) continue;

                    var values = line.Split(';').Select(v => v.TrimStart('#')).ToArray();

                    if (firstLine)
                    {
                        // Create columns based on the header row
                        foreach (var header in values)
                        {
                            dt.Columns.Add(header.Trim());
                        }
                        firstLine = false;
                    }
                    else
                    {
                        // Add data rows
                        dt.Rows.Add(values);
                    }
                }
            }

            return dt;
        }



    }





}
