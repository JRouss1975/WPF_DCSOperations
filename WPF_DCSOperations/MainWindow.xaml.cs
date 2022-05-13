using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using System.Xml.Serialization;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using MathNet.Numerics;
using System.Collections.ObjectModel;
using MNet = MathNet.Numerics;
using ox = OxyPlot.Wpf;

namespace WPF_DCSOperations
{
    public partial class MainWindow : System.Windows.Window
    {
        Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
        public ObservableCollection<Info> lstInfo = new ObservableCollection<Info>();
        Thread thread;

        public MainWindow()
        {
            InitializeComponent();

            main.DataContext = this;
            //dgv1.DataContext = this;
        }

        private void btnOpenDCS_Click(object sender, RoutedEventArgs e)
        {

            // Set filter for file extension and default file extension    
            dlg.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls";

            // Set Multiselect option.
            dlg.Multiselect = true;


            // Display OpenFileDialog by calling ShowDialog method 
            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                foreach (var item in dlg.FileNames)
                {
                    lbExcelFiles.Items.Add(item);
                }
            }

            tbMessage.Text = "Total Files: " + dlg.FileNames.Count();
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            int c = 0;

            btnProcess.IsEnabled = false;



            thread = new Thread(new ThreadStart(() =>
            {
                this.Dispatcher.Invoke(() =>
                {
                    progressBar.Minimum = 0;
                    progressBar.Maximum = dlg.FileNames.Count();
                    tbMessage.Text = String.Format("Files to be examined: {0}.", progressBar.Maximum);
                });

                foreach (string xlsFileName in dlg.FileNames)
                {
                    //Create excel application
                    Excel.Application xlApp;
                    object misValue = System.Reflection.Missing.Value;
                    xlApp = new Excel.Application();
                    //xlApp.Visible = true;

                    //Load excel workbook
                    Excel.Workbook xlWbk = xlApp.Workbooks.Open(xlsFileName);

                    foreach (Excel.Worksheet item in xlWbk.Sheets)
                    {
                        if (item.Name == "Appendix 3")
                        {
                            Info info = new Info();

                            Excel.Worksheet xlSheet;


                            xlSheet = xlWbk.Sheets["Ship - Company Info Input"];

                            string company = "UNDIFINED";
                            if (xlSheet.Range["D35"].Value2 != null)
                            {
                                info.CompanyName = ((string)xlSheet.Range["D35"].Value2.ToString()).ToUpper();
                            }
                            else
                            {
                                info.CompanyName = company;
                            }



                            xlSheet = xlWbk.Sheets["Appendix 3"];

                            info.FileName = xlsFileName;

                            // Start Date
                            string sd = xlSheet.Range["AI5"].Value2.ToString();
                            double tempsd = 0.0;
                            if (Double.TryParse(sd, out tempsd))
                            {
                                info.StartDate = DateTime.FromOADate(Math.Round(double.Parse(sd)));
                            }
                            else
                            {
                                info.StartDate = DateTime.Parse(sd);
                            }


                            //End Date
                            string ed = xlSheet.Range["AH5"].Value2.ToString();
                            double temped = 0.0;
                            if (Double.TryParse(ed, out temped))
                            {
                                info.EndDate = DateTime.FromOADate(Math.Round(double.Parse(ed)));
                            }
                            else
                            {
                                info.EndDate = DateTime.Parse(ed);
                            }


                            info.IMONumber = xlSheet.Range["AG5"].Value2.ToString();

                            info.ShipType = ((string)xlSheet.Range["AF5"].Value2.ToString()).ToUpper();

                            int gt = -1;
                            Int32.TryParse(xlSheet.Range["AE5"].Value2.ToString(), out gt);
                            info.GrossTonnage = gt;

                            int nt = -1;
                            Int32.TryParse(xlSheet.Range["AD5"].Value2.ToString(), out nt);
                            info.NetTonnage = nt;

                            int dwt = -1;
                            Int32.TryParse(xlSheet.Range["AC5"].Value2.ToString(), out dwt);
                            info.DeadWeight = dwt;

                            double eedi = 0;
                            Double.TryParse(xlSheet.Range["AB5"].Value2.ToString(), out eedi);
                            info.EEDI = eedi;


                            // ICE Class
                            string tempICEClass = xlSheet.Range["Z5"].Value2.ToString();
                            if (tempICEClass == "NA" || tempICEClass == "-" || tempICEClass == "0")
                                tempICEClass = "N/A";
                            info.ICEClass = tempICEClass;


                            double mppower = 0;
                            Double.TryParse(xlSheet.Range["V5"].Value2.ToString(), out mppower);
                            info.MPPower = mppower;

                            double eppower = 0;
                            Double.TryParse(xlSheet.Range["U5"].Value2.ToString(), out eppower);
                            info.EPPower = eppower;

                            double dist = 0;
                            Double.TryParse(xlSheet.Range["T5"].Value2.ToString(), out dist);
                            info.DistanceTraveled = dist;

                            double hours = 0;
                            Double.TryParse(xlSheet.Range["S5"].Value2.ToString(), out hours);
                            info.HoursUnderway = hours;

                            double doil = 0;
                            Double.TryParse(xlSheet.Range["P5"].Value2.ToString(), out doil);
                            info.DO = doil;

                            double lfo = 0;
                            Double.TryParse(xlSheet.Range["O5"].Value2.ToString(), out lfo);
                            info.LFO = lfo;

                            double hfo = 0;
                            Double.TryParse(xlSheet.Range["M5"].Value2.ToString(), out hfo);
                            info.HFO = hfo;

                            lstInfo.Add(info);
                        }
                    }

                    //Cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    ////release com objects to fully kill excel process from running in the background
                    //Marshal.ReleaseComObject(xlWksResult);

                    //close and release
                    xlWbk.Close(false, xlsFileName, misValue);
                    Marshal.ReleaseComObject(xlWbk);

                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);

                    this.Dispatcher.Invoke(() =>
                    {
                        c++;
                        progressBar.Value = c;
                        tbMessage.Text = String.Format("Files examined: {0}/{1}.", c, progressBar.Maximum);
                    });
                }

                this.Dispatcher.Invoke(() =>
                {
                    btnProcess.IsEnabled = true;
                    MessageBox.Show("Export Completed!");
                });

                this.Dispatcher.Invoke(() =>
                {
                    if (lstInfo != null)
                        dgv1.ItemsSource = lstInfo;
                });

            }));

            thread.Start();

        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            if (lstInfo != null)
                lstInfo.Clear();

            OpenFileDialog _openFileDialog = new OpenFileDialog();
            if (_openFileDialog.ShowDialog() == true)
            {
                try
                {
                    XmlSerializer _xmlFormatter = new XmlSerializer(typeof(ObservableCollection<Info>));
                    using (Stream _fileStream = new FileStream(_openFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        _fileStream.Position = 0;
                        lstInfo = (ObservableCollection<Info>)_xmlFormatter.Deserialize(_fileStream);
                    }
                    this.Title = "DCS Calculator v1.0 - " + _openFileDialog.FileName;
                    dgv1.ItemsSource = lstInfo;

                    tabControl.SelectedItem = tiValues;
                }
                catch (Exception)
                {
                    MessageBox.Show("Please try open the correct file type.");
                }
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog _saveFileDialog = new SaveFileDialog();
            if (_saveFileDialog.ShowDialog() == true)
            {
                XmlSerializer _xmlFormatter = new XmlSerializer(typeof(ObservableCollection<Info>));
                using (Stream _fileStream = new FileStream(_saveFileDialog.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    _fileStream.Position = 0;
                    _xmlFormatter.Serialize(_fileStream, lstInfo);
                }
            }
        }

        private void tabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabControl tb = (TabControl)sender;

            if (((TabItem)tb.SelectedItem).Name == "tiGraphs")
            {
                CreateGraphs();
            }

            if (((TabItem)tb.SelectedItem).Name == "tiAnalysis")
            {
                Analysis();
            }

        }

        private void menuRemoveItem_Click(object sender, RoutedEventArgs e)
        {
            if (dgv1.Items.Count > 0 && dgv1.SelectedIndex > -1)
            {
                var selection = dgv1.SelectedItems;

                List<Info> lstSelection = new List<Info>();
                foreach (var item in selection)
                {
                    lstSelection.Add((Info)item);
                }

                foreach (Info item in lstSelection)
                {
                    lstInfo.Remove(item);
                }
            }
            else
            {
                MessageBox.Show("No item to remove!");
            }
        }

        private void menuAddItem_Click(object sender, RoutedEventArgs e)
        {
            if (dgv1.SelectedIndex > -1)
            {
                lstInfo.Insert(dgv1.SelectedIndex + 1, new Info() { CompanyName = "CHARTWORLD", StartDate = new DateTime(2019, 1, 1), EndDate = new DateTime(2019, 12, 31), ShipType = "BULK CARRIER", ICEClass = "N/A" });
            }
            else
            {

                lstInfo.Add(new Info() { CompanyName = "CHARTWORLD", StartDate = new DateTime(2019, 1, 1), EndDate = new DateTime(2019, 12, 31), ShipType = "BULK CARRIER" });
                dgv1.ItemsSource = lstInfo;
            }
        }


        private void CreateGraphs()
        {

            //Select Points
            var selection = dgv1.SelectedItems;
            ObservableCollection<Info> selectedInfos = new ObservableCollection<Info>();
            foreach (Info item in selection)
                selectedInfos.Add(item);

            //Create Graph Model and Graph Axis
            //PlotModel myModel = new PlotModel() { Title = "Bureau Veritas Bulk Carriers Fleet - Size: " + selectedInfos.FirstOrDefault().VesselSize.ToString() };
            PlotModel myModel = new PlotModel() { Title = "StarBulk Vessels" };

            var valueXAxis = new LinearAxis() { Position = AxisPosition.Bottom, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Title = "Deadweight [Mt]" };
            var valueYAxis = new LinearAxis() { Position = AxisPosition.Left, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Title = "Di-Value" };
            myModel.Axes.Add(valueXAxis);
            myModel.Axes.Add(valueYAxis);


            //Create Scatter Points
            ObservableCollection<Info> values = selectedInfos;
            double[] x = values.Select(i => i.DeadWeight).ToArray<double>();
            double[] y = values.Select(i => i.Di).ToArray<double>();
            var scatterSeries = new ScatterSeries { MarkerType = MarkerType.Circle };
            myModel.Axes.Add(new LinearColorAxis { Position = AxisPosition.Right, Minimum = y.Min(), Maximum = y.Max(), HighColor = OxyColors.Red, LowColor = OxyColors.Blue });
            for (int i = 0; i < x.Count(); i++)
                scatterSeries.Points.Add(new ScatterPoint(x[i], y[i], 4, y[i]));
            myModel.Series.Add(scatterSeries);


            #region Create AER Trajectory
            //var lineSeries = new LineSeries { Color = OxyColors.Blue, Title = "AER Trajectory value" };
            //double[] a = new double[] { 0, 10000, 34999, 59999, 99999, 199999 };
            //double[] b = new double[] { 26.3, 7, 4.9, 3.9, 2.5, 2.4 };
            //List<DataPoint> dPoints = new List<DataPoint>();
            //for (int i = 0; i < a.Length - 1; i++)
            //{
            //    if (a[i] >= x.Min() && a[i] <= x.Max())
            //    {
            //        dPoints.Add(new DataPoint(a[i], b[i]));
            //        dPoints.Add(new DataPoint(a[i + 1], b[i]));
            //    }
            //}
            //if (dPoints.Count > 0)
            //{
            //    dPoints.Insert(0, new DataPoint(x.Min(), dPoints[0].Y));
            //    dPoints.Add(new DataPoint(x.Max(), dPoints[dPoints.Count - 1].Y));
            //}
            //else
            //{
            //    if (x.Min() <= 9999)
            //    {
            //        dPoints.Add(new DataPoint(x.Min(), b[0]));
            //        dPoints.Add(new DataPoint(x.Max(), b[0]));
            //    }
            //    else
            //    {
            //        for (int i = 0; i < a.Length; i++)
            //        {
            //            if (a[i] >= x.Min())
            //            {
            //                dPoints.Add(new DataPoint(x.Min(), b[i - 1]));
            //                dPoints.Add(new DataPoint(x.Max(), b[i - 1]));
            //                break;
            //            }
            //        }
            //    }
            //}
            //lineSeries.Points.AddRange(dPoints);
            //myModel.Series.Add(lineSeries);
            #endregion








            //Create Trend Line
            int order = 1;
            if (x.Length > order)
            {
                double[] p = MNet.Fit.Polynomial(x, y, order);
                myModel.Series.Add(new FunctionSeries(z => Polynomial.Evaluate(z, p), x.Min(), x.Max(), 3, "Test")
                {
                    Title = "Trend Line: y = " + p[0].ToString("F6") + "x + " + p[1].ToString("F6"),
                    Color = OxyColors.Green
                });
            }


            ////STARBULK
            //var query1 = lstInfo.Where(c => c.CompanyName == "STARBULK S.A." && c.VesselSize == selectedInfos.FirstOrDefault().VesselSize).ToList();
            //double[] x1 = query1.Select(i => i.DeadWeight).ToArray<double>();
            //double[] y1 = query1.Select(i => i.AER).ToArray<double>();
            //var scatterSeries1 = new ScatterSeries { MarkerType = MarkerType.Diamond };
            //myModel.Axes.Add(new LinearColorAxis { Position = AxisPosition.None, Minimum = y.Min(), Maximum = y.Max(), HighColor = OxyColors.Red, LowColor = OxyColors.Blue });
            //for (int i = 0; i < x1.Count(); i++)
            //    scatterSeries1.Points.Add(new ScatterPoint(x1[i], y1[i], 4, x1.Min()));
            //myModel.Series.Add(scatterSeries1);



            //y=0

            double[] x1 = new double[] { 10000, 200000, };
            double[] y1 = new double[] { 0, 0, };
            var lineSeries = new LineSeries { Color = OxyColors.Blue, Title = "AER Trajectory value" };
            List<DataPoint> dPoints = new List<DataPoint>();
            dPoints.Add(new DataPoint(50000, 0));
            dPoints.Add(new DataPoint(211000, 0));

            var lineSeries1 = new LineSeries { Color = OxyColors.Blue, Title = "" };
            lineSeries1.Points.AddRange(dPoints);
            myModel.Series.Add(lineSeries1);





            //Create Graph
            Graph.Model = myModel;


            //Export graph to Clipboard
            var pngExporter = new ox.PngExporter
            {
                Width = 1024,
                Height = 768,
                Background = OxyColors.White
            };
            var bitmap = pngExporter.ExportToBitmap(myModel);
            Clipboard.SetImage(bitmap);

        }


        private void Analysis()
        {
            //var query1 = lstInfo
            //                   .GroupBy(x => new { x.CompanyName, x.VesselSize })
            //                   .OrderBy(g => g.Key.CompanyName).ThenBy(g => g.Key.VesselSize)
            //                   .Select(g => new
            //                   {
            //                       Company = g.Key.CompanyName,
            //                       VSize = g.Key.VesselSize,
            //                       AvAER = g.Average(x => x.AER),
            //                       AvDAER = g.Average(x => x.DAER),
            //                       TrajectoryAER = g.FirstOrDefault().AERTrajectory,
            //                       StandardDev = MNet.Statistics.Statistics.StandardDeviation(g.Select(x => x.AER).ToList())
            //                   });


            //var query1 = lstInfo
            //                  .GroupBy(x => new { x.VesselSize })
            //                  .OrderBy(g => g.Key.VesselSize)
            //                  .Select(g => new
            //                  {
            //                      VSize = g.Key.VesselSize,
            //                      AvAER = g.Average(x => x.AER),
            //                      AvDAER = g.Average(x => x.DAER),
            //                      TrajectoryAER = g.FirstOrDefault().AERTrajectory,
            //                      StandardDev = MNet.Statistics.Statistics.StandardDeviation(g.Select(x => x.AER).ToList())
            //                  });


            //var query2 = lstInfo
            //               .GroupBy(x => new { x.VesselSize })
            //               .OrderBy(g => g.Key.VesselSize)
            //               .Select(g => new
            //               {
            //                   VSize = g.Key.VesselSize,
            //                   AvAER = g.Average(x => x.AER),
            //                   AvDAER = g.Average(x => x.DAER),
            //                   TrajectoryAER = g.FirstOrDefault().AERTrajectory,
            //                   StandardDev = MNet.Statistics.Statistics.StandardDeviation(g.Select(x => x.AER).ToList()),
            //                   Quantile25 = MNet.Statistics.Statistics.PopulationStandardDeviation(g.Select(x => x.AER).ToList()),
            //                   Quantile50 = MNet.Statistics.Statistics.Percentile(g.Select(x => x.AER).ToList(), 50),
            //                   Quantile75 = MNet.Statistics.Statistics.Percentile(g.Select(x => x.AER).ToList(), 75)
            //               });







            //var query2 = from info in lstInfo
            //             group info by new { info.CompanyName, info.VesselSize } into eGroup
            //             orderby eGroup.Key.CompanyName, eGroup.Key.VesselSize
            //             select new
            //             {
            //                 Company = eGroup.Key.CompanyName,
            //                 VSize = eGroup.Key.VesselSize,
            //                 AvAER = eGroup.Average(x => x.AER),
            //                 AvDAER = eGroup.Average(x => x.DAER),
            //                 TrajectoryAER = eGroup.FirstOrDefault().AERTrajectory,
            //                 StandardDev = MNet.Statistics.Statistics.StandardDeviation(eGroup.Select(x => x.AER).ToList())
            //             };


            var query2 = from info in lstInfo
                         group info by new { info.CompanyName, info.VesselSize } into eGroup
                         orderby eGroup.Key.CompanyName, eGroup.Key.VesselSize
                         let xm = eGroup.Average(x => x.AER)
                         let sd = MNet.Statistics.Statistics.StandardDeviation(eGroup.Select(x => x.AER).ToList())
                         select new
                         {
                             Company = eGroup.Key.CompanyName,
                             VSize = eGroup.Key.VesselSize,
                             AvAER = lstInfo.Average(x => x.AER),
                             TrajectoryAER = eGroup.FirstOrDefault().AERTrajectory,
                             StandardDev = MNet.Statistics.Statistics.StandardDeviation(eGroup.Select(x => x.AER).ToList()),
                             ZScore = eGroup.Select(x => (x.AER - xm) / sd).ToList()
                         };




            var query1 = lstInfo.Select(g => new
            {
                Company = g.CompanyName,
                IMO = g.IMONumber,
                GT = g.GrossTonnage,
                DWT = g.DeadWeight,
                AER_Value = g.AER,
                StandardDev = MNet.Statistics.Statistics.StandardDeviation(lstInfo.Select(x => x.AER).ToList())
            });


            var query3 = from g in lstInfo
                         let xm = lstInfo.Where(y => y.VesselSize == g.VesselSize).Select(x => x.AER).Average()
                         let sd = MNet.Statistics.Statistics.StandardDeviation(lstInfo.Where(y => y.VesselSize == g.VesselSize).Select(x => x.AER).ToList())
                         select new
                         {
                             Company = g.CompanyName,
                             VesselSize = g.VesselSize,
                             IMO = g.IMONumber,
                             GT = g.GrossTonnage,
                             DWT = g.DeadWeight,
                             AER_Value = g.AER,
                             StandardDev = sd,
                             ZScore = (g.AER - xm) / sd
                         };


            var query4 = lstInfo
                              .GroupBy(x => new { x.VesselSize })
                              .OrderBy(g => g.Key.VesselSize)
                              .Select(g => new
                              {
                                  VSize = g.Key.VesselSize,
                                  AvAER = g.Average(x => x.AER),
                                  StandardDev = MNet.Statistics.Statistics.StandardDeviation(g.Select(x => x.AER).ToList())
                              });





            dgv2.ItemsSource = query4;
            dgv3.ItemsSource = query3;
        }




    }
}