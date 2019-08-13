using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace TEMAutomatedReports
{
    class GraphicalReport
    {



        public static void GenerateGraph(string file, DataTable dt, string graphbind, string name, string filepath, string format)
        {
            try
            {

                Chart chart = new Chart();
                chart.Legends.Add("Legend1");
                chart.Width = 700;
                chart.Height = 360;
                chart.ChartAreas.Add("ChartArea1");
                chart.Palette = ChartColorPalette.SemiTransparent;
                chart.BackColor = System.Drawing.ColorTranslator.FromHtml("#3B414D");
                chart.BackGradientStyle = GradientStyle.DiagonalLeft;
                chart.BorderlineColor = Color.DarkSalmon;
                //chart.gr.BackGradientEndColor = "White";
                Color[] seriescolurs = new Color[]
            {
             System.Drawing.Color.MidnightBlue,             //dashboard
             System.Drawing.Color.LightSeaGreen ,           //Landline
             System.Drawing.Color.Firebrick,                //Tariff
             System.Drawing.Color.Goldenrod,                //assert
             System.Drawing.Color.DarkSalmon,               //bill val
             System.Drawing.Color.OrangeRed,                //KPI
             System.Drawing.Color.Crimson,                  //Inventory
             System.Drawing.Color.Gold,                     //mobile
             System.Drawing.Color.DarkRed ,
             System.Drawing.Color.Cyan ,                    //Admin
             System.Drawing.Color.Tomato,                   //Trend
             System.Drawing.Color.MediumSeaGreen          //Reporting
            };

                string[] ax = graphbind.Split('#');
                string xaxis = ax[0].Split('@')[0];
                if (chart.Series.Count() >= 0)
                {
                    chart.Series.Clear();
                }
                bool Durationformat = false;
                Series ser; int i = 0;
                foreach (string s in ax[1].Split(','))
                {
                    if (s != string.Empty)
                    {
                        ser = new Series();
                        ser.Name = s.Split('@')[0];
                        ser.XValueMember = xaxis;
                        ser.YValueMembers = s.Split('@')[0];
                        ser.YValueType = ChartValueType.Double;
                        SetGraphtype(format, ser);
                        ser.BorderWidth = 0;
                        ser.BorderColor = Color.Gray;
                        ser.Color = seriescolurs[i];
                        chart.Series.Add(ser);
                        if (s.Split('@')[0].Contains("Duration"))
                        {
                            Durationformat = true;
                        }
                        i++;
                    }
                }
                chart.Titles.Clear();
                chart.Titles.Add(name);
                chart.ChartAreas["ChartArea1"].AxisY.Title = "Calls";
                chart.BackColor = Color.Empty;
                chart.ChartAreas["ChartArea1"].BorderWidth = 0;
                chart.ChartAreas["ChartArea1"].BackColor = Color.Gray;
                chart.ChartAreas["ChartArea1"].ShadowColor = Color.Transparent;
                chart.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                chart.ChartAreas["ChartArea1"].AxisX.LineColor = Color.Gray;
                chart.ChartAreas["ChartArea1"].AxisX.TitleFont = new Font("Century Gothic", 8.25f, FontStyle.Bold);
                chart.ChartAreas["ChartArea1"].AxisX.MajorGrid.LineColor = Color.Gray;
                chart.ChartAreas["ChartArea1"].AxisX.Title = xaxis;
                chart.ChartAreas["ChartArea1"].AxisX.LabelStyle.Angle = 45;
                chart.ChartAreas["ChartArea1"].AxisY.LineColor = Color.Gray;
                chart.ChartAreas["ChartArea1"].AxisY.TitleFont = new Font("Century Gothic", 8.25f, FontStyle.Bold);
                chart.ChartAreas["ChartArea1"].AxisY.MajorGrid.LineColor = Color.Gray;
                chart.ChartAreas["ChartArea1"].BackGradientStyle = GradientStyle.TopBottom;
                chart.ChartAreas["ChartArea1"].BackHatchStyle = ChartHatchStyle.None;
                chart.BorderSkin.BackColor = Color.LightSkyBlue;
                chart.BorderSkin.PageColor = Color.Transparent;
                chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;


                if (Durationformat == false)
                {
                    chart.DataSource = dt.AsEnumerable().Take(12);
                }
                else
                {

                    DataTable dc = new DataTable();
                    dc.Columns.Add(xaxis);
                    foreach (string s in ax[1].Split(','))
                    {
                        dc.Columns.Add(s.Split('@')[0]);

                    }
                    var result = from db in dt.AsEnumerable()
                                 select new
                                 {
                                     Period = db.Field<string>("Period"),
                                     AvgRing = GenerateReports.GetDurationinsec(db.Field<string>("Avg Ring")),
                                     AvgCallDuration = GenerateReports.GetDurationinsec(db.Field<string>("Avg Call Duration")),
                                     AvgAbandonedTime = GenerateReports.GetDurationinsec(db.Field<string>("Avg Abandoned Time")),
                                 };

                    foreach (var v in result)
                    {
                        dc.Rows.Add(v.Period, v.AvgRing, v.AvgCallDuration, v.AvgAbandonedTime);

                    }
                    chart.DataSource = dc.AsEnumerable().Take(12);
                }
                chart.DataBind();
                chart.SaveImage(file, ChartImageFormat.Png);


            }
            catch
            {

            }
        }
        private static void SetGraphtype(string format , Series ser)
        {
            switch (format)
            {
                case "Bar":

                    ser.ChartType = SeriesChartType.Bar;
                    ser.BorderColor = Color.White;
                    ser.BackGradientStyle = GradientStyle.None;
                    ser.BackHatchStyle = ChartHatchStyle.None;
                    break;
                case "Radar":
                    ser.ChartType = SeriesChartType.Radar;
                    break;
                case "SplineArea":
                    ser.ChartType = SeriesChartType.SplineArea;
                    break;
                case "SplineRange":
                    ser.ChartType = SeriesChartType.SplineRange;
                    break;
                case "Stock":
                    ser.ChartType = SeriesChartType.Stock;
                    break;
                case "Pyramid":
                    ser.ChartType = SeriesChartType.Pyramid;
                    break;
                case "Pie":
                    ser.ChartType = SeriesChartType.Pie;
                    break;
                case "Bubble":
                    ser.ChartType = SeriesChartType.Bubble;
                    break;
                case "Doughnut":
                    ser.ChartType = SeriesChartType.Doughnut;
                    break;
                case "Area":
                    ser.ChartType = SeriesChartType.StackedArea;
                    break;
                case "100% Area":
                    ser.ChartType = SeriesChartType.StackedArea100;
                    break;
                case "Column":
                    ser.ChartType = SeriesChartType.Column;
                    ser.CustomProperties = "DrawingStyle=Cylinder";
                    ser.BackGradientStyle = GradientStyle.None;
                    ser.BackHatchStyle = ChartHatchStyle.None;
                    break;
                case "Sacked Column":
                    ser.ChartType = SeriesChartType.StackedColumn;
                    ser.CustomProperties = "DrawingStyle=Cylinder";
                    break;
                case "100% Column":
                    ser.ChartType = SeriesChartType.StackedColumn100;
                    ser.CustomProperties = "DrawingStyle=Cylinder";
                    break;
                case "Line":
                    ser.ChartType = SeriesChartType.Line;
                    ser.BorderWidth = 5;
                    break;
                case "StepLine":
                    ser.ChartType = SeriesChartType.StepLine;
                    ser.BorderWidth = 5;
                    break;
                
                default:
                   ser.ChartType = SeriesChartType.StackedColumn;
                   ser.CustomProperties = "DrawingStyle=Cylinder";
                    break;
            }
           
        }

        public static void CreateGraphsinExcel(List<string> files, string filename , string reportname)
        {
            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Cells[1, 1] = "Graphical Reoport";
                xlWorkSheet.Cells[2, 1] = reportname;
                float top = 50;
                foreach (string pic in files)
                {
                    xlWorkSheet.Shapes.AddPicture(pic, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 50, top, 600, 400);
                    top = 50 + 500;
                   // xlWorkSheet.Shapes.AddPicture(pic, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 50, 50, 300, 45);
                }

                xlWorkBook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();


            }
            catch { }
        }
        public static byte[] Imagetobytearray(string imagefilepath)
        {
            Image im = Image.FromFile(imagefilepath);
            byte[] imageByte = ImagetoBytearraybyimageconverter(im);
            return imageByte;
        
        }
        public static byte[] ImagetoBytearraybyimageconverter(Image im)
        {
            ImageConverter imcon = new ImageConverter();
            byte[] imageByte = (byte[])imcon.ConvertTo(im, typeof(byte[]));
            return imageByte;
        
        }
        public static void GenerateCircularGauge(string file, double min, double medium, double max, double pointervalue, string name, string color)
        {
            try
            {
                AGauge Gauge = new AGauge();
                // Gauge properties
                Gauge.BaseArcColor = Color.Gray;
                Gauge.BaseArcRadius = 80;
                Gauge.BaseArcStart = 135;
                Gauge.BaseArcSweep = 270;
                Gauge.BaseArcWidth = 2;
                //Gauge.Center.X= 100;
                //Gauge.Center.Y = 100;
                Gauge.MaxValue = (float)max + 3;
                Gauge.MinValue = 0;
                Gauge.NeedleColor1 = AGaugeNeedleColor.Blue;
                Gauge.NeedleColor2 = Color.Blue;
                Gauge.NeedleRadius = 80;
                Gauge.NeedleType = NeedleType.Advance;
                Gauge.NeedleWidth = 2;
                Gauge.ScaleLinesInterColor = Color.Black;
                Gauge.ScaleLinesInterInnerRadius = 73;
                Gauge.ScaleLinesInterOuterRadius = 80;
                Gauge.ScaleLinesInterWidth = 1;
                Gauge.ScaleLinesMajorColor = Color.Black;
                Gauge.ScaleLinesMajorInnerRadius = 70;
                Gauge.ScaleLinesMajorOuterRadius = 80;
                Gauge.ScaleLinesMajorStepValue = 1;
                Gauge.ScaleLinesMajorWidth = 2;
                Gauge.ScaleLinesMinorColor = Color.LightGray;
                Gauge.ScaleLinesMinorInnerRadius = 75;
                Gauge.ScaleLinesMinorOuterRadius = 80;
                Gauge.ScaleLinesMinorTicks = 9;
                Gauge.ScaleLinesMinorWidth = 1;
                Gauge.ScaleNumbersColor = Color.Black;
                Gauge.ScaleNumbersRadius = 62;
                Gauge.ScaleNumbersRotation = 0;
                Gauge.ScaleNumbersStartScaleLine = 0;
                Gauge.ScaleNumbersStepScaleLines = 1;
                Gauge.Value = (float)pointervalue;
                // Appearance
                Gauge.BackColor = Color.White;
                Gauge.Name = "Gauge1";
                Gauge.Width = 250;
                Gauge.Height = 200;
                AGaugeRange range1 = new AGaugeRange();
                range1.Color = Color.FromArgb(141, 203, 42); range1.InnerRadius = 70; range1.OuterRadius = 80;
                range1.StartValue = 0;
                range1.EndValue = (float)min;
                AGaugeRange range2 = new AGaugeRange();
                range2.Color = Color.FromArgb(255, 122, 0); range2.InnerRadius = 70; range2.OuterRadius = 80;
                range2.StartValue = (float)min;
                range2.EndValue = (float)medium;
                AGaugeRange range3 = new AGaugeRange();
                range3.Color = Color.FromArgb(194, 0, 0); range3.InnerRadius = 70; range3.OuterRadius = 80;
                range3.StartValue = (float)medium;
                range3.EndValue = (float)max + 3;
                Gauge.GaugeRanges.Add(range1);
                Gauge.GaugeRanges.Add(range2);
                Gauge.GaugeRanges.Add(range3);
                AGaugeLabel label1 = new AGaugeLabel();
                label1.Color = Color.Black;
                label1.Text = name; Point p1 = new Point(); p1.X = 48; p1.Y = 1; label1.Position = p1;
                Gauge.GaugeLabels.Add(label1);
                AGaugeLabel label2 = new AGaugeLabel();
                label2.Color = Color.Black;
                if (color == "Table_Rows_Green")
                    label2.Color = Color.FromArgb(141, 203, 42);
                else if (color == "Table_Rows_orange")
                    label2.Color = Color.FromArgb(255, 122, 0);
                else
                    label2.Color = Color.FromArgb(194, 0, 0);
                label2.Text = pointervalue.ToString(); Point p2 = new Point(); p2.X = 85; p2.Y = 130; label2.Position = p2;
                Gauge.GaugeLabels.Add(label2);
                Bitmap bmp = new Bitmap(Gauge.Width, Gauge.Height);
                Gauge.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                bmp.Save(file);
            }
            catch { }
        }

    }
}


