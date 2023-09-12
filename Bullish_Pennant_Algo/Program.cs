using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Spire.Xls;

namespace Bullish_Pennant_Algo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelGrapher Graph_StockExcel = new ExcelGrapher();

            List<double> StockPricesList = new List<double>(); int StockPricesList_Iterations = 0;
            List<double> StandardDeviations_StockPricesList = new List<double>(); //Anchored by StockPricesList_Iterations therefore StockPrciesList_Iterations is added after StandardDeviation_StockPricesList.RemoveAt() is executed
            List<CsvLine> FlagPoleStreak_InformationList = new List<CsvLine>();
            CSVImporter RawData = new CSVImporter();
            RawData.Import("AAPL.csv");

            RawData.CsvLine_List.ForEach(RawDataValue => {

                StockPricesList.Add(Convert.ToDouble(RawDataValue.Close));

                if (StockPricesList.Count > 7) StockPricesList.RemoveAt(StockPricesList.Count - 8);
                
                double Average = 0;
                StockPricesList.ForEach(AllValuesInStockPricesList => {
                    Average += AllValuesInStockPricesList;
                }); Average = Average / StockPricesList.Count;

                double StandardDeviation = 0;
                double E = 0;
                StockPricesList.ForEach(AllValuesInStockPricesList_2 => {
                    E += Math.Pow(AllValuesInStockPricesList_2 - Average, 2);
                }); StandardDeviation = Math.Sqrt(E / (StockPricesList.Count - 1));
                

                if (Double.IsNaN(StandardDeviation)) { } else StandardDeviations_StockPricesList.Add(StandardDeviation);
                
                if (StandardDeviations_StockPricesList.Count > 7) StandardDeviations_StockPricesList.RemoveAt(StandardDeviations_StockPricesList.Count - 8);
                StockPricesList_Iterations++;

                double AverageOfStandardDeviations = 0;
                StandardDeviations_StockPricesList.ForEach(StandardDeviations_ForEachValue => {
                    AverageOfStandardDeviations += StandardDeviations_ForEachValue;
                }); AverageOfStandardDeviations = AverageOfStandardDeviations / StandardDeviations_StockPricesList.Count;
                
                if (StandardDeviation > AverageOfStandardDeviations)
                {
                    FlagPoleStreak_InformationList.Add(RawDataValue);
                }
                else
                {
                    
                    if (FlagPoleStreak_InformationList.Count > 2)
                    {
                        double Difference = Convert.ToDouble(FlagPoleStreak_InformationList.Last().Close) - Convert.ToDouble(FlagPoleStreak_InformationList.First().Close);
                        Console.WriteLine(FlagPoleStreak_InformationList.First().Date + " " + FlagPoleStreak_InformationList.Last().Date + " " + Difference);

                        Graph_StockExcel.Initialize();
                        Graph_StockExcel.AddRow(FlagPoleStreak_InformationList.Last().Date, FlagPoleStreak_InformationList.First().Date, Difference);
                        FlagPoleStreak_InformationList.Clear();
                    }
                    
                }
            });
            Graph_StockExcel.CreateChart();
            Graph_StockExcel.SaveExcelBook();
            Console.WriteLine("Done");
        }
    }
}
class CsvLine
{
    public string Date { get; set; }
    public string Open { get; set; }
    public string High { get; set; }
    public string Low { get; set; }
    public string Close { get; set; }

    public override string ToString()
    {
        return Date + " " + Open + " " + High + " " + Low + " " + Close;
    }
}

class CSVImporter
{
    public List<CsvLine> CsvLine_List = new List<CsvLine>();

    public void Import(string File_Name)
    {
        try
        {
            var File_Stream = new StreamReader(File_Name);
            

                CsvLine_List = new CsvHelper.CsvReader(File_Stream, System.Globalization.CultureInfo.CurrentCulture).GetRecords<CsvLine>().ToList();
            
        }
        catch (Exception e)
        {
            Console.WriteLine("Exception was thrown:");
            Console.WriteLine(e);
        }
        
    }
}

class ExcelGrapher
{
   private int Anchored_ExcelRowIndicator = 2;

    Workbook Excel_Book;
    Worksheet Worksheet1;
    public ExcelGrapher()
    {
        Excel_Book = new Workbook();
        Excel_Book.Version = ExcelVersion.Version2016;
    }

    public void Initialize()
    {
        Worksheet1 = Excel_Book.Worksheets[0];
        Worksheet1.Range["A1"].Text = "Starting_Date";
        Worksheet1.Range["B1"].Text = "Ending_Date";
        Worksheet1.Range["C1"].Text = "Deviations";
        Excel_Book.SaveToFile("AAPL_Indicative_Deviations.xlsx" , FileFormat.Version2016);
    }

    public void AddRow(string Starting_Date, string Ending_Date, double Deviations)
    {
        Worksheet1 = Excel_Book.Worksheets[0];
        Worksheet1.Range["A" + Anchored_ExcelRowIndicator].Text = Starting_Date;
        Worksheet1.Range["B" + Anchored_ExcelRowIndicator].Text = Ending_Date;
        Worksheet1.Range["C" + Anchored_ExcelRowIndicator].NumberValue = Deviations;
        Anchored_ExcelRowIndicator++;
    }

    public void CreateChart()
    {
        Chart ExcelChart = Worksheet1.Charts.Add(ExcelChartType.Line);
        ExcelChart.DataRange = Worksheet1.Range["A1:C201"];
        ExcelChart.SeriesDataFromRange = false;
        ExcelChart.PrimaryCategoryAxis.HasMajorGridLines = false;
        ExcelChart.LeftColumn = 4;
        ExcelChart.TopRow = 2;
        ExcelChart.RightColumn = 12;
        ExcelChart.BottomRow = 22;
    }

    public void SaveExcelBook()
    {
        Excel_Book.SaveToFile("AAPL_Indicative_Deviations.xlsx", FileFormat.Version2016);
    }
}