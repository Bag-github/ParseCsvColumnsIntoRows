using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ParseCarvaygoPricingSheet
{
    class Program
    {
        static void Main(string[] args)
        {
            ParseLucidPricing();
            //ParseCarvaygoPricing();
            //FormatLatitudeLongitude();
        }

        protected static void FormatLatitudeLongitude()
        {
            using (XLWorkbook workBook = new XLWorkbook(@"E:\Webfolder\RPMFrieght\ParseCarvaygoPricingSheet\Resource\Markets 202207.xlsx"))
            {

                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;

                //custom variables
                string UnitTypeId = string.Empty;
                string MarketCombination = string.Empty;
                string MarketOrigin = string.Empty;
                string MarketDestination = string.Empty;
                string DesireabilityFactor = string.Empty;

                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                            firstRow = false;
                        }
                    }
                    else
                    {

                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            //(double)num1/(double)num2;
                            if (i == 4)
                                dt.Rows[dt.Rows.Count - 1][i] = (double)Int32.Parse(cell.Value.ToString().Replace(" N", "")) / 10000;
                            else if (i == 5)
                                dt.Rows[dt.Rows.Count - 1][i] = "-" + (double)Int32.Parse(cell.Value.ToString().Replace(" W", "")) / 10000;
                            else
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();

                            i++;
                        }
                    }
                }


                //Name of File  
                string fileName = "TransformedMarkets 202207.xlsx";
                string path = Path.Combine(@"E:\Webfolder\RPMFrieght\ParseCarvaygoPricingSheet\Output\", fileName);
                using (XLWorkbook wb = new XLWorkbook())
                {
                    //Add DataTable in worksheet  
                    dt.TableName = "CarvaygoMarketDefinitions";

                    var ws = wb.Worksheets.Add(dt, "CarvaygoMarketDefinitions");

                    wb.SaveAs(path);
                    wb.Dispose();


                }

            }
        }

        protected static void ParseCarvaygoPricing()
        {

            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(@"E:\Webfolder\RPMFrieght\ParseCarvaygoPricingSheet\Resource\PricingDictionary.xlsx"))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;

                //custom variables
                string UnitTypeId = string.Empty;
                string MarketCombination = string.Empty;
                string MarketOrigin = string.Empty;
                string MarketDestination = string.Empty;
                string DesireabilityFactor = string.Empty;

                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        //foreach (IXLCell cell in row.Cells())
                        //{
                        //    dt.Columns.Add(cell.Value.ToString());
                        //}
                        dt.Columns.Add("UnitTypeId");
                        dt.Columns.Add("MarketCombination");
                        dt.Columns.Add("MarketOrigin");
                        dt.Columns.Add("MarketDestination");
                        dt.Columns.Add("DesireabilityFactor");
                        dt.Columns.Add("MilesLow");
                        dt.Columns.Add("MilesHigh");
                        dt.Columns.Add("BasePrice");

                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            if (i > 4)
                            {

                                if (i == 5)
                                {
                                    dt.Rows[dt.Rows.Count - 1][i - 1] = cell.Value.ToString();
                                    DesireabilityFactor = cell.Value.ToString();
                                }
                                else
                                {

                                    if (i == 6)
                                    {
                                        dt.Rows[dt.Rows.Count - 1][5] = "0";
                                        dt.Rows[dt.Rows.Count - 1][6] = "50";
                                        dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                    }
                                    else
                                    {
                                        //Add rows to DataTable.
                                        dt.Rows.Add();
                                        dt.Rows[dt.Rows.Count - 1][0] = UnitTypeId;
                                        dt.Rows[dt.Rows.Count - 1][1] = MarketCombination;
                                        dt.Rows[dt.Rows.Count - 1][2] = MarketOrigin;
                                        dt.Rows[dt.Rows.Count - 1][3] = MarketDestination;
                                        dt.Rows[dt.Rows.Count - 1][4] = DesireabilityFactor;

                                        if (i == 7)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "51";
                                            dt.Rows[dt.Rows.Count - 1][6] = "75";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 8)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "76";
                                            dt.Rows[dt.Rows.Count - 1][6] = "100";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 9)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "101";
                                            dt.Rows[dt.Rows.Count - 1][6] = "150";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 10)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "151";
                                            dt.Rows[dt.Rows.Count - 1][6] = "250";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 11)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "251";
                                            dt.Rows[dt.Rows.Count - 1][6] = "300";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 12)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "301";
                                            dt.Rows[dt.Rows.Count - 1][6] = "400";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 13)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "401";
                                            dt.Rows[dt.Rows.Count - 1][6] = "500";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 14)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "501";
                                            dt.Rows[dt.Rows.Count - 1][6] = "600";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 15)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "601";
                                            dt.Rows[dt.Rows.Count - 1][6] = "700";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 16)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "701";
                                            dt.Rows[dt.Rows.Count - 1][6] = "800";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 17)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "801";
                                            dt.Rows[dt.Rows.Count - 1][6] = "900";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 18)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "901";
                                            dt.Rows[dt.Rows.Count - 1][6] = "1000";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 19)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "1001";
                                            dt.Rows[dt.Rows.Count - 1][6] = "1200";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 20)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "1201";
                                            dt.Rows[dt.Rows.Count - 1][6] = "1500";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 21)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "1501";
                                            dt.Rows[dt.Rows.Count - 1][6] = "1750";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 22)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "1751";
                                            dt.Rows[dt.Rows.Count - 1][6] = "2000";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 23)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "2001";
                                            dt.Rows[dt.Rows.Count - 1][6] = "2500";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 24)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][5] = "2501";
                                            dt.Rows[dt.Rows.Count - 1][6] = "999999";
                                            dt.Rows[dt.Rows.Count - 1][7] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                    }

                                }


                            }

                            else
                            {
                                if (i == 0)
                                    UnitTypeId = cell.Value.ToString();
                                else if (i == 1)
                                    MarketCombination = cell.Value.ToString();
                                else if (i == 2)
                                    MarketOrigin = cell.Value.ToString();
                                else if (i == 3)
                                    MarketDestination = cell.Value.ToString();


                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();

                            }

                            i++;
                        }
                    }

                }

                //Name of File  
                string fileName = "TransformedCarvaygoPricing.xlsx";
                string path = Path.Combine(@"E:\Webfolder\RPMFrieght\ParseCarvaygoPricingSheet\Output\", fileName);
                using (XLWorkbook wb = new XLWorkbook())
                {
                    //Add DataTable in worksheet  
                    dt.TableName = "CarvaygoPricing";

                    var ws = wb.Worksheets.Add(dt, "CarvaygoPricing");

                    wb.SaveAs(path);
                    wb.Dispose();


                }
            }

        }

        protected static void ParseLucidPricing()
        {

            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(@"C:\Webfolder\RPMFrieght\ParseCarvaygoPricingSheet\Resource\LucidFinalMileEnclosedPricing.xlsx"))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;

                //custom variables
                string City = string.Empty;
                string State = string.Empty;
                string Permile = string.Empty;

                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        //foreach (IXLCell cell in row.Cells())
                        //{
                        //    dt.Columns.Add(cell.Value.ToString());
                        //}
                        dt.Columns.Add("City");
                        dt.Columns.Add("State");
                        dt.Columns.Add("PerMile");
                        dt.Columns.Add("MilesLow");
                        dt.Columns.Add("MilesHigh");
                        dt.Columns.Add("Cost");


                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            if (i > 1)
                            {

                                if (i == 2)
                                {
                                    dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                    Permile = cell.Value.ToString();
                                }
                                else
                                {

                                    if (i == 3)
                                    {
                                        dt.Rows[dt.Rows.Count - 1][3] = "0";
                                        dt.Rows[dt.Rows.Count - 1][4] = "25";
                                        dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                    }
                                    else
                                    {
                                        //Add rows to DataTable.
                                        dt.Rows.Add();
                                        dt.Rows[dt.Rows.Count - 1][0] = City;
                                        dt.Rows[dt.Rows.Count - 1][1] = State;
                                        dt.Rows[dt.Rows.Count - 1][2] = Permile;

                                        if (i == 4)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "26";
                                            dt.Rows[dt.Rows.Count - 1][4] = "50";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 5)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "51";
                                            dt.Rows[dt.Rows.Count - 1][4] = "75";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 6)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "76";
                                            dt.Rows[dt.Rows.Count - 1][4] = "100";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 7)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "200";
                                            dt.Rows[dt.Rows.Count - 1][4] = "299";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 8)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "300";
                                            dt.Rows[dt.Rows.Count - 1][4] = "399";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 9)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "400";
                                            dt.Rows[dt.Rows.Count - 1][4] = "400";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 10)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "500";
                                            dt.Rows[dt.Rows.Count - 1][4] = "500";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 11)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "600";
                                            dt.Rows[dt.Rows.Count - 1][4] = "600";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 12)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "700";
                                            dt.Rows[dt.Rows.Count - 1][4] = "700";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 13)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "800";
                                            dt.Rows[dt.Rows.Count - 1][4] = "800";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 14)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "900";
                                            dt.Rows[dt.Rows.Count - 1][4] = "900";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 15)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "1000";
                                            dt.Rows[dt.Rows.Count - 1][4] = "1000";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 16)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "1100";
                                            dt.Rows[dt.Rows.Count - 1][4] = "1100";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }
                                        else if (i == 17)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][3] = "1200";
                                            dt.Rows[dt.Rows.Count - 1][4] = "1200";
                                            dt.Rows[dt.Rows.Count - 1][5] = decimal.Parse(cell.Value.ToString()).ToString("########.00");
                                        }

                                    }

                                }


                            }

                            else
                            {
                                if (i == 0)
                                    City = cell.Value.ToString();
                                else if (i == 1)
                                    State = cell.Value.ToString();

                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();

                            }

                            i++;
                        }
                    }

                }

                //Name of File  
                string fileName = "TransformedLucidPricing.xlsx";
                string path = Path.Combine(@"C:\Webfolder\RPMFrieght\ParseCarvaygoPricingSheet\Output\", fileName);
                using (XLWorkbook wb = new XLWorkbook())
                {
                    //Add DataTable in worksheet  
                    dt.TableName = "CarvaygoPricing";

                    var ws = wb.Worksheets.Add(dt, "CarvaygoPricing");

                    wb.SaveAs(path);
                    wb.Dispose();


                }
            }

        }
    }
}
