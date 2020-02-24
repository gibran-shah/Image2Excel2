using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using ExcelRef = Microsoft.Office.Interop.Excel;

namespace Image2Excel
{
    class Engine
    {
        public static void go(string imageFilename, string excelFilename = ".\\output.xlsx")
        {
            Console.WriteLine("Image file: " + imageFilename);
            if (!excelFilename.StartsWith(".\\")) excelFilename = ".\\" + excelFilename;

            Console.WriteLine("Writing to " + excelFilename);

            Bitmap btm = (Bitmap)Bitmap.FromFile(imageFilename, false);

            if (btm != null)
            {
                int totalPixels = btm.Width * btm.Height;
                int pixelCount = 0;
                int percentDone = 0;
                int lastPercentDone = 0;
                int cursorTop = Console.CursorTop;
                Console.WriteLine("Reading image file: 0%");
                Color[][] colorArray = new Color[btm.Width][];
                for (int x = 0; x < btm.Width; x++)
                {
                    colorArray[x] = new Color[btm.Height];
                    for (int y = 0; y < btm.Height; y++)
                    {
                        colorArray[x][y] = btm.GetPixel(x, y);
                        pixelCount++;
                        percentDone = (int)(((float)pixelCount / (float)totalPixels) * 100.0f);
                        if (percentDone > lastPercentDone)
                        {
                            Console.SetCursorPosition(0, cursorTop);
                            Console.WriteLine("Reading image file: " + percentDone + "%");
                            lastPercentDone = percentDone;
                        }
                    }
                }

                _Application excel = new ExcelRef.Application();
                Workbook workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet worksheet = workbook.Worksheets[1];

                worksheet.Columns.ColumnWidth = 2;
                worksheet.Rows.RowHeight = 15;

                pixelCount = 0;
                percentDone = 0;
                lastPercentDone = 0;
                cursorTop = Console.CursorTop;
                Console.WriteLine("Writing excel file: 0%");
                for (int x = 0; x < colorArray.Length; x++)
                {
                    for (int y = 0; y < colorArray[x].Length; y++)
                    {
                        worksheet.Cells[y + 1, x + 1].Interior.Color = colorArray[x][y];
                        pixelCount++;
                        percentDone = (int)(((float)pixelCount / (float)totalPixels) * 100.0f);
                        if (percentDone > lastPercentDone)
                        {
                            Console.SetCursorPosition(0, cursorTop);
                            Console.WriteLine("Writing excel file: " + percentDone + "%");
                            lastPercentDone = percentDone;
                        }
                    }
                }

                workbook.SaveAs(excelFilename);
                workbook.Close();

                Console.WriteLine("Done!");
            }
        }
    }
}
