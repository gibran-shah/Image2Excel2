using System;

namespace Image2Excel
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string imageFilename = "";
            string excelFilename = "";

            switch (args.Length)
            {
                case 0:
                    Console.WriteLine("Please supply an image file name and an optional Excel file name.");
                    Console.WriteLine("Example: ./Image2Excel [\"path/imagefile.ext\" [\"path/excelfile.xls\"]]");
                    return;
                case 1:
                    imageFilename = args[0];
                    break;
                case 2:
                default:
                    imageFilename = args[0];
                    excelFilename = args[1];
                    break;
            }

            Engine.go(imageFilename, excelFilename);
        }
    }
}

// https://github.com/gibran-shah/Image2Excel2

