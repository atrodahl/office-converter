using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeConverter
{
    class Converter
    {
        static void Main(string[] args)
        {
            try
            {
                validateArguments(args);

                var inFile = new System.IO.FileInfo(args[0]);
                var outFile = new System.IO.FileInfo(args[1]);

                var converter = new Converter();
                converter.convert(inFile, outFile);

                Console.WriteLine(String.Format("Converted [{0}] to [{1}]", inFile.FullName, outFile.FullName));
                Environment.Exit(0);

            }
            catch (ArgumentException e)
            {
                printHelp(e.Message);
                Environment.Exit(1);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Environment.Exit(2);
            }
        }


        private static void validateArguments(string[] args)
        {
            if (args.Length < 2)
            {
                throw new ArgumentException("Not enough arguments");
            }

            var inFilePath = args[0];
            var outFilePath = args[1];
            
            if (!System.IO.File.Exists(inFilePath))
            {
                throw new ArgumentException(String.Format("Input file [{0}] does not exist", inFilePath));
            }
            
            var outDirectory = System.IO.Path.GetDirectoryName(outFilePath);
            if (!System.IO.Directory.Exists(outDirectory))
            {
                throw new ArgumentException(String.Format("Output directory [{0}] does not exist", outDirectory));
            }
        }

        private static void printHelp(string message = "")
        {
            Console.WriteLine(message);
            Console.WriteLine();
            Console.WriteLine("converter <in file> <out format>");
        }

        private void convert(System.IO.FileInfo inFile, System.IO.FileInfo outFile)
        {
            var inFormat = inFile.Extension.ToLower().Substring(1);
            switch (inFormat)
            {
                case "doc":
                case "docx":
                    convertDocument(inFile, outFile);
                    break;
                case "xls":
                case "xml":
                case "xlsx":
                    convertSpreadsheet(inFile, outFile);
                    break;
                case "ppt":
                case "pptx":
                    convertPresentation(inFile, outFile);
                    break;
                default:
                    throw new ArgumentException(String.Format("Input format [{0}] is not supported", inFormat));
            }
        }

        private void convertDocument(System.IO.FileInfo inFile, System.IO.FileInfo outFile)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;

            var document = app.Documents.Open(inFile.FullName);
            document.ExportAsFixedFormat(outFile.FullName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            document.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            app.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
        }

        private void convertPresentation(System.IO.FileInfo inFile, System.IO.FileInfo outFile)
        {
            var app = new Microsoft.Office.Interop.PowerPoint.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.PowerPoint.PpAlertLevel.ppAlertsNone;
            
            var presentation = app.Presentations.Open(inFile.FullName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            presentation.SaveAs(outFile.FullName, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
            presentation.Close();
            app.Quit();
        }

        private void convertSpreadsheet(System.IO.FileInfo inFile, System.IO.FileInfo outFile)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false;

            var workbook = app.Workbooks.Open(inFile.FullName);
            workbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outFile.FullName);
            workbook.Close(false);
            app.Quit();
        }
    }
}
