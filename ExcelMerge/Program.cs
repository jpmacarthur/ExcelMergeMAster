using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            var dir = GetExecutingDirectory();
            var files = Directory.GetFiles(dir.FullName).Where(x => x.EndsWith("csv"));
            var filesInfo = new List<FileInfo>();
            var firstFile = new FileInfo(dir.FullName + "/Merged Observations.xlsx");
            var processedDir = dir.CreateSubdirectory("ProcessedFiles");


            if (files.Count() == 0)
            {
                Console.WriteLine("No Observations Found.");
                Console.ReadKey();
                return;
            }
            using (var mergeExcel = new ExcelPackage(firstFile))
            {
                try
                {
                    mergeExcel.Workbook.Worksheets.Add("Merged Observations");
                }
                catch (Exception ex)
                {
                    var isInvalid = true;
                    while (isInvalid)
                    {
                        Console.WriteLine("Previous Merge already exists. Overwrite? (Y/N)");
                        var response = Console.ReadLine();
                        if (response == "Y" || response == "y")
                        {
                            mergeExcel.Workbook.Worksheets.Delete("Merged Observations");
                            mergeExcel.Workbook.Worksheets.Add("Merged Observations");

                            isInvalid = false;
                        }
                        if (response == "N" || response == "n")
                        {
                            return;
                        }
                    }
                }
                var mergeWs = mergeExcel.Workbook.Worksheets["Merged Observations"];
                var newRow = GetNextRow(mergeExcel);
                foreach (var file in files)
                {
                    //create a WorkSheet

                    //load the CSV data into cell A1
                    var fi = new FileInfo(file);
                    filesInfo.Add(fi);
                    using (var p = new ExcelPackage())
                    {
                        ExcelWorksheet ws = p.Workbook.Worksheets.Add("Sheet 1");

                        var format = new ExcelTextFormat();
                        format.TextQualifier = '"';
                        format.SkipLinesBeginning = 1;

                        format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
                        format.Culture.DateTimeFormat.ShortDatePattern = "M/d/yyyy H:mm";

                        try
                        {
                            ws.Cells["A1"].LoadFromText(fi, format);

                        }
                        catch (Exception)
                        {
                            Console.WriteLine("Observation file in use. Please close all files and try again.");
                            Console.ReadKey();
                            return;
                        }


                        int iColCnt = ws.Dimension.End.Column;
                        int iRowCnt = ws.Dimension.End.Row;
                        for (var i = 0; i <= iRowCnt; i++)
                        {
                            var cells = ws.Cells[i + 1, 1, i + 1, iColCnt];
                            if (((object[,])cells.Value)[0, 0] != null)
                            {
                                cells.Copy(newRow);
                            }
                            newRow = GetNextRow(mergeExcel);
                        }
                        //Get the Worksheet created in the previous codesample. 
                        //Set the cell value using row and column.
                        //Save and close the package.

                    }
                }
                mergeWs.Cells[1, 1, mergeWs.Dimension.End.Row, 1].Style.Numberformat.Format = "M/d/yyyy H:mm";

                try
                {
                    mergeExcel.Save();
                    foreach (var file in filesInfo)
                    {
                        File.Move(file.FullName, dir.FullName + "\\ProcessedFiles\\" + file.Name);
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine("Error moving files to processed folder. Save Successful");
                    Console.ReadKey();
                }
                catch (InvalidOperationException ex)
                {
                    var isInvalid = true;
                    while (isInvalid)
                    {
                        Console.WriteLine("Error. If the merge file is open, please close it.");
                        Console.ReadKey();
                        if (!Utils.IsFileLocked(firstFile))
                        {
                            mergeExcel.Save();
                            isInvalid = false;
                        }
                        else
                        {
                            Console.WriteLine("Press any key to try again");
                        }

                    }
                }
                //Open the workbook (or create it if it doesn't exist)
            }

        }
        private static DirectoryInfo GetExecutingDirectory()
        {
            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            return new FileInfo(location.AbsolutePath).Directory;
        }
        private static ExcelRangeBase GetNextRow(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets["Merged Observations"];
            int iColCnt;
            int iRowCnt;
            try
            {
                iColCnt = ws.Dimension.End.Column;
                iRowCnt = ws.Dimension.End.Row;
            }
            catch (NullReferenceException)
            {
                iColCnt = 1;
                iRowCnt = 0;
            }

            var range = ws.Cells[iRowCnt + 1, 1, iRowCnt + 1, iColCnt];
            return range;
        }
    }
    public class Utils
    {
        static DirectoryInfo _outputDir = null;
        public static DirectoryInfo OutputDir
        {
            get
            {
                return _outputDir;
            }
            set
            {
                _outputDir = value;
                if (!_outputDir.Exists)
                {
                    _outputDir.Create();
                }
            }
        }
        public static FileInfo GetFileInfo(string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(OutputDir.FullName + Path.DirectorySeparatorChar + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }
        public static FileInfo GetFileInfo(DirectoryInfo altOutputDir, string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(altOutputDir.FullName + Path.DirectorySeparatorChar + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }

        internal static DirectoryInfo GetDirectoryInfo(string directory)
        {
            var di = new DirectoryInfo(_outputDir.FullName + Path.DirectorySeparatorChar + directory);
            if (!di.Exists)
            {
                di.Create();
            }
            return di;
        }
        public static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }
}
