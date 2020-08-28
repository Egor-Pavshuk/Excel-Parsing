using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;


namespace Excel_Parsing
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            FileInfo file = new FileInfo(@"C:\Users\egorp\Downloads\Telegram Desktop\ПОВЗВОДНИЙ.xls");
            Console.WriteLine(file.OpenRead().CanRead);

            if (file.Extension.ToString() == ".xls")
            {
                var documet = new Aspose.Cells.Workbook(file.FullName);
                documet.Save("Output.xlsx", Aspose.Cells.SaveFormat.Xlsx);
                file = new FileInfo("Output.xlsx");
            }

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                try
                {
                    excelPackage.Workbook.Worksheets.Delete("Evaluation Warning");
                }
                finally
                {
                    var worckSheet = excelPackage.Workbook.Worksheets.First();
                    var nail = worckSheet.Cells;

                    for (int i = 1; i <= worckSheet.Dimension.Rows; i++)
                    {
                        for (int j = 1; j <= worckSheet.Dimension.Columns; j++)
                        {
                            Console.Write(nail[i, j].Value?.ToString().Replace('і', 'i').Replace('І', 'I') + " ");
                        }
                        Console.WriteLine();
                    }
                }
            }
            file.Delete();
            Console.ReadKey();
        }
    }
}
