using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using ExcelFileControlTest.Model;

namespace ExcelFileControlTest.logic
{
    public class ReadExcelFile
    {

        public List<SimpleExcelModel> ReadExcelFileValues()
        {
            var localPath = System.AppDomain.CurrentDomain.BaseDirectory;
            string path = localPath + "//temp";
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            FileInfo[] Files = directoryInfo.GetFiles("*.xlsx");

            List<SimpleExcelModel> simpleExcelModelList = new List<SimpleExcelModel>();

            if (Files.Count() == 1)
            {              
                ExcelPackage package = new ExcelPackage(Files[0]);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                SimpleExcelModel simpleExcelModel;

                for (int i = 1; i <= rows; i++)
                {
                    simpleExcelModel = new SimpleExcelModel();
                    Dictionary<int, string> column = new Dictionary<int, string>();
                    for (int j = 1; j <= columns; j++)
                    {
                        string content = worksheet.Cells[i, j].Value?.ToString();
                        column.Add(j, content);
                    }
                    simpleExcelModel.Column = column;
                    simpleExcelModelList.Add(simpleExcelModel);
                }
            }

            return simpleExcelModelList;
        }
    }
}
