using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using ExcelFileControlTest.Model;

namespace ExcelFileControlTest.logic
{
    public class MakeExcelFiles
    {
        /// <summary>
        /// 엑셀파일 작성
        /// </summary>
        /// <param name="excelModel"></param>
        /// <returns>byte[]</returns>
        public void MakeDetailsAsExcel(List<SimpleExcelModel> list)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Column(1).Width = 3;
                var allRangeHeight = 500;

                var allRange = worksheet.Cells[1, 1, allRangeHeight, 32];
                allRange.Style.Font.Name = "굴림";
                allRange.Style.Font.Size = 11;
                allRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                int rows = list.Count;
                int columns = list[0].Column.Count;

                for (int i = 1; i <= rows; i++)
                {
                    for(int j = 1; j <= columns; j++)
                    {
                        worksheet.Cells[i,j].Value = list[i-1].Column[j];
                    }
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                var localPath = System.AppDomain.CurrentDomain.BaseDirectory;
                var FullPath = localPath + string.Format("File_{0}.xlsx", DateTime.Now.ToString("yyyyMMddhhmmss"));
                File.WriteAllBytes(FullPath, package.GetAsByteArray());
            }
        }
    }
}
