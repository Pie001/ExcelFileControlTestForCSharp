using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelFileControlTest.logic;
using ExcelFileControlTest.Model;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            List<SimpleExcelModel> simpleExcelModelList = new List<SimpleExcelModel>();

            // 엑셀파일 읽기
            ReadExcelFile readExcelFile = new ReadExcelFile();
            simpleExcelModelList = readExcelFile.ReadExcelFileValues();

            // 위에서 읽은 엑셀파일의 텍스트 값을 가지고 그대로 새 엑셀파일을 만들어서 뿌려주기 
            MakeExcelFiles makeExcelFiles = new MakeExcelFiles();
            makeExcelFiles.MakeDetailsAsExcel(simpleExcelModelList);
        }
    }
}
