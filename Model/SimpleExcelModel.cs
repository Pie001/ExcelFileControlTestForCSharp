using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileControlTest.Model
{
    public class SimpleExcelModel
    {
        public SimpleExcelModel()
        {
            Column = new Dictionary<int, string>();
        }

        public Dictionary<int, string> Column;
    }
}
