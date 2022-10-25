using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CoreExcel
{
    public class ExcelSheet
    {
        private readonly ISheet _wrappedSheet;

        public string Name => _wrappedSheet.SheetName;
        public int FirstRowNumber => _wrappedSheet.FirstRowNum;
        public int LastRowNumber => _wrappedSheet.LastRowNum;

        public IEnumerable<ExcelRow> Rows
        {
            get
            {
                foreach (XSSFRow row in _wrappedSheet)
                {
                    if (row.Cells.Count == 0) continue;
                    yield return new ExcelRow(row);
                }
            }
        }

        public ExcelSheet(ISheet wrappedSheet)
        {
            _wrappedSheet = wrappedSheet;
        }
    }
}