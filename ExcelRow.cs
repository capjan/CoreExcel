using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace CoreExcel
{
    public class ExcelRow
    {
        private readonly IRow _wrappedRow;
        public ExcelRow(IRow wrappedRow)
        {
            _wrappedRow = wrappedRow;
        }

        public IEnumerable<ExcelCell> Cells => _wrappedRow.Cells.Select(cell => new ExcelCell(cell)); 
    }
}