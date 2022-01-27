using NPOI.SS.UserModel;

namespace CoreExcel
{
    public class ExcelCell
    {
        private readonly ICell _wrappedCell;

        public string StringValue => _wrappedCell.StringCellValue;

        public int ColumnIndex => _wrappedCell.ColumnIndex;
        
        public ExcelCell(ICell wrappedCell)
        {
            _wrappedCell = wrappedCell;
        }
    }
}