using NPOI.SS.UserModel;

namespace CoreExcel
{
    public class ExcelCell
    {
        private readonly ICell _wrappedCell;
        private readonly DataFormatter _dataFormatter = new DataFormatter();

        public string StringValue => _dataFormatter.FormatCellValue(_wrappedCell);
        
        public int ColumnIndex => _wrappedCell.ColumnIndex;
        
        public ExcelCell(ICell wrappedCell)
        {
            _wrappedCell = wrappedCell;
        }
    }
}