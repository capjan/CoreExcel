using NPOI.SS.UserModel;

namespace CoreExcel
{
    public class ExcelCell
    {
        private readonly ICell _wrappedCell;

        public string StringValue => _wrappedCell.StringCellValue;
        
        public ExcelCell(ICell wrappedCell)
        {
            _wrappedCell = wrappedCell;
        }
    }
}