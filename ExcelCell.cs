using NPOI.SS.UserModel;

namespace ExcelQuery
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