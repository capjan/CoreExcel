using System.Collections.Generic;
using System.Linq;
using Core.Converters;

namespace CoreExcel.Util
{
    public class SpecificRowReader<T>
    {
        private readonly IConverter<ExcelRow, T> _converter;
        private readonly int _skipRows;
        private readonly ExcelSheet _excelSheet;
        /// <summary>
        /// Creates an instance for the given arguments. 
        /// </summary>
        /// <param name="excelSheet">the sheet that needs to be converted</param>
        /// <param name="converter">specific converter to create an output of type T for given ExcelRows </param>
        /// <param name="skipRows">to skip the headline</param>
        public SpecificRowReader(ExcelSheet excelSheet, IConverter<ExcelRow, T> converter, int skipRows = 1)
        {
            _excelSheet = excelSheet;
            _converter = converter;
            _skipRows = skipRows;
        }

        /// <summary>
        /// Returns an iterator of the expected type T.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<T> ReadAll()
        {
            return _excelSheet.Rows.Skip(_skipRows).Select(x => _converter.Convert(x));
        }
    }
}