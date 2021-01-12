using System;
using System.Collections.Generic;
using System.IO;
using Core.IO.Impl;
using NPOI.XSSF.UserModel;

namespace CoreExcel
{
    public class ExcelDocument: IDisposable
    {
        private readonly XSSFWorkbook _workbook;
        private readonly bool _isShadowCopy;
        private readonly string _filePath;

        public IEnumerable<ExcelSheet> Sheets
        {
            get
            {
                foreach (var sheet in _workbook)
                    yield return new ExcelSheet(sheet);
            }
        }
        
        public ExcelDocument(string filePath, bool shadowCopy = true)
        {
            _isShadowCopy = shadowCopy;
            _filePath = InitFilePath(filePath);

            var fi = new FileInfo(_filePath);
            _workbook = new XSSFWorkbook(fi);
        }

        public ExcelDocument(Stream stream)
        {
            _isShadowCopy        = false;
            _filePath            = null;
            _workbook            = new XSSFWorkbook(stream);
        }

        private string InitFilePath(string filePath)
        {
            if (!_isShadowCopy) return filePath;
            var gen = new DefaultPathNameGenerator(postfix: ".xlsx");
            var util = new DefaultTempUtil(fileNameGen: gen);
            var result = util.CreateFile();

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var ws = new FileStream(result, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var read = 0L;
                    var buffer = new byte[4096];
                    while (read < fs.Length)
                    {
                        var currentRead = fs.Read(buffer, 0, buffer.Length);
                        read += currentRead;
                        if (currentRead != 0)
                        {
                            ws.Write(buffer, 0, currentRead);
                        }
                    }
                }
            }
            
            return result;
        }

        public void Dispose() {
            _workbook.Close();
            if (_isShadowCopy)
            {
                File.Delete(_filePath);
            }
        }
    }
}
