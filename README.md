⚠️ This project is archived and no longer maintained

CoreExcel was a convenience wrapper around (https://github.com/nissl-lab/npoi) that aimed to make reading Excel files as simple as possible.
It served its purpose, but the ecosystem has moved on, and maintaining a thin wrapper no longer makes sense.

## Recommended alternatives

If you're starting a new project, consider one of these instead:

- [ClosedXML](https://github.com/ClosedXML/ClosedXML): The easiest high-level API for .xlsx files. MIT-licensed, actively maintained, and the de facto standard for readable Excel code in .NET. ws.Cell("A3").Value = "Product" is about as simple as it gets.
- [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader): If you only need to read data, this is the fastest and lightest option, with streaming support for large files.
- [MiniExcel](https://github.com/mini-software/MiniExcel): Streaming-based, minimal memory footprint, ideal for large datasets and export scenarios.

## Migration

If you were already using CoreExcel, switching is straightforward: replace the wrapper calls with the equivalent ClosedXML calls. The concepts map one-to-one (workbook → sheet → row → cell), so most migrations are mechanical rather than conceptual.

## License

CoreExcel remains available under the (./LICENSE) for anyone who still depends on it. The NuGet package stays published, but no new versions will be released.

## Info

* Supports .NET Standard 2.0, .NET 6
* Depends on  [NPOI](https://github.com/nissl-lab/npoi) and [CoreLib](https://github.com/capjan/CoreLib).
