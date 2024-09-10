using SpreadCheetah;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zeyt.ExcelDocument
{
    public class ExcelDocumentWriter<TClass, TClassMap> where TClass : class where TClassMap : ExcelDocumentMap<TClass>
    {
        private readonly List<ExcelColumnMap<TClass>> _excelColumnMapList;
        private readonly ExcelDocumentStyle _excelDocumentStyle;
        private readonly string _sheetName;

        private StyleId? _headerStyleId;
        private StyleId? _zebraFirstRowStyleId;
        private StyleId? _zebraSecondRowStyleId;

        public ExcelDocumentWriter(string sheetName, ExcelDocumentStyle? excelDocumentStyle = null)
        {
            _sheetName = string.IsNullOrWhiteSpace(sheetName) ? "SheetName1" : sheetName;
            _excelDocumentStyle = excelDocumentStyle ?? new ExcelDocumentStyle();
            _excelColumnMapList = Activator.CreateInstance<TClassMap>().ExcelColumnMapList;
        }

        public byte[] Write(List<TClass> recordList)
        {
            using (var stream = new MemoryStream())
            {
                using (var spreadSheet = Spreadsheet.CreateNewAsync(stream).ConfigureAwait(false).GetAwaiter().GetResult())
                {
                    spreadSheet.StartWorksheetAsync(_sheetName, WriteOptions()).ConfigureAwait(false).GetAwaiter().GetResult();

                    WriteStyles(spreadSheet, _excelDocumentStyle);
                    WriteHeader(spreadSheet, _excelColumnMapList);
                    WriteRecordList(spreadSheet, recordList);

                    spreadSheet.FinishAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                }
                return stream.ToArray();
            }
        }
        private WorksheetOptions WriteOptions()
        {
            var worksheetOptions = new WorksheetOptions();

            for (int i = 0; i < _excelColumnMapList.Count; i++)
            {
                if (_excelColumnMapList[i]?.ExcelColumnMapData.Width > 255)
                {
                    worksheetOptions.Column(i + 1).Width = 255;
                }
                else if (_excelColumnMapList[i]?.ExcelColumnMapData.Width == 0)
                {
                    worksheetOptions.Column(i + 1).Width = 10;
                }
                else
                {
                    worksheetOptions.Column(i + 1).Width = _excelColumnMapList[i]?.ExcelColumnMapData?.Width;
                }
            }

            return worksheetOptions;
        }

        private void WriteStyles(Spreadsheet spreadsheet, ExcelDocumentStyle? excelDocumentStyle = null)
        {
            if (excelDocumentStyle != null && excelDocumentStyle?.HeaderStyle != null)
            {
                _headerStyleId = spreadsheet.AddStyle(excelDocumentStyle.HeaderStyle);
            }
            else
            {
                var headerStyle = new Style();
                headerStyle.Font.Size = 12;
                headerStyle.Font.Bold = true;
                headerStyle.Font.Color = Color.FromArgb(255, 255, 255);
                headerStyle.Fill.Color = Color.FromArgb(15, 158, 213);
                _headerStyleId = spreadsheet.AddStyle(headerStyle);
            }

            if (excelDocumentStyle != null && excelDocumentStyle?.ZebraStyle == true)
            {
                var zebraFirstRowStyle = new Style();
                zebraFirstRowStyle.Font.Size = 11;
                zebraFirstRowStyle.Font.Color = Color.FromArgb(0, 0, 0);
                zebraFirstRowStyle.Fill.Color = Color.FromArgb(202, 237, 251);
                _zebraFirstRowStyleId = spreadsheet.AddStyle(zebraFirstRowStyle);

                var zebraSecondRowStyle = new Style();
                zebraSecondRowStyle.Font.Size = 11;
                zebraSecondRowStyle.Font.Color = Color.FromArgb(0, 0, 0);
                zebraSecondRowStyle.Fill.Color = Color.FromArgb(255, 255, 255);
                _zebraSecondRowStyleId = spreadsheet.AddStyle(zebraSecondRowStyle);
            }
        }

        private void WriteHeader(Spreadsheet spreadsheet, List<ExcelColumnMap<TClass>> excelColumnMapList)
        {
            var cellColumnList = new List<Cell>();

            foreach (var excelColumnMap in excelColumnMapList)
            {
                cellColumnList.Add(new Cell(excelColumnMap.ExcelColumnMapData.Name, _headerStyleId));
            }

            spreadsheet.AddRowAsync(cellColumnList).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        private void WriteRecordList(Spreadsheet spreadsheet, List<TClass> recordList)
        {
            int index = 0;

            foreach (var record in recordList)
            {
                if (index % 2 == 0 && _zebraFirstRowStyleId != null)
                {
                    WriteRecord(spreadsheet, record, _zebraFirstRowStyleId);
                }
                else if (index % 2 != 0 && _zebraSecondRowStyleId != null)
                {
                    WriteRecord(spreadsheet, record, _zebraSecondRowStyleId);
                }
                else
                {
                    WriteRecord(spreadsheet, record);
                }
                index++;
            }
        }

        private void WriteRecord(Spreadsheet spreadsheet, TClass record, StyleId? styleId = null)
        {
            var cellList = new List<Cell>();

            foreach (var excelColumnMap in _excelColumnMapList)
            {
                if (styleId != null)
                {
                    cellList.Add(new Cell(ResolveValue(excelColumnMap, record), styleId));
                }
                else
                {
                    cellList.Add(new Cell(ResolveValue(excelColumnMap, record)));
                }
            }

            spreadsheet.AddRowAsync(cellList).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        private string? ResolveValue(ExcelColumnMap<TClass> excelColumnMap, TClass record)
        {
            var value = excelColumnMap?.ExcelColumnMapData?.Property?.GetValue(record)?.ToString();
            var defaultValue = excelColumnMap?.ExcelColumnMapData?.Default?.ToString();

            if (string.IsNullOrWhiteSpace(value) && string.IsNullOrWhiteSpace(defaultValue))
            {
                return string.Empty;
            }
            else if (excelColumnMap?.ExcelColumnMapData?.WriteUsing != null)
            {
                return excelColumnMap?.ExcelColumnMapData?.WriteUsing(record)?.ToString();
            }
            else if (string.IsNullOrWhiteSpace(value) && !string.IsNullOrWhiteSpace(defaultValue))
            {
                return defaultValue;
            }
            else
            {
                return value;
            }
        }
    }
}
