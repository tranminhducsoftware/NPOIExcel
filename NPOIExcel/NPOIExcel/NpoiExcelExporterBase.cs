using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOIExcel.Attributes;
using NPOIExcel.Constant;
using NPOIExcel.Dto;

namespace NPOIExcel
{
    public class NpoiExcelExporterBase
    {
        protected FileDto CreateExcelPackage(string fileName, Action<XSSFWorkbook> creator)
        {
            var file = new FileDto(fileName, MimeTypeNames.ApplicationVndOpenxmlformatsOfficedocumentSpreadsheetmlSheet);
            var workbook = new XSSFWorkbook();

            creator(workbook);
            Save(workbook, file);
            return file;
        }

        protected void AddHeader(ISheet sheet, params string[] headerTexts)
        {
            if (!headerTexts.Any())
            {
                return;
            }

            sheet.CreateRow(0);

            for (var i = 0; i < headerTexts.Length; i++)
            {
                AddHeader(sheet, i, headerTexts[i]);
            }
        }

        protected void AddHeader(ISheet sheet, params InvalidExportAttribute[] headers)
        {
            if (!headers.Any())
            {
                return;
            }

            var rowCount = headers.Max(o => o.RowIndex) + 1;
            var colCount = headers.Where(o => o.RowIndex == 0).Sum(o => o.ColSpan);

            // Init Header
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = sheet.CreateRow(rowIndex);
                row.Height = 600;
                var headersByRow = headers.Where(o => o.RowIndex == rowIndex).ToList();
                for (var colIndex = 0; colIndex < colCount; colIndex++)
                {
                    var attr = headersByRow.FirstOrDefault(o => o.ColIndex == colIndex);
                    if (attr == null && colIndex < headersByRow.Count)
                        attr = headersByRow[colIndex];

                    var cell = row.CreateCell(colIndex);
                    var cellStyle = sheet.Workbook.CreateCellStyle();
                    var font = sheet.Workbook.CreateFont();

                    font.IsBold = true;
                    font.FontHeightInPoints = 11;
                    font.Color = attr is {Required: true} ? HSSFColor.Red.Index : HSSFColor.Black.Index;
                    
                    cellStyle.SetFont(font);

                    cellStyle.BorderBottom = BorderStyle.Medium;
                    cellStyle.BorderRight = BorderStyle.Medium;
                    cellStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.WrapText = true;
                    cell.CellStyle = cellStyle;
                }
            }

            // Fill Text
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                var headersByRow = headers.Where(o => o.RowIndex == rowIndex).ToList();
                var cellColIndex = 0;
                foreach (var attr in headersByRow)
                {
                    var actualColIndex = attr.ColIndex == 0 ? cellColIndex : attr.ColIndex;
                    var cell = row.GetCell(actualColIndex);
                    cell.SetCellValue(attr.ColName);
                    sheet.SetColumnWidth(actualColIndex, attr.ColWidth * 100);
                    cellColIndex += attr.ColSpan;
                }
            }

            // Add Merge
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var headersByRow = headers.Where(o => o.RowIndex == rowIndex).ToList();
                var cellColIndex = 0;
                foreach (var attr in headersByRow)
                {
                    if (attr.RowSpan > 1)
                    {
                        var cra = new CellRangeAddress(rowIndex, rowIndex + attr.RowSpan - 1, cellColIndex,
                            cellColIndex);
                        sheet.AddMergedRegion(cra);
                    }

                    if (attr.ColSpan > 1)
                    {
                        var cra = new CellRangeAddress(rowIndex, rowIndex, cellColIndex,
                            cellColIndex + attr.ColSpan - 1);
                        sheet.AddMergedRegion(cra);
                    }

                    cellColIndex += attr.ColSpan;
                }
            }
        }
        
        private void AddHeader(ISheet sheet, int columnIndex, string headerText)
        {
            var cell = sheet.GetRow(0).CreateCell(columnIndex);
            cell.SetCellValue(headerText);
            var cellStyle = sheet.Workbook.CreateCellStyle();
            var font = sheet.Workbook.CreateFont();
            font.IsBold = true;
            font.FontHeightInPoints = 12;
            cellStyle.SetFont(font);
            cell.CellStyle = cellStyle;
        }

        protected void AddObjects<T>(ISheet sheet, IList<T> items, IList<string> propertiesName, int startRowIndex = 1)
        {
            if (!items.Any() || !propertiesName.Any()) return;

            var defaultCellStyle = sheet.Workbook.CreateCellStyle();
            defaultCellStyle.VerticalAlignment = VerticalAlignment.Center;

            for (var i = startRowIndex; i < items.Count + startRowIndex; i++)
            {
                var row = sheet.CreateRow(i);
                for (var j = 0; j < propertiesName.Count; j++)
                {
                    var cell = row.CreateCell(j);
                    cell.CellStyle = defaultCellStyle;
                    cell.CellStyle.WrapText = true;
                    object? value = null;
                    if (items[i - startRowIndex]?.GetType().GetProperty(propertiesName[j]) != null)
                    {
                        value = items[i - startRowIndex]?.GetType().GetProperty(propertiesName[j])?.GetValue(items[i - startRowIndex], null);
                    }

                    if (value != null)
                    {
                        var strValue = value.ToString()?.Replace("[", "").Replace("]", "");
                        cell.SetCellValue(strValue);
                    }
                }
            }
        }

        protected void AddObjects<T>(ISheet sheet, int startRowIndex, IList<T> items, params Func<T, object>[] propertySelectors)
        {
            if (!items.Any() || !propertySelectors.Any()) return;
            

            var defaultCellStyle = sheet.Workbook.CreateCellStyle();
            defaultCellStyle.VerticalAlignment = VerticalAlignment.Center;

            for (var i = 1; i <= items.Count; i++)
            {
                var row = sheet.CreateRow(i);

                for (var j = 0; j < propertySelectors.Length; j++)
                {
                    var cell = row.CreateCell(j);
                    cell.CellStyle = defaultCellStyle;
                    cell.CellStyle.WrapText = true;
                    var value = propertySelectors[j](items[i - 1]);
                    if (value != null) cell.SetCellValue(value.ToString());
                    
                }
            }
        }

        private void Save(XSSFWorkbook excelPackage, FileDto file)
        {
            using var stream = new MemoryStream();
            excelPackage.Write(stream);
        }

        protected void SetCellDataFormat(ICell? cell, string dataFormat)
        {
            if (cell == null) return;

            var dateStyle = cell.Sheet.Workbook.CreateCellStyle();
            var format = cell.Sheet.Workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat(dataFormat);
            cell.CellStyle = dateStyle;
            if (DateTime.TryParse(cell.StringCellValue, out var datetime))
            {
                cell.SetCellValue(datetime);
            }
        }
    }
}