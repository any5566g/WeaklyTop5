using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Table
{
    public static partial class EPPlusTableExtensions
    {
        public static void EnableFilters(this ExcelTable tblData, int filterColumnIndex, IEnumerable<string> filterValues, bool blanks)
        {
            var ws = tblData.WorkSheet;
            var xdoc = tblData.TableXml;
            var nsm = new XmlNamespaceManager(xdoc.NameTable);

            // "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            var schemaMain = xdoc.DocumentElement.NamespaceURI;
            if (nsm.HasNamespace("x") == false)
                nsm.AddNamespace("x", schemaMain);

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.table.aspx
            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.autofilter.aspx
            var autoFilter = xdoc.SelectSingleNode("/x:table/x:autoFilter", nsm);
            if (autoFilter == null)
            {
                var table = xdoc.SelectSingleNode("/x:table", nsm);
                autoFilter = table.AppendElement(schemaMain, "x:autoFilter");
            }

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.filtercolumn.aspx
            var filterColumn = autoFilter.AppendElement(schemaMain, "x:filterColumn");
            int colId = filterColumnIndex - tblData.Address.Start.Column;
            filterColumn.AppendAttribute("colId", colId.ToString());

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.filters.aspx
            var filters = filterColumn.AppendElement(schemaMain, "x:filters");
            if (blanks)
                filters.AppendAttribute("blank", "true");

            foreach (var filterValue in filterValues)
            {
                // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.filter.aspx
                var filter = filters.AppendElement(schemaMain, "x:filter");
                filter.AppendAttribute("val", filterValue);
            }

            // hide rows
            var filterCells = ws.Cells[tblData.Address.Start.Row + 1, filterColumnIndex, tblData.Address.End.Row, filterColumnIndex];
            var cellValues = filterCells.Select(cell => new { Value = (cell.Value ?? string.Empty).ToString(), cell.Start.Row });
            var hiddenRows = cellValues.Where(c => filterValues.Contains(c.Value) == false).Select(c => c.Row);
            if (blanks)
                hiddenRows = hiddenRows.Except(cellValues.Where(c => string.IsNullOrEmpty(c.Value)).Select(c => c.Row));
            hiddenRows = hiddenRows.OrderByDescending(r => r);
            foreach (var row in hiddenRows)
                ws.Row(row).Hidden = true;
        }
    }
}