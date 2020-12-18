using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    // https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.showdataasvalues.aspx
    public enum ShowDataAsValues
    {
        [StringValue("normal")]
        Normal,
        [StringValue("difference")]
        Difference,
        [StringValue("percent")]
        Percent,
        [StringValue("percentDiff")]
        PercentageDifference,
        [StringValue("runTotal")]
        RunTotal,
        [StringValue("percentOfRow")]
        PercentOfRow,
        [StringValue("percentOfCol")]
        PercentOfColumn,
        [StringValue("percentOfTotal")]
        PercentOfTotal,
        [StringValue("index")]
        Index
    }

    // https://msdn.microsoft.com/en-us/library/dd910980.aspx
    public enum PivotShowAsValues
    {
        [StringValue("percentOfParent")]
        PercentOfParent,
        [StringValue("percentOfParentRow")]
        PercentOfParentRow,
        [StringValue("percentOfParentCol")]
        PercentOfParentCol,
        [StringValue("percentOfRunningTotal")]
        PercentOfRunningTotal,
        [StringValue("rankAscending")]
        RankAscending,
        [StringValue("rankDescending")]
        RankDescending
    }

    public static partial class EPPlusPivotTableExtensions
    {
        public static void SortOnDataField(this ExcelPivotTable pivotTable, ExcelPivotTableField field, ExcelPivotTableDataField dataField, bool descending = false)
        {
            var xdoc = pivotTable.PivotTableXml;
            var nsm = new XmlNamespaceManager(xdoc.NameTable);

            // "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            var schemaMain = xdoc.DocumentElement.NamespaceURI;
            if (nsm.HasNamespace("x") == false)
                nsm.AddNamespace("x", schemaMain);

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pivotfield.aspx
            var pivotField = xdoc.SelectSingleNode("/x:pivotTableDefinition/x:pivotFields/x:pivotField[position()=" + (field.Index + 1) + "]", nsm);
            pivotField.AppendAttribute("sortType", (descending ? "descending" : "ascending"));

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.autosortscope.aspx
            var autoSortScope = pivotField.AppendElement(schemaMain, "x:autoSortScope");

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pivotarea.aspx
            var pivotArea = autoSortScope.AppendElement(schemaMain, "x:pivotArea");

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pivotareareferences.aspx
            var references = pivotArea.AppendElement(schemaMain, "x:references");
            references.AppendAttribute("count", "1");

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pivotareareference.aspx
            var reference = references.AppendElement(schemaMain, "x:reference");
            // Specifies the index of the field to which this filter refers. A value of -2 indicates the 'data' field.
            // int -> uint: -2 -> ((2^32)-2) = 4294967294
            reference.AppendAttribute("field", "4294967294");

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fielditem.aspx
            var x = reference.AppendElement(schemaMain, "x:x");
            int v = 0;
            foreach (ExcelPivotTableDataField pivotDataField in pivotTable.DataFields)
            {
                if (pivotDataField == dataField)
                {
                    x.AppendAttribute("v", v.ToString());
                    break;
                }
                v++;
            }
        }

        public static void Top10(this ExcelPivotTable pivotTable, ExcelPivotTableField field, ExcelPivotTableDataField dataField, int number = 10, bool bottom = false, bool percent = false)
        {
            var xdoc = pivotTable.PivotTableXml;
            var nsm = new XmlNamespaceManager(xdoc.NameTable);

            // "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            var schemaMain = xdoc.DocumentElement.NamespaceURI;
            if (nsm.HasNamespace("x") == false)
                nsm.AddNamespace("x", schemaMain);

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pivotfilters.aspx
            var filters = xdoc.SelectSingleNode("/x:pivotTableDefinition/x:filters", nsm);
            int filtersCount = 0;
            if (filters == null)
            {
                var pivotTableDefinition = xdoc.SelectSingleNode("/x:pivotTableDefinition", nsm);
                filters = pivotTableDefinition.AppendElement(schemaMain, "x:filters");
                filtersCount = 1;
            }
            else
            {
                XmlAttribute countAttr = filters.Attributes["count"];
                int count = int.Parse(countAttr.Value);
                filtersCount = count + 1;
            }

            filters.AppendAttribute("count", filtersCount.ToString());

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pivotfilter.aspx
            var filter = filters.AppendElement(schemaMain, "x:filter");
            filter.AppendAttribute("id", filtersCount.ToString());
            filter.AppendAttribute("type", (percent ? "percent" : "count"));

            int fld = 0;
            foreach (ExcelPivotTableField pivotField in pivotTable.Fields)
            {
                if (pivotField == field)
                {
                    filter.AppendAttribute("fld", fld.ToString());
                    break;
                }
                fld++;
            }

            int iMeasureFld = 0;
            foreach (ExcelPivotTableDataField pivotDataField in pivotTable.DataFields)
            {
                if (pivotDataField == dataField)
                {
                    filter.AppendAttribute("iMeasureFld", iMeasureFld.ToString());
                    break;
                }
                iMeasureFld++;
            }

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.autofilter.aspx
            var autoFilter = filter.AppendElement(schemaMain, "x:autoFilter");

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.filtercolumn.aspx
            var filterColumn = autoFilter.AppendElement(schemaMain, "x:filterColumn");
            filterColumn.AppendAttribute("colId", "0"); // the first auto filter in the pivot table

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.top10.aspx
            var top10 = filterColumn.AppendElement(schemaMain, "x:top10");
            top10.AppendAttribute("val", number.ToString());
            top10.AppendAttribute("top", (bottom ? "0" : "1"));
            top10.AppendAttribute("percent", (percent ? "1" : "0"));
        }

        public static void ShowValueAs(this ExcelPivotTable pivotTable, ExcelPivotTableDataField dataField, ShowDataAsValues showDataAs, ExcelPivotTableField baseField = null)
        {
            var xdoc = pivotTable.PivotTableXml;
            var nsm = new XmlNamespaceManager(xdoc.NameTable);

            // "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            var schemaMain = xdoc.DocumentElement.NamespaceURI;
            if (nsm.HasNamespace("x") == false)
                nsm.AddNamespace("x", schemaMain);

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.datafield.aspx
            var dataFieldNode = xdoc.SelectSingleNode("/x:pivotTableDefinition/x:dataFields/x:dataField[@name='" + dataField.Name + "']", nsm);
            dataFieldNode.AppendAttribute("showDataAs", showDataAs.StringValue());

            if (baseField != null)
                dataFieldNode.AppendAttribute("baseField", baseField.Index.ToString());
        }

        public static void ShowValueAs(this ExcelPivotTable pivotTable, ExcelPivotTableDataField dataField, PivotShowAsValues pivotShowAs, ExcelPivotTableField baseField = null)
        {
            var xdoc = pivotTable.PivotTableXml;
            var nsm = new XmlNamespaceManager(xdoc.NameTable);

            // "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            var schemaMain = xdoc.DocumentElement.NamespaceURI;
            if (nsm.HasNamespace("x") == false)
                nsm.AddNamespace("x", schemaMain);

            var schemaMainX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
            if (nsm.HasNamespace("x14") == false)
                nsm.AddNamespace("x14", schemaMainX14);

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.pivottabledefinition.aspx
            var pivotTableDefinition = xdoc.SelectSingleNode("/x:pivotTableDefinition", nsm);
            pivotTableDefinition.AppendAttribute("updatedVersion", "5");

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.datafield.aspx
            var dataFieldNode = xdoc.SelectSingleNode("/x:pivotTableDefinition/x:dataFields/x:dataField[@name='" + dataField.Name + "']", nsm);

            if (baseField != null)
                dataFieldNode.AppendAttribute("baseField", baseField.Index.ToString());

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.datafieldextensionlist.aspx
            var extLst = dataFieldNode.AppendElement(schemaMain, "x:extLst");

            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.datafieldextension.aspx
            var ext = extLst.AppendElement(schemaMain, "x:ext");

            // https://msdn.microsoft.com/en-us/library/dd950685.aspx
            ext.AppendAttribute("uri", "{E15A36E0-9728-4e99-A89B-3F7291B0FE68}");

            // https://msdn.microsoft.com/en-us/library/dd949774.aspx
            var x14DataField = ext.AppendElement(schemaMainX14, "x14:dataField");
            x14DataField.AppendAttribute("pivotShowAs", pivotShowAs.StringValue());
        }
    }
}