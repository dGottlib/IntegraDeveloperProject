using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
using System.Data;

namespace IntegraDevProject
{
    public static class DataTabletoExcel
    {
        public static void ExporttabletoExcel(DataTable dt, string filename, List<string> allowed)
        {
            for (int i = dt.Columns.Count - 1; i >= 0; i--)
            {
                DataColumn column = dt.Columns[i];

                if (!allowed.Contains(column.ColumnName))
                {
                    if (column.ColumnName == "ActiveDate")
                    {
                        column.ColumnName = "Live";
                    }
                    else
                    {
                        dt.Columns.Remove(column);
                    }
                }
            }

            var workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add(dt, "Sheet 1");

            workbook.SaveAs(filename);
        }

    }
}

