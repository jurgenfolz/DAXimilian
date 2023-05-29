using OfficeOpenXml;
using System.Data;
using System.IO;

public class ExcelREader
{
   
    public static DataTable ExcelToDataTable(string path)
    {
        using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];  // Assuming 1st worksheet
            DataTable dt = new DataTable();

            bool hasHeader = true; //false if the  excel file doesn't have a header
            foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }

            var startRow = hasHeader ? 2 : 1;

            for (var rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                var row = dt.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                dt.Rows.Add(row);
            }

            return dt;
        }
    }
}
