namespace Npoiextend
{
    using NPOI.SS.UserModel;
    
    using Cell = NPOI.SS.UserModel.ICell;
    using Sheet = NPOI.SS.UserModel.ISheet;
    using XlsWorkbook = NPOI.HSSF.UserModel.HSSFWorkbook;
    using XlsxWorkbook = NPOI.XSSF.UserModel.XSSFWorkbook;

    public static class NPOIExtend
    {
        public static Cell Cell(this Sheet sheet, int row, int col)
        {
            var targetRow = sheet.GetRow(row - 1) ?? sheet.CreateRow(row - 1);
            return targetRow.GetCell(col - 1) ?? targetRow.CreateCell(col - 1);
        }

        public static dynamic Get(this Sheet sheet, int row, int col)
        {
            Cell cell = sheet.Cell(row, col);
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                default:
                    return "null";
            }
        }

        public static void Set(this Sheet sheet, int row, int col, object value)
        {
            Cell cell = sheet.Cell(row, col);
            if (value == null)
                cell.SetCellValue("");
            else if (value is string)
                cell.SetCellValue((string)value);
            else if (value is double || value is int || value is float || value is long || value is decimal)
                cell.SetCellValue(Convert.ToDouble(value));
            else if (value is DateTime)
                cell.SetCellValue((DateTime)value);
            else if (value is bool)
                cell.SetCellValue((bool)value);
            else
                cell.SetCellValue(value.ToString());
        }

        public static Sheet Sheet(this XlsWorkbook workbook, int sheetIndex)
         => workbook.GetSheetAt(sheetIndex - 1) ?? workbook.CreateSheet();
        public static Sheet Sheet(this XlsxWorkbook workbook, int sheetIndex)
        => workbook.GetSheetAt(sheetIndex - 1) ?? workbook.CreateSheet();

    }
}
