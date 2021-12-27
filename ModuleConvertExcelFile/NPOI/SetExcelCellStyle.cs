using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ModuleConvertExcelFile.NPOI
{
    public class SetExcelCellStyle
    {
        private static SetExcelCellStyle instance;

        public static SetExcelCellStyle Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new SetExcelCellStyle();
                }

                return instance;
            }

            set { SetExcelCellStyle.instance = value; }
        }

#pragma warning disable SA1201 // Elements should appear in the correct order
        private SetExcelCellStyle() { }

#pragma warning restore SA1201 // Elements should appear in the correct order

        public void FontChange(XSSFWorkbook wb, string caseFont, IRow row, int index)
        {
            IFont font = wb.CreateFont();
            ICellStyle cellStyle = wb.CreateCellStyle();

            switch (caseFont)
            {
                case "title":
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 16;
                    break;
                case "label":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    cellStyle.FillForegroundColor = HSSFColor.Gold.Index;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    break;
                case "content":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    break;
                case "table":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    cellStyle.DataFormat = wb.CreateDataFormat().GetFormat("text");
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    break;
                case "description":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    break;
                case "numberic":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    cellStyle.DataFormat = wb.CreateDataFormat().GetFormat("text");
                    break;
                case "bgBlack":
                    /*cellStyle.BorderBottom = BorderStyle.None;
                    cellStyle.BorderLeft = BorderStyle.None;
                    cellStyle.BorderRight = BorderStyle.None;
                    cellStyle.BorderTop = BorderStyle.None;*/
                    cellStyle.FillForegroundColor = HSSFColor.Grey50Percent.Index;
                    cellStyle.FillBackgroundColor = HSSFColor.Red.Index;
                    break;
                case "tableBody":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    cellStyle.DataFormat = wb.CreateDataFormat().GetFormat("text");
                    cellStyle.Alignment = HorizontalAlignment.Left;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    break;
                default:
                    font.FontName = "Calibri";
                    cellStyle.FillBackgroundColor = HSSFColor.Gold.Index;
                    break;
            }
            cellStyle.SetFont(font);
            row.GetCell(index).CellStyle = cellStyle;
        }
    }
}
