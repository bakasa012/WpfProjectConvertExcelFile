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

            set {instance = value; }
        }
        private SetExcelCellStyle() { }

        [System.Obsolete]
        public void FontChange(XSSFWorkbook wb, string caseFont, IRow row, int indexCell)
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
                    cellStyle.FillForegroundColor = HSSFColor.LightGreen.Index;
                    cellStyle.DataFormat = wb.CreateDataFormat().GetFormat("text");
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    break;
                case "bgBlack":
                    cellStyle.FillForegroundColor = HSSFColor.Grey50Percent.Index;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    //cellStyle.FillBackgroundColor = HSSFColor.Red.Index;
                    break;
                case "bgYellow":
                    cellStyle.FillForegroundColor = HSSFColor.LightYellow.Index;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
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
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    break;
            }
            cellStyle.SetFont(font);
            row.GetCell(indexCell).CellStyle = cellStyle;
        }
    }
}
