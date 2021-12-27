using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModuleConvertExcelFile.NPOI
{
    public class SetExcelColumnStyle
    {
        public void SetColumnStyle(IWorkbook wb, ISheet sheet)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.FillForegroundColor = HSSFColor.White.Index;
            cellStyle.FillPattern = FillPattern.LeastDots;
            cellStyle.BorderBottom = BorderStyle.Thick;
            cellStyle.BottomBorderColor = HSSFColor.White.Index;
            for (int i = 0; i < 200; i++)
            {
                sheet.SetDefaultColumnStyle(i, cellStyle);
            }
        }
    }
}
