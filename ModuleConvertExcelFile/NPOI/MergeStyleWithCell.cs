using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModuleConvertExcelFile.NPOI
{
    public class MergeStyleWithCell
    {
        private static MergeStyleWithCell instance;

        public static MergeStyleWithCell Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new MergeStyleWithCell();
                }

                return instance;
            }

            set { instance = value; }
        }
        private MergeStyleWithCell() { }

        [Obsolete]
        public void CreateCellStyleTitle(XSSFWorkbook wb,IRow row,int indexRow,string caseStr = "title")
        {
            SetExcelCellStyle.Instance.FontChange(wb, caseStr, row, indexRow);
        }

        [Obsolete]
        public void SetCellStyleOneRow(XSSFWorkbook wb,ISheet sheet, int indexRow, int startCell, int endCell, string strCase)
        {
            for (int i = startCell; i <= endCell; i++)
            {
                SetExcelCellStyle.Instance.FontChange(wb, strCase, sheet.GetRow(indexRow), i);
            }
        }

        [Obsolete]
        public void SetCellStyleMultiRow(XSSFWorkbook wb, ISheet sheet, int startRow,int endRow, int startCell, int endCell, string strCase)
        {
            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startCell; j <= endCell; j++)
                {
                    SetExcelCellStyle.Instance.FontChange(wb, strCase, sheet.GetRow(i), j);
                }
            }
        }
    }
}
