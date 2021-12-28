using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleConvertExcelFile.NPOI
{
    public class ExportExcelFile
    {
        private static ExportExcelFile instance;

        public static ExportExcelFile Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ExportExcelFile();
                }

                return instance;
            }

            set { ExportExcelFile.instance = value; }
        }

        private ExportExcelFile() { }

        [Obsolete]
        public void ExportExcelFileWithNPOI(string @output, List<DataBinding.DataHeaderExcel> dataHeaderExcels, List<DataBinding.DataBodyExcelFile> dataBodyExcelFiles)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "sheet123");
            SetExcelColumnStyle.Instance.SetColumnStyle(wb, sheet);
            ICellStyle style = wb.CreateCellStyle();
            IRow row = sheet.CreateRow(0);
            row.CreateCell(1);
            //Merge column
            CellRangeAddress cellRange = new CellRangeAddress(0, 0, 1, 7);
            sheet.AddMergedRegion(cellRange);
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 5, 7));
            sheet.AddMergedRegion(new CellRangeAddress(3, 3, 5, 7));
            sheet.AddMergedRegion(new CellRangeAddress(4, 4, 3, 7));
            sheet.AddMergedRegion(new CellRangeAddress(5, 5, 3, 7));
            sheet.AddMergedRegion(new CellRangeAddress(6, 6, 3, 4));
            sheet.AddMergedRegion(new CellRangeAddress(6, 6, 5, 7));
            sheet.AddMergedRegion(new CellRangeAddress(7, 7, 4, 7));
            //set column width
            sheet.AutoSizeColumn(0);
            sheet.SetColumnWidth(1, 4000);
            sheet.SetColumnWidth(2, 6000);
            sheet.SetColumnWidth(3, 4000);
            sheet.SetColumnWidth(4, 3000);
            sheet.SetColumnWidth(5, 3000);
            sheet.SetColumnWidth(6, 3000);
            sheet.SetColumnWidth(7, 3000);
            row.GetCell(1).SetCellValue(dataHeaderExcels[0].column2);
            SetExcelCellStyle.Instance.FontChange(wb, "title", row, 1);
            int rowIndex = 1;
            dataHeaderExcels.RemoveAt(0);
            //data header
            foreach (DataBinding.DataHeaderExcel item in dataHeaderExcels)
            {
                IRow newRow = sheet.CreateRow(rowIndex);
                newRow.CreateCell(0).SetCellValue(item.column1);
                newRow.CreateCell(1).SetCellValue(item.column2);
                newRow.CreateCell(2).SetCellValue(item.column3);
                newRow.CreateCell(3).SetCellValue(item.column4);
                newRow.CreateCell(4).SetCellValue(item.column5);
                newRow.CreateCell(5).SetCellValue(item.column6);
                newRow.CreateCell(6).SetCellValue(item.column7);
                newRow.CreateCell(7).SetCellValue(item.column8);
                rowIndex++;
            }
            //set style for cell Header
            MergeStyleWithCell.Instance.SetCellStyleMultiRow(wb, sheet, 2, 7, 1, 1, "label");
            MergeStyleWithCell.Instance.SetCellStyleMultiRow(wb, sheet, 2, 7, 2, 7, "content");
            MergeStyleWithCell.Instance.SetCellStyleOneRow(wb, sheet, 5, 2, 2, "bgYellow");
            MergeStyleWithCell.Instance.SetCellStyleMultiRow(wb, sheet, 2, 3, 3, 3, "label");
            MergeStyleWithCell.Instance.SetCellStyleMultiRow(wb, sheet, 6, 7, 3, 3, "label");
            MergeStyleWithCell.Instance.SetCellStyleMultiRow(wb, sheet, 8, 9, 1, 7, "description");
            MergeStyleWithCell.Instance.SetCellStyleOneRow(wb, sheet, 3, 5, 7, "bgBlack");
            MergeStyleWithCell.Instance.SetCellStyleOneRow(wb, sheet, 4, 3, 7, "bgBlack");
            //set type value for cell header
<<<<<<< Updated upstream
            SetCellTypeValueForMultiRow.Instance.SetCellValueTypeMultiRow(sheet, 8, 9, 2, 7, CellType.Blank);
=======
            //SetCellValueType.Instance.SetCellValueTypeMultiRow(sheet, 8, 9, 2,7, CellType.Blank);
>>>>>>> Stashed changes
            //data body
            rowIndex++;
            foreach (DataBinding.DataBodyExcelFile item in dataBodyExcelFiles)
            {
                IRow newRow = sheet.CreateRow(rowIndex);
                newRow.CreateCell(0).SetCellValue(item.column1);
                newRow.CreateCell(1).SetCellValue(item.column2);

                newRow.CreateCell(2).SetCellValue(item.column3);
                newRow.CreateCell(3).SetCellValue(item.column4);
                newRow.CreateCell(4).SetCellValue(item.column5);
                newRow.CreateCell(5).SetCellValue(item.column6);
                newRow.CreateCell(6).SetCellValue(item.column7);
                if (item.column2 == "ISBNコード")
                    newRow.CreateCell(6).SetCellValue("申請数");
                newRow.CreateCell(7).SetCellValue(item.column8);
                rowIndex++;
            }
            MergeStyleWithCell.Instance.SetCellStyleOneRow(wb, sheet, 11, 1, 7, "table");
            MergeStyleWithCell.Instance.SetCellStyleMultiRow(wb, sheet, 12, rowIndex - 1, 1, 7, "tableBody");
            MergeStyleWithCell.Instance.SetCellStyleMultiRow(wb, sheet, 12, rowIndex - 1, 6, 6, "bgYellow");
            if (File.Exists(output))
                File.Delete(output);
            FileStream fileStream = new FileStream(output, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            wb.Write(fileStream);
            fileStream.Close();
        }
    }
}
