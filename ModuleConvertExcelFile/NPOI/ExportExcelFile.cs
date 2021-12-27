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
        public void ExportExcelFileWithNPOI(string @output, List<DataBinding.DataHeaderExcel> dataHeaderExcels, List<DataBinding.DataBodyExcelFile> dataBodyExcelFiles)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "sheet123");
            ICellStyle style = wb.CreateCellStyle();
            IRow row = sheet.CreateRow(0);
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
            //sheet.SetColumnWidth(8, 6000);
            //this.SetColumnStyle(wb, sheet);
            //end
            row.GetCell(1).SetCellValue(dataHeaderExcels[0].column2);
            SetExcelCellStyle.Instance.FontChange(wb, "title", row, 1);
            int rowIndex = 1;
            dataHeaderExcels.RemoveAt(0);
            dataHeaderExcels.RemoveAt(dataHeaderExcels.Count - 1);

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
                if (rowIndex > 1 && rowIndex < 8)
                    for (int i = 1; i <= 7; i++)
                    {
                        SetExcelCellStyle.Instance.FontChange(wb, "content", newRow, i);
                    }
                else
                    for (int i = 1; i <= 7; i++)
                        SetExcelCellStyle.Instance.FontChange(wb, "description", newRow, i);
                if (item.column2 != null && item.column2 != "" && item.column2 != "null" && rowIndex <= 8)
                {
                    SetExcelCellStyle.Instance.FontChange(wb, "label", newRow, 1);
                }
                if (rowIndex == 3)
                {
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 5);
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 6);
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 7);
                }
                if (rowIndex == 4)
                {
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 3);
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 4);
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 5);
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 6);
                    SetExcelCellStyle.Instance.FontChange(wb, "bgBlack", newRow, 7);
                }
                rowIndex++;
            }

            //data body
            rowIndex++;
            foreach (DataBinding.DataBodyExcelFile item in dataBodyExcelFiles)
            {
                IRow newRow = sheet.CreateRow(rowIndex);
                newRow.CreateCell(0).SetCellValue(item.column1);
                newRow.CreateCell(1).SetCellValue(item.column2);

                newRow.CreateCell(2).SetCellValue(item.column3);
                newRow.CreateCell(3).SetCellValue(item.column4);
                if (rowIndex != 13)
                    newRow.CreateCell(4).SetCellValue(Int32.Parse(item.column5));
                newRow.CreateCell(4).SetCellValue(item.column5);
                newRow.CreateCell(5).SetCellValue(item.column6);
                newRow.CreateCell(6).SetCellValue(item.column7);
                if (item.column2 == "ISBNコード")
                    newRow.CreateCell(6).SetCellValue("申請数");
                newRow.CreateCell(7).SetCellValue(item.column8);
                if (rowIndex == 13)
                    for (int i = 1; i <= 7; i++)
                    {
                        SetExcelCellStyle.Instance.FontChange(wb, "table", newRow, i);
                    }
                else
                    for (int i = 1; i <= 7; i++)
                    {
                        SetExcelCellStyle.Instance.FontChange(wb, "tableBody", newRow, i);
                    }
                rowIndex++;
            }


            if (File.Exists(output))
                File.Delete(output);
            FileStream fileStream = new FileStream(output, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            wb.Write(fileStream);
            fileStream.Close();
        }
    }
}
