using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuleConvertExcelFile.NPOI
{
    public class ReadExcelFile
    {
        public void ReadFileExcelWitdNPOI(string pathFile, List<DataBinding.DataHeaderExcel> dataHeaderExcels, List<DataBinding.DataBodyExcelFile> dataBodyExcelFiles)
        {
            IWorkbook wb = null;
            dataHeaderExcels.Clear();
            dataBodyExcelFiles.Clear();
            using (FileStream fileStream = new FileStream(@pathFile, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    switch (Path.GetExtension(pathFile).ToLower())
                    {
                        case ".xls":
                            wb = new HSSFWorkbook(fileStream);
                            break;
                        case ".xlsx":
                            wb = new XSSFWorkbook(fileStream);
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception)
                {

                    throw;
                }

                ISheet sheet = wb.GetSheetAt(0);
                int lastRow = sheet.LastRowNum;
                bool isDataBody = false;
                int rowIndex = 4;

                while (true)
                {
                    var nowRow = sheet.GetRow(rowIndex);
                    if (nowRow != null)
                    {
                        if (nowRow.GetCell(0)?.ToString() == "ISBNコード")
                        {
                            isDataBody = true;
                        }
                        if (!isDataBody)
                        {
                            DataBinding.DataHeaderExcel dataHeaderExcel = new DataBinding.DataHeaderExcel
                            {
                                column1 = " ",
                                column2 = nowRow.GetCell(0)?.ToString(),
                                column3 = nowRow.GetCell(1)?.ToString().Trim(),
                                column4 = nowRow.GetCell(2)?.ToString(),
                                column5 = nowRow.GetCell(3)?.ToString(),
                                column6 = nowRow.GetCell(4)?.ToString(),
                                column7 = "",
                                column8 = nowRow.GetCell(5)?.ToString(),
                            };
                            dataHeaderExcels.Add(dataHeaderExcel);
                        }
                        else
                        {
                            DataBinding.DataBodyExcelFile dataBodyExcelFile = new DataBinding.DataBodyExcelFile()
                            {
                                column1 = " ",
                                column2 = nowRow.GetCell(0)?.ToString(),
                                column3 = nowRow.GetCell(1)?.ToString(),
                                column4 = nowRow.GetCell(2)?.ToString(),
                                column5 = nowRow.GetCell(3)?.ToString(),
                                column6 = nowRow.GetCell(4)?.ToString(),
                                column7 = "",
                                column8 = nowRow.GetCell(5)?.ToString(),
                            };
                            dataBodyExcelFiles.Add(dataBodyExcelFile);
                        }

                    }
                    else if (!isDataBody)
                    {
                        DataBinding.DataHeaderExcel dataHeaderExcel = new DataBinding.DataHeaderExcel
                        {
                            column1 = "",
                            column2 = "",
                            column3 = "",
                            column4 = "",
                            column5 = "",
                            column6 = "",
                            column7 = "",
                            column8 = "",
                        };
                        dataHeaderExcels.Add(dataHeaderExcel);
                    }
                    if (rowIndex >= lastRow - 1)
                        break;
                    rowIndex++;
                }

                fileStream.Close();
            };
        }
    }
}
