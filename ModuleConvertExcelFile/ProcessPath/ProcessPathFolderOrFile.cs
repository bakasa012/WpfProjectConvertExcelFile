using ModuleConvertExcelFile.DataBinding;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ModuleConvertExcelFile.ProcessPath
{
    public class ProcessPathFolderOrFile
    {
        public void checkDataJson(string pathFileExcel, string pathFileJson, List<DataBindingCompanyCode> dataBindingCompanyCodes)
        {
            IWorkbook wb = null;
            string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string xslLocation = Path.Combine(executableLocation, pathFileExcel);
            string xslLocationJson = Path.Combine(executableLocation, pathFileJson);
            if (!File.Exists(xslLocationJson))
            {
                using (FileStream fileStream = new FileStream(xslLocation, FileMode.Open, FileAccess.Read))
                {
                    string fileExt = Path.GetExtension(pathFileExcel);
                        switch (fileExt.ToLower())
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
                    ISheet sheet = wb.GetSheetAt(0);
                    int lastRow = sheet.LastRowNum;
                    int rowIndex = 0;
                    while (rowIndex <= lastRow)
                    {
                        var nowRow = sheet.GetRow(rowIndex);
                        if (nowRow != null)
                        {
                            DataBindingCompanyCode dataBindingCompanyCode = new DataBindingCompanyCode()
                            {
                                CDVJMemberNumber = nowRow.GetCell(0)?.ToString(),
                                StoreName = nowRow.GetCell(1)?.ToString(),
                                GEOStoreCode = nowRow.GetCell(2)?.ToString(),
                            };
                            dataBindingCompanyCodes.Add(dataBindingCompanyCode);
                        }
                        rowIndex++;
                    }
                    var json = JsonConvert.SerializeObject(dataBindingCompanyCodes);
                    File.WriteAllText(xslLocationJson, json);
                    fileStream.Close();
                };
            }
            else
                LoadJson(xslLocationJson,dataBindingCompanyCodes);
        }
        private void LoadJson(string pathFileJson, List<DataBindingCompanyCode> dataBindingCompanyCodes)
        {
            using (StreamReader streamReader = new StreamReader(pathFileJson))
            {
                string json = streamReader.ReadToEnd();
                List<DataBindingCompanyCode> items = JsonConvert.DeserializeObject<List<DataBindingCompanyCode>>(json);
                dataBindingCompanyCodes = items;
                Console.WriteLine(items);
            };
        }
    }
}
