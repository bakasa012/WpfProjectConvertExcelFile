using ModuleConvertExcelFile.DataBinding;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ModuleConvertExcelFile.ProcessPath;
using ModuleConvertExcelFile.NPOI;
using System.IO;

namespace ModuleConvertExcelFile.View
{
    /// <summary>
    /// Interaction logic for ConvertExcelFileView.xaml
    /// </summary>
    public partial class ConvertExcelFileView : UserControl
    {
        private string pathFolderImportFileExcelGlobal = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DataExcelImport";//@"C:\Users\lieu.hong.thai\Downloads\dataExcel";//@"C:\DataExcelImport";
        private string pathFolderExportFileExcelGlobal = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DataExcelExport";//@"C:\DataExcelExport\";
        List<string> files = new List<string>();
        private List<DataHeaderExcel> dataHeaderExcels = new List<DataHeaderExcel>();
        private List<DataBodyExcelFile> dataBodyExcelFiles = new List<DataBodyExcelFile>();
        List<DataBindingCompanyCode> dataBindingCompanyCodes = new List<DataBindingCompanyCode>();
        public ConvertExcelFileView()
        {
            InitializeComponent();
            LoadOnStartUp();
        }

        private void LoadOnStartUp()
        {
            lbLinkUrlUpload.Content = pathFolderImportFileExcelGlobal;
            lbLinkUrlSaveFile.Content = pathFolderExportFileExcelGlobal;
            CreateOrGetFolder createOrGetFolder = new CreateOrGetFolder();
            ProcessPathFolderOrFile processPathFolderOrFile = new ProcessPathFolderOrFile();
            createOrGetFolder.CheckPathOrCreateFolder(pathFolderImportFileExcelGlobal);
            createOrGetFolder.CheckPathOrCreateFolder(pathFolderExportFileExcelGlobal);
            processPathFolderOrFile.CheckDataJson("変換M.xls", "dataStore.json",dataBindingCompanyCodes);
        }


        private void lbLinkUrlUpload_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            pathFolderImportFileExcelGlobal = OpenDialog.Instance.OpentFolderDialog(lbLinkUrlUpload);
            MessageBox.Show(pathFolderImportFileExcelGlobal);
        }

        private void lbLinkUrlSaveFile_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            pathFolderExportFileExcelGlobal = OpenDialog.Instance.OpentFolderDialog(lbLinkUrlSaveFile);
        }

        private void btnCloseApp_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxManager.OK = "Alright";
            MessageBoxManager.Yes = "Yep!";
            MessageBoxManager.No = "Nope";
            MessageBoxManager.Register();
            MessageBoxResult messageBoxResult = MessageBox.Show("some one", "some one!", MessageBoxButton.YesNoCancel);
            if(messageBoxResult == MessageBoxResult.Yes)
            {
                Application.Current.MainWindow.Close();
            }
            else
            {

            }
        }

        [Obsolete]
        private void btnConvertExcelFile_Click(object sender, RoutedEventArgs e)
        {
            CreateOrGetFolder createOrGetFolder = new CreateOrGetFolder();
            createOrGetFolder.RecursiveDirectory(pathFolderImportFileExcelGlobal);
            string[] files = createOrGetFolder.files.ToArray();
            int countFile = 0;
            //if (dpStartDate.Text.Length > 0)
            //{
            string strStartDate = "a";//dpStartDate.Text?.Split('/')[1];
                for (int i = 0; i < files.Length; i++)
                {
                    string fileExt = Path.GetExtension(files[i]);
                    if (fileExt == ".xls" || fileExt == ".xlsx")
                    {
                        countFile++;
                        ReadExcelFile.Instance.ReadFileExcelWitdNPOI(@files[i], dataHeaderExcels, dataBodyExcelFiles);
                        DataBindingCompanyCode retData = GetDataCompanyCode.Instance.Get(dataBindingCompanyCodes, dataHeaderExcels[3].column3, dataHeaderExcels[3].column5);
                        string outputNameFile = retData.GEOStoreCode + "_" + retData.StoreName +
                            "-【" + strStartDate + "月末〆】" + strStartDate + "月度自己調達許諾シール給付申請書.xlsx";
                    ExportExcelFile.Instance.ExportExcelFileWithNPOI(pathFolderExportFileExcelGlobal +
                        @"\" + outputNameFile,
                        dataHeaderExcels, dataBodyExcelFiles);
                }
                }
                MessageBox.Show(pathFolderImportFileExcelGlobal + " " + "Files was conver : " + countFile.ToString(), "Information!");
            //}
            //else
            //{
            //    MessageBox.Show("Vui lòng nhập ngày tháng", "Information!");
            //}
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        }

    }
}
