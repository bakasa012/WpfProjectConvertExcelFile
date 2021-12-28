using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModuleConvertExcelFile.ProcessPath
{
    class OpenDialog
    {
        private static OpenDialog instance;

        public static OpenDialog Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new OpenDialog();
                }

                return instance;
            }

            set { OpenDialog.instance = value; }
        }

        private OpenDialog() { }
        public string OpentFolderDialog(System.Windows.Controls.Label label)
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = folderBrowserDialog.ShowDialog();
            string strPath = "";
            if ((int)result == 1 && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
            {
                if (label.Name == "lbLinkUrlUpload")
                    strPath= folderBrowserDialog.SelectedPath;
                else
                    strPath =  folderBrowserDialog.SelectedPath;
                label.Content = folderBrowserDialog.SelectedPath;
            }
            return strPath;
        }
    }
}
