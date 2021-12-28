using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModuleConvertExcelFile.ProcessPath
{
    public class CreateOrGetFolder
    {
        public List<string> files = new List<string>();
        public void CheckPathOrCreateFolder(string path)
        {
            if (!Directory.Exists(path)) 
            {
                Directory.CreateDirectory(path); 
            }
        }
        public void RecursiveDirectory(string path)
        {
            string[] folder = Directory.GetDirectories(path);
            string[] file = Directory.GetFiles(path);
            for (int i = 0; i < Directory.GetFiles(path).Length; i++)
            {
                files.Add(Directory.GetFiles(path)[i]);

            }
            if (folder.Length > 0)
            {
                for (int i = 0; i < folder.Length; i++)
                {
                    RecursiveDirectory(folder[i]);
                }
            }
        }
    }
}
