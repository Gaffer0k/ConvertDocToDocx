using System;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace WordConverter {
  
    class Program {
        static void ConvertDocToDocx(string path)
        {
            Application WordApp = new Application();

            if (path.ToLower().EndsWith(".doc"))
            {
                var srcFile = new FileInfo(path);
                if (!srcFile.Exists)
                    return;
                var Doc = WordApp.Documents.Open(srcFile.FullName);

                string newDocName = srcFile.FullName.Replace(".doc", ".docx");
                Doc.SaveAs2(newDocName, WdSaveFormat.wdFormatXMLDocument,
                                 CompatibilityMode: WdCompatibilityMode.wdWord2010);

                WordApp.ActiveDocument.Close();
                WordApp.Quit();
            }
        }
        static void Main(string[] args)
        {
            foreach (var path in args) {
                ConvertDocToDocx(path);
            }
        }
    }
}
