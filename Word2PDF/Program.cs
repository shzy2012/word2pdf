using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace Word2PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("缺失参数文件夹路径.格式：[程序] [空格] [路径]  example: ./Word2PDF.exe {0}", @"j:\files");
                return;
            }
                
            string dirPath = args[0];

            if (!Directory.Exists(dirPath))
            {
                Console.WriteLine("当前路径不存在：{0}", dirPath);
                return;
            }

            Console.WriteLine("当前输入路径：{0}", dirPath);

            // Create a new Microsoft Word application object
            Application word = new Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            // Get list of Word files in specified directory
            DirectoryInfo dirInfo = new DirectoryInfo(dirPath);
            FileInfo[] wordFiles = dirInfo.GetFiles("*.docx");


            word.Visible = false;
            word.ScreenUpdating = false;

            foreach (FileInfo wordFile in wordFiles)
            {

                // Cast as Object for word Open method
                Object filename = (Object)wordFile.FullName;
                Console.WriteLine("[生成]=>开始处理:{0}", filename);

                // Use the dummy value as a placeholder for optional arguments
                Document doc = word.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                object outputFileName = wordFile.FullName.Replace(".docx", ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                doc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.                
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;

                Console.WriteLine("[生成]=>处理结束:{0}", outputFileName);
            }

            // word has to be cast to type _Application so that it will find
            // the correct Quit method.
            ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
            
        }
    }
}
