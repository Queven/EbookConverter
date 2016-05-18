using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Configuration;

namespace EbookConverter
{
    class Program
    {
        
        private const string epubQTools = @"epubQTools\epubQTools.exe";
        private const string calibre = "calibredb ";
        private const string epubQToolsArgs = @" {0} -n -e --book-margin -40 -kd";
        private const string calibreArgs = @"add --add=""*.*"" {0}";
        private static string path = ConfigurationManager.AppSettings["Path"];
        private const string mail = "calibre-smtp";
        private static string from = ConfigurationManager.AppSettings["MailFrom"];
        private static string to = ConfigurationManager.AppSettings["MailTo"];
        private const string mailArgs = "{0} {1} \"book\" -a \"{2}\"";
        static void Main(string[] args)
        {
            DoWork(epubQTools, string.Format(epubQToolsArgs, path));
            RenameFile(path, "*_moh.mobi");
            RemoveFiles(path,"*_moh.epub");
            
            DoWork(calibre, string.Format(calibreArgs, path));
            foreach (var mobi in Directory.GetFiles(path, "*.mobi"))
            {
                DoWork(mail, string.Format(mailArgs, from, to, mobi));
            }
            RemoveFiles(path);
            Console.WriteLine("******************************************************");
            Console.WriteLine("*===================End Program !!===================*");
            Console.WriteLine("******************************************************");
            Console.ReadKey();
            
        }

        private static void RenameFile(string path, string pattern)
        {
            foreach (var file in Directory.GetFiles(path, pattern))
            {
                File.Move(file, file.Replace("_moh",""));
            }
        }

        private static void RemoveFiles(string path,string pattern="*")
        {
            foreach (var file in Directory.GetFiles(path, pattern))
                File.Delete(file);
        }

        private static void DoWork(string programName, string programArgs)
        {
            Console.WriteLine("******************************************************");
            Console.WriteLine(string.Format("*================Start Process {0}!!!================*", programName.ToUpper()));
            Console.WriteLine("******************************************************");
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = programName;
            start.Arguments = programArgs;
            start.UseShellExecute = false;
            Process process = Process.Start(start);
            process.WaitForExit();
            Console.WriteLine(string.Format("================End Process {0}!!!================", programName.ToUpper()));
        }
    }
}
