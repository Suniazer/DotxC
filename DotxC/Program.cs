using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;

namespace DotxC
{
    internal class Program
    {
        private static string SelectFile(string filter, string title, bool multi = false)
        {
        re:

            OpenFileDialog OpenFD = new OpenFileDialog
            {
                Multiselect = multi,
                Filter = filter,
                Title = title
            };
            if (!(bool)OpenFD.ShowDialog()) { goto re; }
            return OpenFD.FileName;
        }
        [STAThread]
        static void Main(string[] args)
        {
        start:
            Console.WriteLine("模板位置：");
            string tpath = SelectFile("文档模板|*.dotx", "选取文档模板");
            //if (Console.ReadKey().Key != ConsoleKey.Enter) { goto start; }
            Console.Write(tpath + "\r\n");
            Console.WriteLine();
            Dictionary<string ,string > pairs= new Dictionary<string ,string>();
        rpair:
            Console.WriteLine("键值对位置：");
            string ppath = SelectFile("键值对|*.txt", "选取文本");
            //if (Console.ReadKey().Key != ConsoleKey.Enter) { goto rpair; }
            Console.Write(ppath + "\r\n");
            Console.WriteLine();
            foreach (var line in File.ReadAllLines(ppath))
            {
                try
                {
                    string[] strings= line.Split(new char[] {'|' });
                    if (strings.Length != 2) { continue; }
                    pairs.Add(strings[0].Trim(), strings[1].Trim());
                }
                catch(Exception ex) { Console.WriteLine(ex.Message); goto rpair; }  
            }
            foreach (KeyValuePair<string, string> pair in pairs)
            {
                Console.WriteLine($"[{pair.Key}]:[{pair.Value}]");
            }
        rsave:
            Console.WriteLine("保存位置：");
            SaveFileDialog SaveFD = new SaveFileDialog();
            SaveFD.Filter = "文档|*.doc";
            if (!(bool)SaveFD.ShowDialog()) { goto rsave; }
            string spath = SaveFD.FileName;
            Console.WriteLine($"{spath}");
            try
            {
                HandT.Replace(tpath, spath, pairs);
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); goto start; }
            Console.WriteLine();
            Console.WriteLine("[结束]");
            Console.ReadLine();
        }
    }
}
