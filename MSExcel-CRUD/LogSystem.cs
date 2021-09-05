using System;
using System.IO;

namespace MSExcel_CRUD
{
    class LogSystem
    {
        string MainPath;
        StreamWriter sw;
        public LogSystem(string Path)
        {
            MainPath = Path;
            if (!File.Exists(MainPath))
                sw = File.CreateText(MainPath);
            else
                sw = File.AppendText(MainPath);
        }
        public void ReadLogFile(){
            using (StreamReader sr = File.OpenText(MainPath))
            {
                string s = "";
                while ((s = sr.ReadLine()) != null)
                {
                    Console.WriteLine(s);
                }
            }
        }

        public void WriteLog(string Statement)
        {
            sw.WriteLine(string.Format("#- {0} [{1}]\n", Statement, DateTime.Now.ToString("dddd, dd MMMM yyyy HH: mm:ss tt")));
        }


    }
}
