using System;
using System.Linq;

namespace MSExcel_CRUD
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }

        //Remove The beginning and ending Spaces form given string
        public string RmSpace(string STR)
        {
            while (STR.Last() == ' ')
                STR = STR.Substring(0, STR.Length - 1);
            while (STR[0] == ' ')
                STR = STR.Substring(1, STR.Length - 1);
            return STR;
        }

        //Remove specific characters form given string
        public string RmSystemCharechters(string STR)
        {
            return STR.Replace("|", "").Replace(">", "").Replace("<", "").Replace("\"", "").Replace("?", "").Replace(":", "").Replace(",", "").Replace("/", "").Replace("\\", "");
        }
    }
}
