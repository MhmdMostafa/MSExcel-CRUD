using System;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;

namespace MSExcel_CRUD
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string path;
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"d:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xlsx files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog1.FileName;
                CallMethod(path);
            }

            else
            {
                Console.WriteLine("Bad Path");
            }

        }
        public void test()
        {
            Console.WriteLine("correct");
        }

        static public void CallMethod(string path)
        {
            string command;
            Type main = typeof(Program);
            MethodInfo Method;
            bool loop = true;
            while (loop)
            {
                Console.WriteLine("give me the command: ");
                command = Console.ReadLine();
                switch (command)
                {
                    case "RmSpace":
                        Method = main.GetMethod("RmSpace()");
                        //Method.Invoke(Program, null);
                        break;
                    case "RmSystemCharechters":
                        break;
                    case "GetFilePath":
                        break;
                    case "MSEx_CRUD":


                        MSEx_CRUD ExObj = new MSEx_CRUD(path);
                        Type ClassExcelType = typeof(MSEx_CRUD);
                        break;
                    case "test":
                        Method = main.GetMethod("RmSpace()");
                        //Method.Invoke(this, null);
                        break;
                    case "exit":
                        loop = false;
                        break;
                    default:
                        Console.WriteLine("Wrong command");
                        break;
                }
            }
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


        static public string GetFilePath()
        {
            Console.WriteLine("test");
            return "test";
        }
    }
}
