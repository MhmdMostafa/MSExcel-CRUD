using System;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using System.ServiceModel;

namespace MSExcel_CRUD
{
    class Functions
    {
        static public void Method()
        {
            string command;
            bool loop = true;
            while (loop)
            {
                Console.WriteLine("give me the command: ");
                command = Console.ReadLine();
                switch (command)
                {
                    case "RmSpace":
                        break;
                    case "RmSystemCharechters":
                        break;
                    case "GetFilePath":
                        break;
                    case "MSEx_CRUD":
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
        static public string RmSpace(string STR)
        {
            while (STR.Last() == ' ')
                STR = STR.Substring(0, STR.Length - 1);
            while (STR[0] == ' ')
                STR = STR.Substring(1, STR.Length - 1);
            return STR;
        }

        //Remove specific characters form given string
        static public string RmSystemCharechters(string STR)
        {
            return STR.Replace("|", "").Replace(">", "").Replace("<", "").Replace("\"", "").Replace("?", "").Replace(":", "").Replace(",", "").Replace("/", "").Replace("\\", "");
        }

        static public bool Openfile()
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
            return true;
        }

        static public void TestExceptions(Action codeBlock)
        {
            try
            {
                codeBlock();
            }
            catch (CommonException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
    }
}
