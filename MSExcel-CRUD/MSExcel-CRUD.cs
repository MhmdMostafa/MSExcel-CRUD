using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;
namespace MSExcel_CRUD
{
    class MSExcel_CRUD
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        public static int XlAppProcessID;
        public static Application XlApp = null;
        public static Workbook XlWorkBook = null;
        public static Worksheet XlWorkSheet = null;

        public static void Create(int row, params string[] args)
        {
            for (int index = 0; index < args.Length; index++)
            {
                Update(row, index, String.Format(args[index]));
            }

        }

        public static string Read(int row, int column)
        {
            return XlWorkSheet.Cells[row, column].Text;
        }

        public static bool Update(int row, int column, string STR)
        {
            XlWorkSheet.Cells[row, column].Text = STR;
            return true;
        }

        public static bool Delete(int row, int column)
        {
            XlWorkSheet.Cells[row, column].Text = "";
            return true;
        }


        public Process GetExcelProcess()
        {
            GetWindowThreadProcessId(XlApp.Hwnd, out XlAppProcessID); ;
            return Process.GetProcessById(XlAppProcessID);
        }

    }
}
