using System;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MSExcel_CRUD
{
    class MSEx_CRUD
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        public static int XlAppProcessID;
        public static Application XlApp = null;
        public static Workbook XlWorkBook = null;
        public static Worksheet XlWorkSheet = null;

        public MSEx_CRUD(string FilePath)
        {
            XlWorkBook = XlApp.Workbooks.Open(@FilePath);
            XlWorkSheet = (Excel.Worksheet)XlWorkBook.Worksheets.get_Item(1);
            GetWindowThreadProcessId(XlApp.Hwnd, out XlAppProcessID);
        }


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

        public void SaveExcelWork()
        {
            if (XlApp != null)
            {
                try
                {
                    XlApp.ActiveWorkbook.Save();
                }
                catch (System.NullReferenceException error)
                {
                    Console.WriteLine($"You have closed the Excel File Please open your project again and do not close it again please\nErorr: {error}");
                    XlApp.Quit();
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlWorkBook);
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlWorkSheet);
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlApp);
                    Process.GetProcessById(XlAppProcessID).Kill();
                    XlApp = null;
                }
            }
        }

        public void CloseExcel()
        {
            if (XlApp != null)
            {
                try
                {
                    XlApp.ActiveWorkbook.Close(0);
                    XlApp.Quit();
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlWorkBook);
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlWorkSheet);
                    Process.GetProcessById(XlAppProcessID).Kill();
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlApp);
                    XlApp = null;

                }
                catch (System.NullReferenceException error)
                {

                    XlApp.Quit();
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlWorkBook);
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlWorkSheet);
                    Process.GetProcessById(XlAppProcessID).Kill();
                    //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(XlApp);
                    XlApp = null;

                    Console.WriteLine($"You have closed the Excel File Please open your project again and do not close it again please\nErorr: {error}");
                }
            }
        }
    }
}
