using System;


namespace MSExcel_CRUD
{
    class Program
    {
        
        static void Main(string[] args)
        {


            LogSystem Log = new LogSystem("C:\\Users\\unkno\\Desktop\\drivers\\t1.txt");
            Console.WriteLine("done");
            Console.ReadLine();

        }
        static public void test()
        {
            Functions.TestExceptions(() => { int a = Convert.ToInt32(Console.ReadLine()); int c = Convert.ToInt32(Console.ReadLine()); Console.WriteLine(a / c); });
                
            
            

        }

    }
}
