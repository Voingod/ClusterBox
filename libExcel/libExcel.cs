using System;

namespace libExcel
{
    class libExcel
    {
        static void Main(string[] args)
        {
            Test();
            Read read = new Read();
            read.ReadFile();
        }

        private static void Test()
        {
            Console.WriteLine("Test");
           // Console.ReadLine();
        }
    }

    public class Read
    {
        public void ReadFile ()
        {
            Console.WriteLine("read");
           
            Write();
            Console.ReadLine();
        }

        private void Write()
        {

            Console.WriteLine("write");
        }
    }
}
