using System;

namespace DummyApp
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.ReadLine();

                try
                {
                    throw new ApplicationException("TEST");
                }
                catch (Exception e)
                {
                }
            }
        }
    }
}
