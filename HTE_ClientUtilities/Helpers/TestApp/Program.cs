using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FileSystemWatch;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Watcher watcher = new Watcher();

            Console.WriteLine("Press 'Enter' to exit");
            Console.ReadLine();
        }
    }
}
