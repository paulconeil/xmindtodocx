using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMindInterviewToDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            string destinationFile, destinationPath, sourceFile;

            sourceFile = args[0];

            Console.WriteLine("Loading: " + sourceFile);
            XmindDocLoader.XmindDocLoader xmindDocLoader = new XmindDocLoader.XmindDocLoader(sourceFile);
            Console.WriteLine("Loaded: " + sourceFile);
            Console.WriteLine(xmindDocLoader.GenerateInterview().ToString());
            Console.ReadLine();
            //load xmind interview document



            
            
        

        }
    }
}
