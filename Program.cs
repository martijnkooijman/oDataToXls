using oDataToXls.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace oDataToXls
{
    class Program
    {
        static void Main(string[] args)
        {
            string baseUrl = ConfigurationManager.AppSettings["oDataUrl"];
            string fileName = ConfigurationManager.AppSettings["outputFileName"];
            var builder = new oDataXlsBuilder();
            builder.Build(baseUrl, fileName).Wait();
            Console.WriteLine("finished");
            Console.ReadKey();
        }
    }
}
