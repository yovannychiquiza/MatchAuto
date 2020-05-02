using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.FileExtensions;
using Microsoft.Extensions.Configuration.Json;

namespace MatchAuto
{
    class Program
    {
        static void Main(string[] args)
        {

            //IConfiguration configuration = new ConfigurationBuilder()
            //.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            //.Build();

            //var PathFrom = configuration.GetValue<string>("MySettings:PathFrom");
            //var PathTo = configuration.GetValue<string>("MySettings:PathTo");

            Console.WriteLine("Match Process start");
            Process process = new Process();
            process.Match();

        }
    }
}
