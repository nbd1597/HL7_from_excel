using System;
using System.Diagnostics;
using System.IO;
using NHapi.Base.Model;
using NHapi.Base.Parser;

namespace ExceltoHl7
{
    public class Program
    {
        static void Main(string[] args)
        {
            //ExcelParser.ReadTemplte();
            Console.WriteLine("********HL7 ADT_A01 MESSAGE MAKER**********");
            Console.WriteLine("Select EXCEL.exe location");
            String ExcelPathFile = Directory.GetCurrentDirectory() + "\\excelpath.txt"; //current excel path
            Console.WriteLine("type 'yes' to select new path, 'enter' to skip");
            using (StreamReader ReadExcelPath = new StreamReader(ExcelPathFile))
            {
                Console.WriteLine(String.Format("Default location is at: {0}", ReadExcelPath.ReadLine()));
            }
            //get new location of excel.exe and save to excelpath.txt
            if (Console.ReadLine() == "yes")
            {
                Console.WriteLine("Insert new location:");
                File.WriteAllText(ExcelPathFile, String.Empty);
                using (StreamWriter WriteExcelPath = new StreamWriter(ExcelPathFile))
                {
                    WriteExcelPath.WriteLine(Console.ReadLine());
                }
                using (StreamReader ReadExcelPath = new StreamReader(ExcelPathFile))
                {
                    Console.WriteLine(String.Format("New location is at: {0}", ReadExcelPath.ReadLine()));
                }
                    
            }
            else { }
            //main loop
            while (true)
            {
                Console.Clear();
                Console.WriteLine("********HL7 ADT_A01 MESSAGE MAKER**********");
                try
                {
                    var adtMessage = AdtMessageFactory.CreateMessage("A01");    //make the message object
                    var pipeParser = new PipeParser();  //api parser
                    WriteMessageFile(pipeParser, adtMessage, 
                        String.Format("{0}\\HL7TestOutputs\\hl7", Directory.GetCurrentDirectory()), 
                        String.Format("{0}.txt", ExcelParser.NewFile)); // write to file
                    Console.WriteLine("make another ?");
                    if (Console.ReadLine() == "yes") continue;
                    else return;
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error occured while creating HL7 message {e.Message}");
                }
                
            }
        }
        //write 
        private static void WriteMessageFile(ParserBase parser, IMessage hl7Message, string outputDirectory, string outputFileName)
        {
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            var fileName = Path.Combine(outputDirectory, outputFileName);

            Console.WriteLine("Writing data to file...");

            if (File.Exists(fileName))
                File.Delete(fileName);
            File.WriteAllText(fileName, parser.Encode(hl7Message));
            Console.WriteLine($"Wrote data to file {fileName} successfully...");
        }


    }
}
