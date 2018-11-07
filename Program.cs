using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SautinSoft;

using static System.Console;

namespace PDF_to_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                WriteLine("Program launched without a file path");
                WriteLine("Close the program and drag a file onto this .exe in Windows Explorer");
                ReadKey();
                Environment.Exit(0);
            }
            else if(!args[0].ToLower().EndsWith(".pdf"))
            {
                WriteLine("Cannot convert file as it's not a PDF\nPress any key to close");
                ReadKey();
                Environment.Exit(0);
            }
            else
            {
                string wordDoc = ConvertToDocx(args[0]);

                WriteLine("Press enter to view the file or any other key to close without viewing");

                switch(ReadKey().Key)
                {
                    case ConsoleKey.Enter:
                        System.Diagnostics.Process.Start(wordDoc);
                        break;
                    default:
                        Environment.Exit(0);
                        break;
                }
            }
        }

        static string ConvertToDocx(string filename)
        {
            string docxFilename = filename.Substring(0, filename.Length - 4) + ".docx";

            SautinSoft.PdfFocus pdf = new SautinSoft.PdfFocus();

            pdf.OpenPdf(filename);

            if (pdf.PageCount > 0)
            {
                pdf.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;
            }

            WriteLine("Converting file, this may take a while");
            Stopwatch sw = new Stopwatch();
            sw.Start();
            int result = pdf.ToWord(docxFilename);

            if (result == 0)
            {
                WriteLine("Conversion successful");
                sw.Stop();
                WriteLine("Conversion took " + sw.Elapsed);
                return docxFilename;
            }
            else
            {
                WriteLine("Conversion failed, try again.");
                sw.Stop();
                ReadKey();
                Environment.Exit(0);
                return null;
            }
        }
    }
}
