using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Novacode;

namespace CopyCSVtoWordTable
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Need to drag and drop file onto this EXE");
                Console.WriteLine("Exiting");
                return;
            }

            string[] delimiters = { ",", "\r\n" };
            String[] values = File.ReadAllText(args[0]).Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

            int rows = 0;
            for (int i = 0; i < values.Count(); i++)
            {
                if (values[i] == "Endline")
                {
                    rows++;
                }
            }

            using (
                DocX document =
                    DocX.Create(
                        Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), "Test.docx")))
            {
                Table t = document.AddTable(rows, 1);
                int cell = 0;
                foreach (var value in values)
                {
                    if (value != "" && value != "Endline")
                    {
                        t.Rows[cell].Cells[0].Paragraphs.First().Append(value);
                        t.Rows[cell].Cells[0].Paragraphs.First().Append("\r\n");
                    }
                    if (value == "Endline")
                    {
                        cell++;
                    }
                }
                document.InsertTable(t);

                document.Save();
            }
            Console.WriteLine("Finished output");
            Console.WriteLine("Press anykey to exit");
            Console.ReadKey();
        }
    }
}
