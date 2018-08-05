using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace xml_converter
{
    class Writer
    {
        private static String dir;
        private ExcelPackage p;

        public Writer(String d)
        {
            dir = d;
            p = new ExcelPackage();
        }

        // initialize dir-specific workbook
        public void Headers(String file)
        {
            String content = File.ReadAllText(file);
            String[] heads = content.Split("\r");

            //Creates a template workbook. Package is disposed when done.
            var ws = p.Workbook.Worksheets.Add("Content Values");
            ws.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Row(1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.MidnightBlue);
            ws.Row(1).Style.Font.Color.SetColor(System.Drawing.Color.White);
            ws.Row(1).Style.Font.Bold = true;
            ws.Row(1).Style.Font.Name = "Calibri Light";
            ws.Row(1).Style.Font.Size = 10;

            ws.Cells.Style.Font.Name = "Calibri Light";
            ws.Cells.Style.Font.Size = 10;

            for (int i = 0; i < heads.Length - 1; i++)
            {
                //Console.WriteLine(i + ": " + heads[i]);
                ws.Cells[1, i + 1].Value = heads[i];
            }

            //p.SaveAs(new FileInfo(dir + "test_wb.xlsx"));

            Contents(heads);
        }

        // write cell data
        private void Write(String value, int col, int row)
        {
            //Console.WriteLine("@[" + row + ", " + col + "]: " + value);
            var ws = p.Workbook.Worksheets[0];
            ws.Cells[row, col].Value = value;
        }

        // check cell data by reading ahead -- refactored out as Scanner
        private String Check(String contents, String type)
        {
            StreamReader sr = new StreamReader(contents);
            String line = "";
            if (type.Equals("tail"))
            {
                line = sr.ReadLine();
                return line.Substring(line.IndexOf(":") + 2);
            }
            else if (type.Equals("attr"))
            {
                String value = "";
                while (line.Contains("\t"))
                {
                    value += line.Substring(line.IndexOf(":") + 2) + ";";
                    line = sr.ReadLine();
                }
                if (line.Substring(0, line.IndexOf(":")).Contains("#text"))
                {
                    return line.Substring(line.IndexOf(":") + 2);
                }
            }
            return "";
        }

        // determine workbook's contents
        private void Contents(String[] h)
        {
            String ndir = dir + "content_files/";
            String line;
            int col = 1;
            int row = 2;
            String value = "";
            StreamReader r;
            foreach (String file in Directory.EnumerateFiles(ndir, "*.txt"))
            {
                r = new StreamReader(file);
                Scanner sc = new Scanner(r);

                while ((line = sc.ReadLine()) != null)
                {
                    value = "";
                    try
                    {
                        if (h.Contains(line.Substring(0, line.IndexOf(":")))) // if is head tag
                        {
                            col = Array.IndexOf(h, line.Substring(0, line.IndexOf(":"))) + 1;   //determine column to write to
                            //Console.WriteLine("Head @ " + col + " > " + line.Substring(0, line.IndexOf(":")) + "      ");

                            if (line.Substring(0, line.IndexOf(":")).Equals("desig"))
                            {
                                row++;
                            }

                            String peekLine = sc.PeekLine();
                            if (line.Contains(" [ "))  //if attributes exist
                            {
                                value = line.Substring(line.IndexOf(":") + 2);
                                //Console.WriteLine("     Attribute @" + col + " > " + value);
                            }
                            if (peekLine.Substring(0, peekLine.IndexOf(":")).Contains("#text")) // if is head content
                            {
                                value = peekLine.Substring(peekLine.IndexOf(":") + 2);
                                //Console.WriteLine("     Text @" + col + " > " + value);
                            }
                            Write(value, col, row);
                        }
                        else { }
                    }
                    catch (Exception e) { Console.WriteLine(e);  };
                }
                p.SaveAs(new FileInfo(dir + "output.xlsx"));
                Console.WriteLine("File: " + Path.GetFileName(file) + " @row=" + row);
            }
        }

        // Contents() but with a rogue + shallow copy scanner
        private void Explore(String[] h)
        {
            String ndir = dir + "test/";
            String line;
            int col = 1;
            int row = 2;
            String value = "";
            StreamReader r;
            foreach (String file in Directory.EnumerateFiles(ndir, "*.txt"))
            {
                r = new StreamReader(file);
                while((line = r.ReadLine()) != null)
                {
                    if (h.Contains(line.Substring(0, line.IndexOf(":")))) //is head tag
                    {
                        col = Array.IndexOf(h, line.Substring(0, line.IndexOf(":"))) + 1;
                        Console.WriteLine("Found Head @ " + col + " > " + line.Substring(0, line.IndexOf(":")) + "      ");
                        //line = r.ReadLine();
                        if (line.Substring(0, line.IndexOf(":")).Contains("#text"))
                        {
                            value = line.Substring(line.IndexOf(":") + 2);
                            //Write(value, col, row);
                            Console.WriteLine(" Text found @" + col + " > " + value);
                            
                        }
                        else
                        {
                            value = ""; //reset
                            while (line.Contains("\t"))
                            {
                                value += line.Substring(line.IndexOf(":") + 2) + ";";
                                line = r.ReadLine();
                            }
                            //Write(value, col, row);
                            Console.WriteLine(" Attribute found @" + col + " > " + value);
                            if (line.Substring(0, line.IndexOf(":")).Contains("#text"))
                            {
                                value = line.Substring(line.IndexOf(":") + 2);
                                //Write(value, col, row);
                                Console.WriteLine(" Overriden > text found @" + col + " > " + value);
                            }
                        }
                        Write(value, col, row);
                        row++;
                    }
                }
                p.SaveAs(new FileInfo(dir + "test_wb.xlsx"));
            }
        }
    }
}
