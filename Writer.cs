using OfficeOpenXml;
using System;
using System.Collections.Generic;
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

        // prepare workbook for wide data
        public void WideHeaders(String file)
        {
            String content = File.ReadAllText(file);
            String[] heads = content.Split("\r");

            //Creates a template workbook. Package is disposed when done.
            var ws = p.Workbook.Worksheets.Add("Wide Data");
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

            WideContents(heads);
        }

        // determine workbook's contents -- doesn't propagate values down
        private void WideContents(String[] h)
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
                                value = line.Substring(line.IndexOf(":") + 1).Trim();
                                //Console.WriteLine("     Attribute @" + col + " > " + value);
                            }
                            if (peekLine.Substring(0, peekLine.IndexOf(":")).Contains("#text")) // if is head content
                            {
                                value = peekLine.Substring(peekLine.IndexOf(":") + 1).Trim();
                                //Console.WriteLine("     Text @" + col + " > " + value);
                            }
                            Write(value, col, row);
                        }
                        else { }
                    }
                    catch (Exception e) { Console.WriteLine(e); };
                }
                p.SaveAs(new FileInfo(dir + "wide_output.xlsx"));
                Console.WriteLine("File: " + Path.GetFileName(file) + " @row=" + row);
            }
        }

        // prepare workbook for narrow data
        public void NarrowHeaders()
        {
            var ws = p.Workbook.Worksheets.Add("Content Values");
            ws.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Row(1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.MidnightBlue);
            ws.Row(1).Style.Font.Color.SetColor(System.Drawing.Color.White);
            ws.Row(1).Style.Font.Bold = true;
            ws.Row(1).Style.Font.Name = "Calibri Light";
            ws.Row(1).Style.Font.Size = 10;

            ws.Cells.Style.Font.Name = "Calibri Light";
            ws.Cells.Style.Font.Size = 10;

            ws.Cells[1, 1].Value = "#";
            //ws.Cells[1, 2].Value = "Source";
            ws.Cells[1, 2].Value = "Content ID";
            ws.Cells[1, 3].Value = "Date Last Updated";
            ws.Cells[1, 4].Value = "Parent Node";
            ws.Cells[1, 5].Value = "Subcontent ID";
            ws.Cells[1, 6].Value = "Subcontent Type";
            ws.Cells[1, 7].Value = "Subcontent Citation";
            ws.Cells[1, 8].Value = "Node Name";
            ws.Cells[1, 9].Value = "Node Value";

            NarrowContents();
        }

        public int[] FindIndex(Array haystack, object needle)
        {
            if (haystack.Rank == 1)
                return new[] { Array.IndexOf(haystack, needle) };

            var found = haystack.OfType<object>()
                              .Select((v, i) => new { v, i })
                              .FirstOrDefault(s => s.v.Equals(needle));
            if (found == null)
                throw new Exception("needle not found in set");

            var indexes = new int[haystack.Rank];
            var last = found.i;
            var lastLength = Enumerable.Range(0, haystack.Rank)
                                       .Aggregate(1,
                                           (a, v) => a * haystack.GetLength(v));
            for (var rank = 0; rank < haystack.Rank; rank++)
            {
                lastLength = lastLength / haystack.GetLength(rank);
                var value = last / lastLength;
                last -= value * lastLength;

                var index = value + haystack.GetLowerBound(rank);
                if (index > haystack.GetUpperBound(rank))
                    throw new IndexOutOfRangeException();
                indexes[rank] = index;
            }

            return indexes;
        }

        // get as narrow as possible
        private void NarrowContents()
        {
            // @params id which makes node.values meaningful (e.g., regulatory body/title)

            String ndir = dir + "content_files/";
            String line;
            int row = 2;
            String value = "";
            String reg_body = "";
            String content_id = "";
            String dc_id = "";
            String cite_normalized = "";
            String cite_type = "";
            String cite_id = "";
            StreamReader r;
            int start = 0;
            Boolean prop;
            Boolean f = false; // if start of new file
            //String[] files = ReadMeta(dir + "meta_data/meta.txt");

            foreach (String file in Directory.EnumerateFiles(ndir, "*.txt"))
            {
                r = new StreamReader(file);
                Scanner sc = new Scanner(r);

                while ((line = sc.ReadLine()) != null)
                {
                    value = "";
                    prop = false;
                    
                    try
                    {
                        String peekLine = sc.PeekLine();
                        if (line.Substring(0, line.IndexOf(":")).Equals("xml"))
                        { //get regulatory body
                            dc_id = "";
                            f = false;
                        }
                        else
                        {
                            if (line.Substring(0, line.IndexOf(":")).Equals("citeForThisResource"))
                            { //get regulatory body
                                reg_body = peekLine.Substring(peekLine.IndexOf(":") + 1).Trim();
                                start = row;
                                f = true;
                            }
                            if (line.Substring(0, line.IndexOf(":")).Equals("citation"))
                            {
                                cite_normalized = GetNthAttribute(line, 3).Substring(GetNthAttribute(line, 3).IndexOf("= ") + 1).Trim();
                                cite_type = GetNthAttribute(line, 1).Substring(GetNthAttribute(line, 1).IndexOf("= ") + 1).Trim();
                                cite_id = GetNthAttribute(line, 2).Substring(GetNthAttribute(line, 2).IndexOf("= ") + 1).Trim();
                            }
                            if (line.Substring(0, line.IndexOf(":")).Equals("content") && line.Substring(line.IndexOf(":")).Length > 5)
                            { //get regulatory body
                                //Console.WriteLine(line);
                                int pos1 = line.IndexOf("src = cid:") + 10;
                                int pos2 = line.Substring(pos1).IndexOf(" ]");
                                //Console.WriteLine(pos1 + "length = " + pos2);
                                content_id = line.Substring(pos1, pos2).Trim();//, line.IndexOf(" ]")- line.IndexOf("src = ")).Trim();
                                f = true;
                            }
                        }
                        if (line.Contains("dc:identifier:[ identifierScheme = LNI ]") || line.Contains("dc:date:"))
                        {
                            dc_id += peekLine.Substring(peekLine.IndexOf(":") + 1).Trim() + "; ";
                            //Console.WriteLine("id = " + dc_id.Substring(0, dc_id.IndexOf(";") - 1));
                            //Console.WriteLine("date = " + dc_id.Substring(dc_id.IndexOf(";") + 2).Replace(";", ""));
                        }
                        if (line.Contains(" [ "))  //if attributes exist
                        {
                            value = line.Substring(line.IndexOf(":") + 1).Trim();
                            //Console.WriteLine("     Attribute @" + col + " > " + value);
                        }
                        if (peekLine.Substring(0, peekLine.IndexOf(":")).Contains("#text"))
                        {
                            value = peekLine.Substring(peekLine.IndexOf(":") + 1).Trim();
                            //Console.WriteLine("     Text @" + col + " > " + value);
                        }
                        if (!value.Equals("") && !value.Equals(null))
                        {
                            Write((row-1).ToString(), 1, row); //row #
                            //Write(Path.GetFileNameWithoutExtension(file), 2, row); // source file
                            if (f)
                            {
                                int count = 0;
                                foreach(char c in dc_id)
                                {
                                    if (c == ';') count++;
                                }
                                if(count >= 2){
                                    //Write(dc_id, 3, row);   //content id
                                    Write(dc_id.Substring(0, dc_id.IndexOf(";") - 1).Trim(), 2, row);   //content id
                                    Write(dc_id.Substring(dc_id.IndexOf(";") + 2).Replace(";", "").Trim(), 3, row);   //content id
                                    if (!prop)
                                    {
                                        //prop = Propagate(dc_id, 3, row, start);
                                        prop = Propagate(dc_id.Substring(0, dc_id.IndexOf(";") - 1).Trim(), 2, row, start);
                                        prop = Propagate(dc_id.Substring(dc_id.IndexOf(";") + 2).Replace(";", "").Trim(), 3, row, start);
                                    }
                                }
                                
                                Write(reg_body, 4, row); // parent node
                                Write(cite_id, 5, row); //subcontent id
                                Write(cite_type, 6, row); //subcontent type
                                Write(cite_normalized, 7, row); //subcontent normalized
                            }
                            
                            Write(line.Substring(0, line.IndexOf(":")), 8, row); // node name
                            Write(System.Text.RegularExpressions.Regex.Replace(value, @"\s+", " "), 9, row); // node text/attribute(s)
                            
                            row++;
                        }
                    }
                    catch (Exception e) { };
                }
                r.Close();
                p.SaveAs(new FileInfo(dir + "narrow_output.xlsx"));
            }
        }

        private String GetNthAttribute(String value, int num)
        {
            int start = GetNthIndex(value, '[', num) + 1;
            int end = GetNthIndex(value, ']', num) - 1;
            return value.Substring(start, end - start);
            
        }

        public int GetNthIndex(String s, char t, int n)
        {
            int count = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == t)
                {
                    count++;
                    if (count == n)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        private Boolean Propagate (String value, int column, int current, int until)
        {
            if (current > until) //upwards
            {
                while (current >= until)
                {
                    Write(value, column, current);
                    current--;
                }
            }
            else
            {
                while (current <= until)
                {
                    Write(value, column, current);
                    current++;
                }
            }
            return true;
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
            sr.Close();
            return "";
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
