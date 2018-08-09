using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace xml_converter
{
    class Reader
    {
        private static String dir;

        public Reader(String d)
        {
            dir = d;
        }
        
        // read all content files
        public void ReadDir()
        {
            List<List<String>> tags = new List<List<String>>();
            StreamWriter heads = new StreamWriter (dir + "/meta_data/meta.txt");
            foreach (String file in Directory.EnumerateFiles(dir, "*.txt"))
            {
                tags.Add(ReadFile(file, heads));
            }

            List<String> flatList = tags.SelectMany(v => v).ToList();

            StreamWriter all_tags = new StreamWriter(dir + "/meta_data/all_tags.txt");
            StreamWriter unique_tags = new StreamWriter(dir + "/meta_data/unique_tags.txt");
            
            all_tags.WriteLine(string.Join("\r", flatList.ToArray()));
            unique_tags.WriteLine(string.Join("\r", flatList.Distinct().ToArray()));

            all_tags.Close();
            unique_tags.Close();
            heads.Close();
        }

        // gathers meta for schema:: ?attributes
        public List<String> Tags (String file)
        {
            List<String> tags = new List<String>();
            try { 
                XmlDocument f = new XmlDocument();
                f.Load(file);
                XmlNodeList entries = f.GetElementsByTagName("*");
                
                //StreamWriter all_tags = new StreamWriter("all_tags.txt");
                foreach (XmlNode n in entries)
                {
                    XmlElement element = (XmlElement)n;
                    //Console.WriteLine(n.Name.ToString());
                    tags.Add(n.Name);
                }
            } catch (Exception e) {
                String msg = Path.GetFileName(file) + ": " + e;
                Console.WriteLine(msg);
            }
            return tags;
        }

        // flatten nested xml format with meta while perserving hierarchy
        private List<String> ReadFile(String file, StreamWriter heads)
        {
            String source = file;
            int i = 0;

            String content = File.ReadAllText(source);
            String[] contents = content.Split("--yytet00pubSubBoundary00tetyy");
            //Console.WriteLine("contents.length = " + contents.Length);
            String destination = "";
            StreamWriter dest = null;

            String header = "";
            String description = "";

            //StreamWriter heads = new StreamWriter(Path.GetDirectoryName(source) + "/meta_data/meta" + ".txt");
            foreach (String c in contents)
            {
                destination = Path.GetFileNameWithoutExtension(file) + "-" + i++;
                Console.Write(">" + source);
                int positionOfXML = c.IndexOf("<?xml");
                //Console.WriteLine(">" + positionOfXML);
                Console.Write(" > " + destination + "\n");
                try
                {
                    header += destination + ".txt " + c.Substring(0, positionOfXML);
                    description = c.Substring(positionOfXML);
                    //heads.Write("\n" + header);
                    //Console.WriteLine("     + " + header);
                }
                catch (Exception e) { }

                dest = new StreamWriter(Path.GetDirectoryName(source) + "/generated_files/" + destination + ".xml");
                dest.WriteLine(description);
                dest.Close();
                //heads.Close();
            }
            heads.Write(header);
            dest.Close();

            return Tags(Path.GetDirectoryName(source) + "/generated_files/" + destination + ".xml");
        }

        // read xml file as is
        static void Read(String path)
        {
            String line;
            try
            {
                StreamReader sr = new StreamReader(path);
                line = sr.ReadLine();

                while(line != null)
                {
                    Console.WriteLine(line);
                    line = sr.ReadLine();
                }
                sr.Close();
                Console.ReadLine();
            }
            catch(Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Completed Read(path);");
            }
        }

        // make ugly xmls beautiful
        public void Beautify(String path)
        {
            XmlReader r = XmlReader.Create(path);
            while (r.NodeType != XmlNodeType.Element)
                r.Read();
            XElement e = XElement.Load(r);
            //Console.WriteLine(e);
            e.Save(Path.GetFileNameWithoutExtension(path) + "_beautified.xml");
        }
        
    }
}
