using System;
using System.IO;
using System.Xml;

namespace xml_converter
{
    class Program
    {
        public static void Main(string[] args)
        {
            //String dir = "C:/Users/Rebecca Ramnauth/source/repos/xml_converter/xml_converter/test_files/";
            String dir = "test_files/";

            Reader r = new Reader(dir);
            r.ReadDir();  //generated directory 
             

            StreamWriter s = null;
            String finders_dir = "test_files/generated_files/"; // r.ReadDir()'s new dir
            foreach(String file in Directory.EnumerateFiles(finders_dir, "*.xml"))
            {
                try
                {
                    s = new StreamWriter("test_files/content_files/" + Path.GetFileNameWithoutExtension(file) + "_content.txt");
                    Console.WriteLine("Reading " + Path.GetFileNameWithoutExtension(file));

                    XmlDocument d = new XmlDocument();
                    d.Load(file);

                    Finder nf = new Finder(file, d);
                    //nf.Beautify(); //beautified directory
                    nf.TraverseNodes(d.ChildNodes, s);  //content directory

                    s.Close();
                }
                catch (Exception e) { }
            } 

            //String h = "C:/Users/Rebecca Ramnauth/source/repos/xml_converter/xml_converter/test_files/meta_data/unique_tags.txt";
            Writer w = new Writer(dir);
            w.NarrowHeaders("");

            //Cleaner c = new Cleaner(dir + "output.xlsx");
        }
    }
}
