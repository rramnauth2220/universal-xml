using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace xml_converter
{
    class Finder
    {
        private static XmlDocument doc;
        private static String path;

        public Finder(String p, XmlDocument d)
        {
            path = p;
            doc = d;
        }

        // make ugly xmls beautiful
        public void Beautify()
        {
            XmlReader r = XmlReader.Create(path);
            while (r.NodeType != XmlNodeType.Element)
                r.Read();
            XElement e = XElement.Load(r);
            //Console.WriteLine(e);
            e.Save("C:/Users/Rebecca Ramnauth/source/repos/xml_converter/xml_converter/test_files/beautified_files/" + Path.GetFileNameWithoutExtension(path) + "_beautified.xml");
        }
       
        // print all/unique tags from document
        public void Tags()
        {
            XmlNodeList entries = doc.GetElementsByTagName("*");
            StreamWriter all_tags = new StreamWriter("all_tags.txt");
            String[] nodes = new String[entries.Count];

            for(int i = 0; i < entries.Count; i++)
            {
                XmlElement element = (XmlElement)entries.Item(i);
                nodes[i] = element.Name;
                all_tags.WriteLine(element.Name);
            }

            String[] distNodes = nodes.Distinct().ToArray();
            StreamWriter unique_tags = new StreamWriter("unique_tags.txt");

            for(int j = 0; j < distNodes.Length; j++)
            {
                unique_tags.WriteLine(distNodes[j]);
            }

            unique_tags.Close();
            all_tags.Close();
        }

        // visits nodes and collects all information
        public void TraverseNodes(XmlNodeList nodes, StreamWriter s)
        {
            //XmlNodeList nodes = node.ChildNodes;
            foreach (XmlNode node in nodes)
            {
                s.Write((node.Name + ": " + node.Value).Trim());

                XmlAttributeCollection attributes = node.Attributes;
                try
                {   // get attributes
                    String values = "";
                    foreach (XmlAttribute attribute in attributes)
                    {
                        values += "[ " + attribute.Name + " = " + attribute.Value + " ] ";
                        //s.Write("[ " + attribute.Name + " = " + attribute.Value + " ] ");
                        //s.WriteLine("\t" + attribute.Name + ": " + attribute.Value);
                    }
                    s.Write(values);
                }
                catch (Exception e) { } //tag has no attributes
                s.Write("\n");
                TraverseNodes(node.ChildNodes, s);
            }
        }
    }
}
