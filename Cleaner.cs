using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace xml_converter
{
    class Cleaner
    {
        private ExcelPackage p;

        public Cleaner(String path)
        {
            FileInfo fpath = new FileInfo(path);
            if (fpath.Exists)
            {
                p = new ExcelPackage(fpath);
            }
            else
            {
                Console.WriteLine("File does not exist. Execute a Writer to create a content xlsx");
            }
        }

        // ?values to propagate down (exceptions ∈ p/?desig/...)
        public void Propagate(String[] values)
        {
            // propagate downwards until new value is found
        }

    }
}
