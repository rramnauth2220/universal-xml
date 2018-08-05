using System;
using System.Collections.Generic;
using System.IO;

namespace xml_converter
{
    class Scanner
    {
        private readonly StreamReader reader;
        private readonly Queue<String> readAheadQueue = new Queue<String>();

        public Scanner(StreamReader r)
        {
            reader = r;
        }

        // read line
        public String ReadLine()
        {
            if (readAheadQueue.Count > 0)
            {
                return readAheadQueue.Dequeue();
            }
            return reader.ReadLine();
        }

        // peek next line
        public String PeekLine()
        {
            if (readAheadQueue.Count > 0)
            {
                return readAheadQueue.Peek();
            }

            String line = reader.ReadLine();

            if (line != null)
            {
                readAheadQueue.Enqueue(line);
            }
            return line;
        }
    }
}
