using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;

namespace GenerateDocxDemo
{
    /// <summary>
    /// Summary description for DocumentSpecification
    /// </summary>
    public class DocumentSpecification
    {
        public DocumentSpecification()
        {

        }

        public string TextContent { get; set; }

        public byte[] DocumentToMerge { get; set; }

        public string HTMLContent { get; set; }
    }
}
