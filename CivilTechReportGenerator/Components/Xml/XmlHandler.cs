using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CivilTechReportGenerator
{    
    class XmlHandler
    {
        public String path;
        public XmlDocument doc;
        public XmlHandler(String path , XmlDocument _doc )
        {
            this.doc = _doc;
            this.doc.Load(path);
        }

        public String getXml()
        {
            return this.doc.InnerXml;
        }
    }
}
