using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CivilTechReportGenerator.Interfaces;
using CivilTechReportGenerator.Types;

namespace CivilTechReportGenerator.Handlers {
    class DocumentHandler : CivilDocumentX {
        
        public String text;

        public DocumentHandler(RichEditDocumentServer wordProcessor) : base() {
            base.srv = wordProcessor;
        }

        public override int count() {
            return 1;
        }

        public String scanDocument() {
            Document document = base.srv.Document;
            DocumentIterator iterator = new DocumentIterator(document, true);
            MyVisitor visitor = new MyVisitor();            

            while (iterator.MoveNext()) {
                String test = iterator.Current.Type.ToString();                                              
                if (test.Equals("Text"))
                    visitor.Buffer.Append(test + "=");
                else
                    visitor.Buffer.AppendLine(test);

                iterator.Current.Accept(visitor);
                



            }                                    

            return visitor.Text;
        }


        public override void create() {                     

        }

        public List<String[]> countElements() {
            List<String[]> elementCounter = new List<String[]>();


            TableHandler th = new TableHandler(base.srv);
            th.loadTemplate(base.templatePath);
            SectionHandler sh = new SectionHandler(base.srv);
            sh.loadTemplate(base.templatePath);
            ParagraphHandler ph = new ParagraphHandler(base.srv);
            ph.loadTemplate(base.templatePath);
            ListHandler lh = new ListHandler(base.srv);
            lh.loadTemplate(base.templatePath);

            elementCounter.Add(new String[] { "Tables:", th.count().ToString()});
            elementCounter.Add(new String[] { "Sections:", sh.count().ToString() });
            elementCounter.Add(new String[] { "Paragraphs:", ph.count().ToString() });
            elementCounter.Add(new String[] { "List:", lh.count().ToString() });
            return elementCounter;
        }
        
        
    }
}
