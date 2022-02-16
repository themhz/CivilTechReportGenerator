using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportGenerator.Interfaces;
using ReportGenerator.Types;

namespace ReportGenerator.Handlers {
    public class DocumentHandler : DocumentX, IDocumentHandler {

        public String text { get; set; }
        public DocumentXItem item { get; set; }

        public DocumentHandler setDocumentItem(DocumentXItem _item) {
            item = _item;
            return this;
        }
        public DocumentXItem getDocumentItem() {
            return item;
        }

        public DocumentHandler(RichEditDocumentServer wordProcessor) : base(wordProcessor) {

        }

        public override int count() {
            return 1;
        }

        public String scanDocument() {
            Document document = base.wordProcessor.Document;
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


            //TableItem th = new TableItem(base.srv);
            //th.loadTemplate(base.templatePath);
            //SectionItem sh = new SectionItem(base.srv);
            //sh.loadTemplate(base.templatePath);
            //ParagraphItem ph = new ParagraphItem(base.srv);
            //ph.loadTemplate(base.templatePath);
            //ListItem lh = new ListItem(base.srv);
            //lh.loadTemplate(base.templatePath);

            //elementCounter.Add(new String[] { "Tables:", th.count().ToString()});
            //elementCounter.Add(new String[] { "Sections:", sh.count().ToString() });
            //elementCounter.Add(new String[] { "Paragraphs:", ph.count().ToString() });
            //elementCounter.Add(new String[] { "List:", lh.count().ToString() });
            return elementCounter;
        }


        public void deleteElement(DocumentXItem item, int index) {
            item.delete(index);
        }
    }
}
