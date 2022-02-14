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
    class DocumentHandler : DocumentX {
        
        public String text;

        //private ListItem listItem;
        //public DocumentHandler setSectionItem(ListItem _listItem) {
        //    listItem = _listItem;
        //    return this;
        //}
        //public ListItem getListItem() {
        //    return listItem;
        //}



        //private ParagraphItem paragraphItem;
        //public DocumentHandler setSectionItem(ParagraphItem _paragraphItem) {
        //    paragraphItem = _paragraphItem;
        //    return this;
        //}
        //public ParagraphItem getParagraphItem() {
        //    return paragraphItem;
        //}


        //private SectionItem sectionItem;
        //public DocumentHandler setSectionItem(SectionItem _sectionItem) {
        //    sectionItem = _sectionItem;
        //    return this;
        //}
        //public SectionItem getSectionItem() {
        //    return sectionItem;
        //}

        //private TableItem tableItem;
        //public DocumentHandler setTableItem(TableItem _tableItem) {
        //    tableItem = _tableItem;
        //    return this;
        //}
        //public TableItem getTableItem() {
        //    return tableItem;
        //}


        private DocumentXItem item;
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
