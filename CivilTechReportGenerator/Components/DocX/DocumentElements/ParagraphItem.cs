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
using System.Text.RegularExpressions;

namespace CivilTechReportGenerator.Handlers {
    class ParagraphItem : DocumentXItem {
        

        public String text;
        public RichEditDocumentServer srv;
        public ParagraphItem(RichEditDocumentServer srv) : base(srv) {
            this.srv = srv;
        }        
        public override void create() {
           
            Document document = base.srv.Document;
            // Start the document update:
            document.BeginUpdate();

            // Append a paragraph:
            Paragraph appendedParagraph = document.Paragraphs.Append();
            document.InsertText(appendedParagraph.Range.Start, this.text);

            // Finalize the document update:
            document.EndUpdate();

            //DocumentPosition position = document.CreatePosition(_pos);                    
        }
        public void replace(int pos, Regex _myRegEx) {

            var myRegEx = _myRegEx;
            DocumentRange dr = this.srv.Document.FindAll(myRegEx).First();
            DocumentPosition dpos = this.srv.Document.CreatePosition(dr.Start.ToInt());
            this.srv.Document.InsertText(dpos, " ");
            this.srv.Document.InsertDocumentContent(dpos, this.srv.Document.Sections[pos].Range);
            this.srv.Document.Delete(dr);            
        }
        public Table findParagraph(int pos) {
            return this.srv.Document.Tables[pos];
        }
        public override void delete(int index) {
            this.srv.Document.Delete(this.srv.Document.Paragraphs[index].Range);            
        }
        public override int count() {
            return this.srv.Document.Paragraphs.Count;
        }

    }
}
