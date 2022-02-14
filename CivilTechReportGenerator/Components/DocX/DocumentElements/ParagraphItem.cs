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
    class ParagraphItem : DocumentX {
        
        public String text;

        public ParagraphItem(RichEditDocumentServer wordProcessor) : base(wordProcessor) {

        }
        

        public override void create() {
           
            Document document = base.srv.Document;
            // Start the document update:
            document.BeginUpdate();

            // Append a paragraph:
            Paragraph appendedParagraph = document.Paragraphs.Append();
            document.InsertText(appendedParagraph.Range.Start, this.text);

            //// Insert a paragraph at the start of the second section:
            //Paragraph paragraph = document.Paragraphs.Insert(document.Sections[1].Range.Start);
            //DocumentPosition position = document.Paragraphs[paragraph.Index - 1].Range.Start;
            //document.InsertText(position, "Inserted paragraph");

            // Finalize the document update:
            document.EndUpdate();

            //DocumentPosition position = document.CreatePosition(_pos);


            base.saveDocument();
            
        }

        public void replace(String generatedfile, int pos, Regex _myRegEx) {

            var myRegEx = _myRegEx;
            DocumentRange dr = this.srv.Document.FindAll(myRegEx).First();
            DocumentPosition dpos = this.srv.Document.CreatePosition(dr.Start.ToInt());
            this.srv.Document.InsertText(dpos, " ");
            this.srv.Document.InsertDocumentContent(dpos, this.srv.Document.Sections[pos].Range);
            this.srv.Document.Delete(dr);
            base.saveDocument(generatedfile);
        }

        public void delete(int index, String generatedfile) {
            this.srv.Document.Delete(this.srv.Document.Paragraphs[index].Range);
            base.saveDocument(generatedfile);
        }


        public override int count() {
            return this.srv.Document.Paragraphs.Count;
        }

    }
}
