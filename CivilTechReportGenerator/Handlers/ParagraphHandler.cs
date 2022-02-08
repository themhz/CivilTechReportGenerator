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

namespace CivilTechReportGenerator.Handlers {
    class ParagraphHandler : CivilDocumentX {
        
        public String text;

        public ParagraphHandler(RichEditDocumentServer wordProcessor) : base() {
            base.srv = wordProcessor;
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
        
    }
}
