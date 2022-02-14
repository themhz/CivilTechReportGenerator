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
    class ListItem : IDocumentXItem, IListItem {

        public String text;
        public RichEditDocumentServer wordProcessor;

        public ListItem(RichEditDocumentServer _wordProcessor) {
            wordProcessor = _wordProcessor;

        }


        public override void create() {

        }

        public override void delete(int index) {

            wordProcessor.Document.Paragraphs.RemoveNumberingFromParagraph(wordProcessor.Document.Paragraphs[index]);

        }



        public override int count() {
            return wordProcessor.Document.NumberingLists.Count();
        }

    }
}
