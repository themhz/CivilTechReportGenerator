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
    class ListItem : IDocumentItem  {
        
        public String text;
        public RichEditDocumentServer _wordProcessor;

        public ListItem(RichEditDocumentServer wordProcessor) {
            _wordProcessor = wordProcessor;

        }
        

        public override void create() {                       
            
        }

   

        public override int count() {
            return _wordProcessor.Document.NumberingLists.Count();
        }

    }
}
