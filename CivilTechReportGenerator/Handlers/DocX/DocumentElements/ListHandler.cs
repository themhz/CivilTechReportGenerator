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
    class ListHandler : CivilDocumentX {
        
        public String text;

        public ListHandler(RichEditDocumentServer wordProcessor) : base() {
            base.srv = wordProcessor;
        }
        

        public override void create() {                       
            
        }

   

        public override int count() {
            return base.document.NumberingLists.Count();
        }

    }
}
