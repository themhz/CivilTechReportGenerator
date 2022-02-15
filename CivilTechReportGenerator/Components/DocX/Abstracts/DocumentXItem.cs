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
using CivilTechReportGenerator.Handlers;

namespace CivilTechReportGenerator {
    public abstract class DocumentXItem : IDocumentXItem {
        
        public RichEditDocumentServer wordProcessor;

        public DocumentXItem() {            
        }    

        public int documentPosition;
        public DocumentXItem setDocumentPosition(int val) {
            documentPosition = val;
            return this;
        }

        public void createSpace(int posTarget) {
            DocumentPosition dpos = this.wordProcessor.Document.CreatePosition(posTarget);
            this.wordProcessor.Document.InsertText(dpos, " ");
        }

    }
}
