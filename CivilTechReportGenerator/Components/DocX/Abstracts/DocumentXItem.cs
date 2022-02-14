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
    abstract class DocumentXItem : IDocumentXItem {
        
        public RichEditDocumentServer srv;

        public DocumentXItem(RichEditDocumentServer _srv) {
            this.srv = _srv;
        }    

        public int documentPosition;
        public DocumentXItem setDocumentPosition(int val) {
            documentPosition = val;
            return this;
        }

        public void createSpace(int posTarget) {
            DocumentPosition dpos = this.srv.Document.CreatePosition(posTarget);
            this.srv.Document.InsertText(dpos, " ");
        }

    }
}
