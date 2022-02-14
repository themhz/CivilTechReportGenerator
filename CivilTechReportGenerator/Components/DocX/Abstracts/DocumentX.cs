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

namespace CivilTechReportGenerator {
    abstract class DocumentX : IDocument {

        public int x, y, pos = 0;

        public RichEditDocumentServer srv;                        

        public String templatePath { get; set; }


        public DocumentX(RichEditDocumentServer _srv ) {
            this.srv = _srv;            
        }



        public void loadTemplate(String template) {
           
            this.srv.LoadDocument(template);          
            this.templatePath = template;

        }

        public void saveDocument() {
            this.srv.Document.EndUpdate();
            this.srv.SaveDocument(this.templatePath, DocumentFormat.OpenXml);
        }

        public void saveDocument(String generatedfile) {
            this.srv.SaveDocument(generatedfile, DocumentFormat.OpenXml);
        }

        public void beginUpdate() {
            this.srv.Document.BeginUpdate();
        }
        

    }
}
