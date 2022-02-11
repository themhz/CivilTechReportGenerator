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
    abstract class CivilDocumentX :IDocumentItem  {

        public int x, y, pos = 0;

        public RichEditDocumentServer srv;        
        
        public Document document { get; set; }

        public String templatePath { get; set; }


        public CivilDocumentX() {
            this.srv = new RichEditDocumentServer();
        }



        public void loadTemplate(String template) {
           
            this.srv.LoadDocument(template);
            this.document = this.srv.Document;
            this.templatePath = template;

        }

        public void saveDocument() {
            this.document.EndUpdate();
            this.srv.SaveDocument(this.templatePath, DocumentFormat.OpenXml);
        }

        public void saveDocument(String generatedfile) {
            this.srv.SaveDocument(generatedfile, DocumentFormat.OpenXml);
        }

        public void beginUpdate() {
            this.document.BeginUpdate();
        }
        

    }
}
