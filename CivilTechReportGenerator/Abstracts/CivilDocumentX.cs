using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace CivilTechReportGenerator {
    abstract class CivilDocumentX {

        public RichEditDocumentServer srv;        
        
        public Document document { get; set; }

        public String templatePath { get; set; }


        public CivilDocumentX() {
            this.srv = new RichEditDocumentServer();
        }



        public void loadTemplate(String template) {
            //using (this.srv) {
                this.srv.LoadDocument(template);
                this.document = this.srv.Document;                
            //}

            this.templatePath = template;

        }

        public void saveDocument() {
            this.srv.SaveDocument(this.templatePath, DocumentFormat.OpenXml);

        }

    }
}
