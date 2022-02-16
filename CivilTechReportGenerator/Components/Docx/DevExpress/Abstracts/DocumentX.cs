using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportGenerator.Interfaces;

namespace ReportGenerator {
    public abstract class DocumentX : IDocumentX {      

        public RichEditDocumentServer wordProcessor;                        

        public String templatePath { get; set; }


        public DocumentX(RichEditDocumentServer _wordProcessor) {
            this.wordProcessor = _wordProcessor;            
        }

        public void loadTemplate(String template) {
           
            this.wordProcessor.LoadDocument(template);          
            this.templatePath = template;

        }

        public void saveDocument() {
            this.wordProcessor.Document.EndUpdate();
            this.wordProcessor.SaveDocument(this.templatePath, DocumentFormat.OpenXml);
        }

        public void saveDocument(String generatedfile) {
            this.wordProcessor.SaveDocument(generatedfile, DocumentFormat.OpenXml);
        }

        public void beginUpdate() {
            this.wordProcessor.Document.BeginUpdate();
        }
        

    }
}
