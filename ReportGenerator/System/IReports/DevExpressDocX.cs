using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using ReportGenerator;
using ReportGenerator.Interfaces.Elements;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ReportGenerator_v1.System {

    class DevExpressDocX : IReport {

        public RichEditDocumentServer wordProcessor;
        public String template { set; get; }
        public String generatedfile { set; get; }



        public DevExpressDocX(RichEditDocumentServer _wordProcessor) {
            this.wordProcessor = _wordProcessor;
        }
        public IReport create() {
            Console.WriteLine("creating file report");
            using (this.wordProcessor) {
                this.load();
                this.parse();
                this.save();
            }
            Console.WriteLine("file report created");
            return this;
        }

        public void load() {
            this.wordProcessor.LoadDocument(this.template);
        }

        public void save() {
            Console.WriteLine("Saving file report");            
            this.wordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);            
            Console.WriteLine("Report save in :"+ this.template);            
        }


        public void parse() {
            this.replaceTextWithNewTable("{table1}", 2,3);
        }

        
        public void replaceTextWithNewTable(String text, int rows, int cols) {
            this.wordProcessor.Document.BeginUpdate();
            DocumentRange documentRange = this.getTextRange(text);
            this.wordProcessor.Document.Tables.Create(documentRange.Start, rows, cols);
            this.wordProcessor.Document.Delete(documentRange);
        }
        public void delete(DocumentRange documentRange) {
            this.wordProcessor.Document.Delete(documentRange);
        }

        protected DocumentRange getElementRange() {
            return null;
        }

        public DocumentRange getTableRange(int position) {
           return this.wordProcessor.Document.Tables[position].Range;            
        }

        public DocumentRange getParagrapgRange(int position) {
            return this.wordProcessor.Document.Paragraphs[position].Range;
        }

        public DocumentRange getTextRange(String search) {
            Regex myRegEx = new Regex(search);
            return this.wordProcessor.Document.FindAll(myRegEx).First();
        }

        public void delete() {
            throw new NotImplementedException();
        }
    }
}
