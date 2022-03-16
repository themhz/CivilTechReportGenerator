using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using ReportGenerator;
using ReportGenerator.DataSources;
using ReportGenerator.Helpers;
using ReportGenerator_v1.DataSources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Text.Json.Nodes;

namespace ReportGenerator_v1.System {

    class DevExpressDocX002 : IReport {

        public RichEditDocumentServer wordProcessor { set; get; }
        public RichEditDocumentServer tempWordProcessor { set; get; }
        public IDataSource datasource { set; get; }
        public String template { set; get; }
        public String generatedfile { set; get; }
        private DocumentRange sourceRange { set; get; }
        private DocumentRange targetRange { set; get; }
        public DevExpressDocX002(RichEditDocumentServer _wordProcessor, IDataSource _datasource) {
            this.wordProcessor = _wordProcessor;
            this.datasource = _datasource;
        }
        public IReport create() {
            Console.WriteLine("creating file report");
            using (this.wordProcessor) {
                this.load();
                this.parse();
                this.save();
                this.openfile();
            }
            Console.WriteLine("file report created");
            return this;
        }
        public void load() {
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.LoadDocument(this.template);
        }
        public void save() {
            this.wordProcessor.Document.EndUpdate();
            Console.WriteLine("Saving file report");
            this.wordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);
            Console.WriteLine("Report save in :" + this.template);
        }
        public void openfile() {
            Process.Start(new ProcessStartInfo(this.generatedfile) { UseShellExecute = true });
        }
        //This function is under construction. It will be used to parse the word template document and conscrtruct the report
        //However some commans are implemented.. more to come..
        public void parse() {
            //Collect all comments and loop through them in order to parse each one of them individually and create the document elemtnts
            CommentCollection comments = this.wordProcessor.Document.Comments;
            foreach (Comment comment in comments.ToList()) {
                //this block of code is used to get the comments and save them in a string variable
                //all the block is needed according to devexpress
                SubDocument doc = comment.BeginUpdate();
                String field = doc.GetText(doc.Range).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");
                comment.EndUpdate(doc);
                
                //After collecting the individual comment witch is in json format, it will be passed to the internal function parseFieldTypes in order to 
                //parse and create the fields
                //this.parseFieldTypes(field, comment);

                //then move to the next comment..

            }
            Regex r = new Regex("{PBR}");
            wordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
            
            //Console.ReadLine();       
        }

        public void delete() {
            throw new NotImplementedException();
        }
    }
}
