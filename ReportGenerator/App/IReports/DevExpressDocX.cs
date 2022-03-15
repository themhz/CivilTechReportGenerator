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

    class DevExpressDocX : IReport {

        public RichEditDocumentServer wordProcessor { set; get; }
        public RichEditDocumentServer tempWordProcessor { set; get; }
        public IDataSource datasource { set; get; }
        public String template { set; get; }
        public String generatedfile { set; get; }
        private DocumentRange sourceRange { set; get; }
        private DocumentRange targetRange { set; get; }
        public DevExpressDocX(RichEditDocumentServer _wordProcessor, IDataSource _datasource) {
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
                this.parseFieldTypes(field, comment);

                //then move to the next comment..

            }
            Regex r = new Regex("{PBR}");
            wordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
            
            //Console.ReadLine();       
        }
        public void populatePageADetails(XmlNodeList DetailList, Table table, JObject jo, string id) {            
            foreach (XmlNode node in DetailList) {                
                Dictionary<String, String> cols = new Dictionary<string, string>();
                foreach (XmlNode row in node) {
                    cols.Add(row.Name, row.InnerText);
                }
                this.addTableRow(table, cols, jo);                
            }
            this.populateTableTotals(DetailList, table,jo, id);
        }
        public void replaceTextWithNewTable(string text, int rows, int cols) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(text);
            this.wordProcessor.Document.Tables.Create(this.targetRange.Start, rows, cols);
            this.delete();
        }
        public void replaceRangeWithNewText(DocumentRange sourceRange, string targetText) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = sourceRange;
            if (this.targetRange != null)
                this.wordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithNewText(string sourceText, string targetText) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(sourceText);
            if (this.targetRange != null)
                this.wordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithImage(DocumentRange sourceRange, string targetText, string id) {
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.Document.Unit = DevExpress.Office.DocumentUnit.Inch;
            this.targetRange = sourceRange;
            if (this.targetRange != null) {
                
                byte[] bytes = Convert.FromBase64String(this.datasource.GetValue(targetText).ToString());
                bytes = ImageResizer.resize(bytes, 700, 700);
                using (MemoryStream ms = new MemoryStream(bytes)) {
                    DocumentImageSource image = DocumentImageSource.FromStream(ms);
                    this.wordProcessor.Document.Images.Insert(this.targetRange.Start, image);
                }
                this.delete();
            }
            //this.wordProcessor.Document.Replace(targetRange, targetText);
        }        
        //Get the table on the document based on the index
        public Table getTable(int position) {
            return this.wordProcessor.Document.Tables[position];
        }
        public Paragraph getParagraph(int position) {
            return this.wordProcessor.Document.Paragraphs[position];
        }
        public DocumentRange getTextRange(string search) {
            try {
                Regex myRegEx = new Regex(search);
                return this.wordProcessor.Document.FindAll(myRegEx).First();
            } catch (Exception ex) {
                return null;
            }

        }
        public void delete() {
            this.wordProcessor.Document.Delete(this.targetRange);
        }                
        private void addTableRow(Table targetTable, Dictionary<string, string> cols, JObject jo) {
            int rowcount = targetTable.Rows.Count() - 1;
            targetTable.Rows.InsertAfter(rowcount);
            var colvalues = jo.GetValue("cols");
            if(colvalues != null) {
                for(int i = 0; i < colvalues.Count(); i++) {
                    this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, i].Range.Start, cols[colvalues[i].ToString()].ToString());
                }
                
            }            
        }        
        private void populateTableTotals(XmlNodeList DetailList, Table table, JObject jo, string id) {
            int rowcount = table.Rows.Count() - 1;
            var totals = jo.GetValue("total");
            //table.Rows.InsertAfter(rowcount);

            
            
            int index = -1;
            for (int i = 0; i < totals.Count(); i++) {
                for (int j=0;j< jo.GetValue("cols").Count(); j++) {                
                    if (totals[i]["col"].ToString() == jo.GetValue("cols")[j].ToString()) {
                        var value = this.datasource.GetValueByID("PageA."+ totals[0]["field"].ToString().Trim(), id);
                        this.wordProcessor.Document.InsertSingleLineText(table[rowcount, j].Range.Start, value);
                    }
                }
            }

            //for (int i = 0; i < totals.Count(); i++) {
            //    this.wordProcessor.Document.InsertSingleLineText(table[rowcount, Int32.Parse(totals[i].ToString())].Range.Start, "3");
            //}

        }
        //Routed to different elements depending on the type
        private void parseFieldTypes(string field, Comment comment, string id="") {
            JObject jo = JObject.Parse(field);                        
            switch (jo.GetValue("type").ToString()) {
                case "field":
                    this.parseField(jo, comment, id);
                    break;
                case "image":
                    this.parseImage(jo, comment, id);
                    break;
                case "list":
                    this.parseList(jo, comment, id);
                    break;
                case "template":
                    if (jo.ContainsKey("recursion")) {
                        this.parseTemplateRecursively(jo, comment, id);
                    } else {
                        this.parseTemplate(jo, comment, id);
                    }
                    break;
                case "table":
                    this.parseTable(jo, comment, id);
                    break;
            }

        }
        private void replaceTextWithTemplate(Comment comment, string file) {
            string documentTemplate = Path.Combine(ConfigurationManager.AppSettings["templates"] + file + ".docx");

            using (RichEditDocumentServer subWordPrecessor = new RichEditDocumentServer()) {
                subWordPrecessor.LoadDocumentTemplate(documentTemplate);

                this.tempWordProcessor = this.wordProcessor;
                this.wordProcessor = subWordPrecessor;
                foreach (Comment c in subWordPrecessor.Document.Comments.ToList()) {
                    SubDocument doc = c.BeginUpdate();
                    string field = doc.GetText(doc.Range).Replace("”", "\"").Replace("“", "\"");
                    this.parseFieldTypes(field, c);
                }
                this.wordProcessor = this.tempWordProcessor;
                this.wordProcessor.Document.InsertDocumentContent(comment.Range.End, subWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
            }
            this.wordProcessor.Document.Delete(comment.Range);
        }
        private void replaceTextWithTemplate(Comment comment, string file, string id) {
            string documentTemplate = Path.Combine(ConfigurationManager.AppSettings["templates"] + file + ".docx");

            using (RichEditDocumentServer subWordPrecessor = new RichEditDocumentServer()) {
                subWordPrecessor.LoadDocumentTemplate(documentTemplate);

                this.tempWordProcessor = this.wordProcessor;
                this.wordProcessor = subWordPrecessor;
                foreach (Comment c in subWordPrecessor.Document.Comments.ToList()) {
                    SubDocument doc = c.BeginUpdate();
                    String field = doc.GetText(doc.Range).Replace("”", "\"").Replace("“", "\"");

                    if(field!="")
                        this.parseFieldTypes(field, c, id);
                }
                this.wordProcessor = this.tempWordProcessor;
                this.wordProcessor.Document.InsertDocumentContent(comment.Range.End, subWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
            }
            
        }
        public void parseField(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is field");
            this.replaceRangeWithNewText(comment.Range, datasource.GetValue(jo.GetValue("name").ToString()).ToString());
        }
        public void parseList(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is List");
        }
        public void parseTemplate(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is template");
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString());
        }        
        public void parseImage(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is image");
            XmlNodeList DetailList = ((Xml)datasource).getList("PageA[ns:ID='" + id + "']");

            //this.replaceTextWithImage(comment.Range, jo.GetValue("name").ToString(), id);
            this.replaceTextWithImage(comment.Range, DetailList[0]["Image"].InnerText, id);            
        }      
        public void parseTable(JObject jo, Comment comment, string id = "") {            
            string data = jo.GetValue("cols").ToString().Replace("[", " ").Replace("]", " ").Replace(Environment.NewLine, "");            
            TableCell tableCell = this.wordProcessor.Document.Tables.GetTableCell(comment.Range.Start);

            if(tableCell != null) {
                XmlNodeList DetailList = ((Xml)datasource).getList("PageADetails[ns:PageADetailID='" + id + "']");
                this.populatePageADetails(DetailList, tableCell.Table, jo, id);                              
            }
        }
        public void parseTemplateRecursively(JObject jo, Comment comment, string id="") {

            string recursiveElement = jo.GetValue("recursion").ToString();
            XmlNodeList PageAList = ((Xml)datasource).getList(recursiveElement);
            foreach (XmlNode page in PageAList) {                
                this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), page["ID"].InnerText);
            }
            this.wordProcessor.Document.Delete(comment.Range);
        }
    }
}
