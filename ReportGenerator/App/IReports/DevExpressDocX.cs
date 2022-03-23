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

        public RichEditDocumentServer mainWordProcessor { set; get; }
        public RichEditDocumentServer tempWordProcessor { set; get; }
        public IDataSource datasource { set; get; }
        public String template { set; get; }
        public String generatedfile { set; get; }
        private DocumentRange sourceRange { set; get; }
        private DocumentRange targetRange { set; get; }


        public DevExpressDocX(RichEditDocumentServer _wordProcessor, IDataSource _datasource) {
            this.mainWordProcessor = _wordProcessor;
            this.datasource = _datasource;
        }
        public IReport create() {
            Console.WriteLine("creating file report");
            using (this.mainWordProcessor) {
                this.load();
                this.parse();
                this.save();
                this.openfile();
            }
            Console.WriteLine("file report created");
            return this;
        }
        public void load() {
            this.mainWordProcessor.Document.BeginUpdate();
            this.mainWordProcessor.LoadDocument(this.template);
        }
        public void save() {
            this.mainWordProcessor.Document.EndUpdate();
            Console.WriteLine("Saving file report");
            this.mainWordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);
            Console.WriteLine("Report save in :" + this.template);
        }
        public void openfile() {
            Process.Start(new ProcessStartInfo(this.generatedfile) { UseShellExecute = true });
        }                                    
        private void addTableRow(Table targetTable, Dictionary<string, string> cols, JObject jo) {
            int rowcount = targetTable.Rows.Count() - 1;
            targetTable.Rows.InsertAfter(rowcount);
            var colvalues = jo.GetValue("cols");
            if(colvalues != null) {
                for(int i = 0; i < colvalues.Count(); i++) {
                    this.mainWordProcessor.Document.InsertSingleLineText(targetTable[rowcount, i].Range.Start, cols[colvalues[i].ToString()].ToString());

                }
                
            }            
        }        
        private void populateTableTotals(XmlNodeList DetailList, Table table, JObject jo, string id) {
            int rowcount = table.Rows.Count() - 1;
            var totals = jo.GetValue("total");                     
            
            int index = -1;
            for (int i = 0; i < totals.Count(); i++) {
                for (int j=0;j< jo.GetValue("cols").Count(); j++) {                
                    if (totals[i]["col"].ToString() == jo.GetValue("cols")[j].ToString()) {
                        var value = this.datasource.GetValueByID("PageA."+ totals[0]["field"].ToString().Trim(), id);
                        this.mainWordProcessor.Document.InsertSingleLineText(table[rowcount, j].Range.Start, value);
                    }
                }
            }         
        }        
        public void loopTable(JObject jo, Comment comment, string id = "", string foreignKey = "") {

            XmlNodeList tables = ((Xml)datasource).getList(jo.GetValue("loopTable").ToString(), jo.GetValue("foreignKey").ToString(), id);
            foreach (XmlNode table in tables) {
                this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), table[jo.GetValue("id").ToString()].InnerText, table[jo.GetValue("foreignKey").ToString()].InnerText);
            }
            this.mainWordProcessor.Document.Delete(comment.Range);
        }        
        public void populateTable(XmlNodeList DetailList, Table table, JObject jo, string id) {
            foreach (XmlNode node in DetailList) {
                Dictionary<String, String> cols = new Dictionary<string, string>();
                foreach (XmlNode row in node) {
                    cols.Add(row.Name, row.InnerText);
                }
                this.addTableRow(table, cols, jo);
            }
            this.populateTableTotals(DetailList, table, jo, id);
        }
        public void delete() {
            this.mainWordProcessor.Document.Delete(this.targetRange);
        }

        public void parse() {
            //Collect all comments and loop through them in order to parse each one of them individually and create the document elemtnts
            CommentCollection comments = this.mainWordProcessor.Document.Comments;
            foreach (Comment comment in comments.ToList()) {
                //this block of code is used to get the comments and save them in a string variable
                //all the block is needed according to devexpress
                string commentText = this.getCommentText(comment);
                //After collecting the individual comment witch is in json format, it will be passed to the internal function parseFieldTypes in order to 
                //parse and create the fields
                this.parseCommentTypes(commentText, comment);
                //then move to the next comment..
            }
            Regex r = new Regex("{PBR}");
            mainWordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
        }        
        private void parseCommentTypes(string field, Comment comment, string id = "", string foreignKey = "") {
            try {
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
                        this.parseTemplate(jo, comment, id);
                        break;
                    case "table":
                        this.parseTable(jo, comment, id);
                        break;
                    case "complexTable":
                        this.parseComplexTable(jo, comment, id, foreignKey);
                        break;
                }
            } catch (Exception ex) {

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
            if (jo.ContainsKey("loopTable")) {
                this.loopTable(jo, comment, id);
            } else {                
                this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id);
            }
            
        }        
        public void parseImage(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is image");
            //XmlNodeList DetailList = ((Xml)datasource).getList("PageA[ns:ID='" + id + "']");
            XmlNodeList DetailList = ((Xml)datasource).getList("PageA","ID", id);

            //this.replaceTextWithImage(comment.Range, jo.GetValue("name").ToString(), id);
            this.replaceTextWithImage(comment.Range, DetailList[0]["Image"].InnerText, id);            
        }      
        public void parseTable(JObject jo, Comment comment, string id = "") {            
            string data = jo.GetValue("cols").ToString().Replace("[", " ").Replace("]", " ").Replace(Environment.NewLine, "");            
            TableCell tableCell = this.mainWordProcessor.Document.Tables.GetTableCell(comment.Range.Start);

            if(tableCell != null) {
                //XmlNodeList DetailList = ((Xml)datasource).getList("PageADetails[ns:PageADetailID='" + id + "']");
                XmlNodeList DetailList = ((Xml)datasource).getList("PageADetails","PageADetailID", id);
                this.populateTable(DetailList, tableCell.Table, jo, id);                              
            }
        }
        public void parseComplexTable(JObject jo, Comment comment, string id = "", string foreignKey = "") {
            
            this.mainWordProcessor.Document.AppendText("newTable");
            var newTableRange = this.getTextRange("newTable");
            TableCell tableCell = null;
            Table table;

            int headerCount = Int32.Parse(jo.GetValue("headerCount").ToString());
            int rowCount = Int32.Parse(jo.GetValue("rowCount").ToString());
            int footerCount = Int32.Parse(jo.GetValue("footerCount").ToString());

            // Copy header
            tableCell = this.mainWordProcessor.Document.Tables.GetTableCell(comment.Range.Start);
            table = tableCell.Table;
            var headerRange = getRowsRange(table, 0, headerCount);
            mainWordProcessor.Document.InsertDocumentContent(newTableRange.End, headerRange, InsertOptions.KeepSourceFormatting);

            // Do row repetition
            DocumentPosition lastPos = newTableRange.End;

            XmlNodeList nodes = ((Xml)this.datasource).getList(jo.GetValue("loopTable").ToString(), jo.GetValue("foreignKey").ToString(), id);
            
            foreach (XmlNode node in nodes) {
                
                var bodyRange = getRowsRange(table, headerCount, rowCount);
                foreach(String field in jo.GetValue("fields")) {
                    this.replaceTextWithNewText(field, node[field].InnerText);
                }
                
                lastPos = mainWordProcessor.Document.InsertDocumentContent(lastPos, bodyRange, InsertOptions.KeepSourceFormatting).End;                
            }
                
         

            // Copy footer
            var footerRange = getRowsRange(table, headerCount + rowCount, footerCount);
            mainWordProcessor.Document.InsertDocumentContent(lastPos, footerRange, InsertOptions.KeepSourceFormatting);

            this.mainWordProcessor.Document.Delete(comment.Range);
            this.mainWordProcessor.Document.Delete(table.Range);

            
        }


        public void replaceTextWithNewTable(string text, int rows, int cols) {
            this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(text);
            this.mainWordProcessor.Document.Tables.Create(this.targetRange.Start, rows, cols);
            this.delete();
        }
        public void replaceRangeWithNewText(DocumentRange sourceRange, string targetText) {
            this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = sourceRange;
            if (this.targetRange != null)
                this.mainWordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithNewText(string sourceText, string targetText) {
            this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(sourceText);
            if (this.targetRange != null)
                this.mainWordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithImage(DocumentRange sourceRange, string targetText, string id) {
            this.mainWordProcessor.Document.BeginUpdate();
            this.mainWordProcessor.Document.Unit = DevExpress.Office.DocumentUnit.Inch;
            this.targetRange = sourceRange;
            if (this.targetRange != null) {

                byte[] bytes = Convert.FromBase64String(this.datasource.GetValue(targetText).ToString());
                bytes = ImageResizer.resize(bytes, 700, 700);
                using (MemoryStream ms = new MemoryStream(bytes)) {
                    DocumentImageSource image = DocumentImageSource.FromStream(ms);
                    this.mainWordProcessor.Document.Images.Insert(this.targetRange.Start, image);
                }
                this.delete();
            }
            //this.wordProcessor.Document.Replace(targetRange, targetText);
        }
        private void replaceTextWithTemplate(Comment comment, string file, string id, string foreightKey = "") {
            string documentTemplate = Path.Combine(ConfigurationManager.AppSettings["templates"] + file);

            //create child wordprocessor
            using (RichEditDocumentServer childWordPrecessor = new RichEditDocumentServer()) {
                //load document to child wordprocessor
                childWordPrecessor.LoadDocumentTemplate(documentTemplate);
                //swap the objects main wordprocessor in order to be accessed globaly
                this.tempWordProcessor = this.mainWordProcessor;
                this.mainWordProcessor = childWordPrecessor;

                //loop through child processor comments
                foreach (Comment c in childWordPrecessor.Document.Comments.ToList()) {
                    SubDocument doc = c.BeginUpdate();
                    String field = doc.GetText(doc.Range).Replace("”", "\"").Replace("“", "\"");
                    if (field != "") {
                        this.parseCommentTypes(field, c, id, foreightKey);
                    }
                }

                //get the main wordprocessor back
                this.mainWordProcessor = this.tempWordProcessor;
                this.mainWordProcessor.Document.InsertDocumentContent(comment.Range.End, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
            }
        }

        
        protected DocumentRange getRowsRange(Table table, int rowIndex, int rowCount) {
            DocumentPosition start = table.Rows[rowIndex].Range.Start;
            DocumentPosition end = table.Rows[rowIndex + rowCount - 1].Range.End;
            int length = end.ToInt() - start.ToInt();

            return this.mainWordProcessor.Document.CreateRange(start, length);
        }
        public Table getTable(int position) {
            return this.mainWordProcessor.Document.Tables[position];
        }
        public Paragraph getParagraph(int position) {
            return this.mainWordProcessor.Document.Paragraphs[position];
        }
        public DocumentRange getTextRange(string search) {
            try {
                Regex myRegEx = new Regex(search);
                return this.mainWordProcessor.Document.FindAll(myRegEx).First();
            } catch (Exception ex) {
                return null;
            }
        }
        public string getCommentText(Comment comment) {
            SubDocument doc = comment.BeginUpdate();
            string commentText = doc.GetText(doc.Range).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");
            comment.EndUpdate(doc);

            return commentText;
        }
    }
}
