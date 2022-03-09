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
            CommentCollection comments = this.wordProcessor.Document.Comments;

            foreach (Comment comment in comments.ToList()) {

                SubDocument doc = comment.BeginUpdate();
                String field = doc.GetText(doc.Range).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");
                comment.EndUpdate(doc);
                this.checkField(field, comment);
            }
            //Console.ReadLine();       
        }
        public void populatePageADetails(XmlNodeList DetailList, Table table) {            
            foreach (XmlNode node in DetailList) {
                //List<string> rows = new List<string>();
                Dictionary<String, String> cols = new Dictionary<string, string>();
                foreach (XmlNode row in node) {
                    cols.Add(row.Name, row.InnerText);
                }
                this.addTableRow(table, cols);                
            }


        }
        //Τhe only way to copy and paste something is via InsertDocumentContent method
        //https://supportcenter.devexpress.com/ticket/details/t725837/richeditdocumentserver-copy-paste-problem
        public void copy() {
            wordProcessor.Document.InsertDocumentContent(this.targetRange.End, this.sourceRange);
        }
        //Moves any table 
        public void move() {
            wordProcessor.Document.InsertDocumentContent(this.targetRange.End, this.sourceRange);
            this.targetRange = this.sourceRange;
            this.delete();
        }
        public void replaceTextWithNewTable(String text, int rows, int cols) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(text);
            this.wordProcessor.Document.Tables.Create(this.targetRange.Start, rows, cols);
            this.delete();
        }
        public void replaceRangeWithNewText(DocumentRange sourceRange, String targetText) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = sourceRange;
            if (this.targetRange != null)
                this.wordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithNewText(String sourceText, String targetText) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(sourceText);
            if (this.targetRange != null)
                this.wordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithImage(DocumentRange sourceRange, String targetText) {
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.Document.Unit = DevExpress.Office.DocumentUnit.Inch;
            this.targetRange = sourceRange;  //this.getTextRange(sourceText);
            if (this.targetRange != null) {
                byte[] bytes = Convert.FromBase64String(targetText);
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
        public DocumentRange getTextRange(String search) {
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
        private void addTableRow(Table targetTable, Dictionary<String, String> cols) {
            int rowcount = targetTable.Rows.Count() - 1;
            targetTable.Rows.InsertAfter(rowcount);
            try {
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 0].Range.Start, cols["Index"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 1].Range.Start, cols["Name"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 2].Range.Start, cols["Density"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 3].Range.Start, cols["d"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 4].Range.Start, cols["λ"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 5].Range.Start, MathOperations.formatTwoDecimalWithoutRound(cols["dλ"].ToString(), 4));
            } catch (Exception ex) {

            }

        }        

        private void checkField(String field, Comment comment) {
            JObject jo = JObject.Parse(field);
            //Console.WriteLine(jo);
            switch (jo.GetValue("type").ToString()) {
                case "field":
                    this.parseField(jo, comment);
                    break;
                case "image":
                    this.parseImage(jo, comment);
                    break;
                case "list":
                    this.parseList(jo, comment);
                    break;
                case "template":
                    this.parseTemplate(jo, comment);
                    break;
                case "table":
                    this.parseTable(jo, comment);
                    break;
            }

        }

        private void replaceTextWithTemplate(Comment comment, String file) {
            string documentTemplate = Path.Combine(ConfigurationManager.AppSettings["templates"] + file + ".docx");

            using (RichEditDocumentServer subWordPrecessor = new RichEditDocumentServer()) {
                subWordPrecessor.LoadDocumentTemplate(documentTemplate);

                this.tempWordProcessor = this.wordProcessor;
                this.wordProcessor = subWordPrecessor;
                foreach (Comment c in subWordPrecessor.Document.Comments.ToList()) {
                    SubDocument doc = c.BeginUpdate();
                    String field = doc.GetText(doc.Range).Replace("”", "\"").Replace("“", "\"");
                    this.checkField(field, c);
                }
                this.wordProcessor = this.tempWordProcessor;
                this.wordProcessor.Document.InsertDocumentContent(comment.Range.End, subWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
            }
            this.wordProcessor.Document.Delete(comment.Range);
        }


        //To be continued
        public void parseField(JObject jo, Comment comment) {
            Console.WriteLine(jo + " is field");
            this.replaceRangeWithNewText(comment.Range, datasource.GetValue(jo.GetValue("name").ToString()).ToString());
        }

        public void parseList(JObject jo, Comment comment) {
            Console.WriteLine(jo + " is List");
        }

        public void parseTemplate(JObject jo, Comment comment) {
            Console.WriteLine(jo + " is template");
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString());
        }

        public void parseImage(JObject jo, Comment comment) {
            Console.WriteLine(jo + " is image");
            this.replaceTextWithImage(comment.Range, jo.GetValue("name").ToString());
        }
        public void parseTable(JObject jo, Comment comment) {            
            string data = jo.GetValue("fields").ToString().Replace("[", " ").Replace("]", " ").Replace(Environment.NewLine, "");
            string[] fields = data.Split(new char[]{ ',' }, StringSplitOptions.RemoveEmptyEntries);

            TableCell tableCell = this.wordProcessor.Document.Tables.GetTableCell(comment.Range.Start);

            if(tableCell != null) {
                int rowcount = tableCell.Table.Rows.Count() - 1;
                tableCell.Table.Rows.InsertAfter(rowcount);

                //XmlNodeList DetailList = ((Xml)datasource).getList("PageADetails[ns:PageADetailID='" + page["ID"].InnerText + "']");
                //this.populatePageADetails(DetailList, this.wordProcessor.Document.Tables[counter + 1]);

                int counter = 0;
                foreach (String value in fields) {
                    Console.WriteLine(value.Replace("\"", ""));
                    this.wordProcessor.Document.InsertSingleLineText(tableCell.Table[rowcount, counter].Range.Start, value.Replace("\"", ""));
                    counter++;
                }
            }
        }
    }
}
