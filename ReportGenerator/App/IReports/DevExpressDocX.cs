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
        /// <summary>
        /// Creats the document, it loads the master template, parses, saves and finaly opens it
        /// </summary>
        /// <returns>
        /// and returns the document itself
        /// </returns>
        public IReport create() {
            Console.WriteLine("creating document "+ this.generatedfile.ToString());
            using (this.mainWordProcessor) {
                this.load();
                this.parse();
                this.save();
                this.openfile();
            }
            Console.WriteLine("file report created");
            return this;
        }
        /// <summary>
        /// Loads the document and begins the update
        /// </summary>
        public void load() {
            this.mainWordProcessor.Document.BeginUpdate();
            this.mainWordProcessor.LoadDocument(this.template);
        }
        /// <summary>
        /// Saves the document to the folder specified in the App.config, and the path is in the this.generatedfile
        /// </summary>        
        public void save() {
            this.mainWordProcessor.Document.EndUpdate();
            Console.WriteLine("Saving file report");
            this.mainWordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);
            Console.WriteLine("Report save in :" + this.template);
        }
        /// <summary>
        /// Opens the file with the word application 
        /// </summary>
        public void openfile() {
            Process.Start(new ProcessStartInfo(this.generatedfile) { UseShellExecute = true });
        }                                    
        /// <summary>
        /// Adds a row to the target table that is specified
        /// </summary>
        /// <param name="targetTable">The target table that the row will be added</param>
        /// <param name="cols">You need to specify a dictionary with the column name and values</param>
        /// <param name="jo">the json object with the actually values of the row</param>
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
        /// <summary>
        /// populates the totals of the table. It is the final row of a normal table
        /// </summary>
        /// <param name="DetailList">the xml node list object that is created by testReport.xml</param>
        /// <param name="table">The table object</param>
        /// <param name="jo">the json object</param>
        /// <param name="id">and the primary key of the table</param>
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
        /// <summary>
        /// Loops through the xml datasource on a specific table or node and parses it by executing the replaceTextWithTemplate function
        /// </summary>
        /// <param name="jo">the json object</param>
        /// <param name="comment">the comment that was parsed</param>
        /// <param name="id">the id of the table to be looped</param>
        /// <param name="foreignKey">the foreign key name</param>
        public void loopTable(JObject jo, Comment comment, string id = "", string foreignKey = "") {

            XmlNodeList tables = ((Xml)datasource).getList(jo.GetValue("loopTable").ToString(), jo.GetValue("foreignKey").ToString(), id);
            foreach (XmlNode table in tables) {
                this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), table[jo.GetValue("id").ToString()].InnerText, table[jo.GetValue("foreignKey").ToString()].InnerText);
            }
            this.mainWordProcessor.Document.Delete(comment.Range);
        }
        /// <summary>
        /// populates the simple table by adding rows to it. It reads the xmlNode
        /// </summary>
        /// <param name="DetailList">the xml node list object that is created by testReport.xml</param>
        /// <param name="table">The table object</param>
        /// <param name="jo">the json object</param>
        /// <param name="id">and the primary key of the table</param>
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
        /// <summary>
        /// Deletes the object specified within the current object range
        /// </summary>
        public void delete() {
            this.mainWordProcessor.Document.Delete(this.targetRange);
        }
        /// <summary>
        /// Parses all the comments of a document. Basically everything starts here
        /// </summary>
        public void parse() {
            //Collect all comments and loop through them in order to parse each one of them individually and create the document elemtnts
            CommentCollection comments = this.mainWordProcessor.Document.Comments;
            foreach (Comment comment in comments.ToList()) {
                                
                //After collecting the individual comment witch is in json format, it will be passed to the internal function parseFieldTypes in order to 
                //parse and create the fields
                this.parseCommentTypes(comment);
                //then move to the next comment..
            }
            Regex r = new Regex("{PBR}");
            mainWordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
        }
        ///<summary>
        ///This function is used to check the type of parsing that will be used        
        ///</summary>
        ///<param name="json">the json string to be parsed as jobject and get the parse type</param>
        ///<param name="comment">the comment object that is passsed</param>
        private void parseCommentTypes(Comment comment, string id = "", string foreignKey = "") {
            string json = this.getCommentText(comment);
            try {
                JObject jsonObject = JObject.Parse(json);
                switch (jsonObject.GetValue("type").ToString()) {
                    case "field":
                        this.parseField(jsonObject, comment, id);
                        break;
                    case "image":
                        this.parseImage(jsonObject, comment, id);
                        break;                   
                    case "template":
                        this.parseTemplate(jsonObject, comment, id);
                        break;
                    case "table":
                        this.parseTable(jsonObject, comment, id);
                        break;
                    case "complexTable":
                        this.parseComplexTable(jsonObject, comment, id);
                        break;
                }
            } catch (Exception ex) {

            }


        }
        /// <summary>
        /// replaces the targeted range that the comment points to with the text that relates to the name
        /// </summary>
        ///<param name="json">the json string to be parsed as jobject and get the parse type</param>
        ///<param name="comment">the comment object that is passsed</param>
        /// <param name="id"></param>
        public void parseField(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is field");
            this.replaceRangeWithNewText(comment.Range, datasource.GetValue(jo.GetValue("name").ToString()).ToString());
        }      
        /// <summary>
        /// Parses the template in the comment
        /// Finds the looptable attribute in the json comment and loops so it prints the table as many times as it is found in the xml
        /// else it replaces with a template
        /// </summary>
        /// <param name="jo">the json object</param>
        /// <param name="comment">the comment object</param>
        /// <param name="id">the id as primary key</param>
        public void parseTemplate(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is template");
            if (jo.ContainsKey("loopTable")) {
                this.loopTable(jo, comment, id);
            } else {                
                this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id);
            }            
        }        
        /// <summary>
        /// Gets the image and prints in the document. Notice that the image is in binary format not an actual jpg and its embedid in the xml
        /// </summary>
        /// <param name="jo">the json object</param>
        /// <param name="comment">the comment object</param>
        /// <param name="id">the id as primary key</param>
        public void parseImage(JObject jo, Comment comment, string id = "") {
            Console.WriteLine(jo + " is image");            
            XmlNodeList DetailList = ((Xml)datasource).getList("PageA","ID", id);            
            this.replaceTextWithImage(comment.Range, DetailList[0]["Image"].InnerText, id);            
        }      
        /// <summary>
        /// Parses the table comment and populates the table in the document
        /// </summary>
        /// <param name="jo">the json object</param>
        /// <param name="comment">the comment object</param>
        /// <param name="id">the id</param>
        public void parseTable(JObject jo, Comment comment, string id = "") {
            //string data = jo.GetValue("cols").ToString().Replace("[", " ").Replace("]", " ").Replace(Environment.NewLine, "");
            string loopTable = jo.GetValue("loopTable").ToString();
            string foreignKey = jo.GetValue("foreignKey").ToString();
            TableCell tableCell = this.mainWordProcessor.Document.Tables.GetTableCell(comment.Range.Start);

            if(tableCell != null) {                
                XmlNodeList DetailList = ((Xml)datasource).getList(loopTable, foreignKey, id);
                this.populateTable(DetailList, tableCell.Table, jo, id);
            }
        }
        /// <summary>
        /// parses a compex table where you can expand n rows a time, keeping header and footer in place
        /// </summary>
        /// <param name="jo">json object</param>
        /// <param name="comment">comment object</param>
        /// <param name="id">the primary key of the table</param>        
        public void parseComplexTable(JObject jo, Comment comment, string id = "") {
            
            this.mainWordProcessor.Document.AppendText("newTable");
            var newTableRange = this.getTextRange("newTable");
            TableCell tableCell = null;
            Table table;

            int headerCount = Int32.Parse(jo.GetValue("headerCount").ToString());
            int rowCount = Int32.Parse(jo.GetValue("rowCount").ToString());
            int footerCount = Int32.Parse(jo.GetValue("footerCount").ToString());

            //Get table 
            tableCell = this.mainWordProcessor.Document.Tables.GetTableCell(comment.Range.Start);
            table = tableCell.Table;

            // Copy header
            var headerRange = getRowsRange(table, 0, headerCount);
            mainWordProcessor.Document.InsertDocumentContent(newTableRange.End, headerRange, InsertOptions.KeepSourceFormatting);

            // Do row repetition
            DocumentPosition lastPos = newTableRange.End;
            XmlNodeList nodes = ((Xml)this.datasource).getList(jo.GetValue("loopTable").ToString(), jo.GetValue("foreignKey").ToString(), id);
            
            foreach (XmlNode node in nodes) {
               
                var bodyRange = getRowsRange(table, headerCount, rowCount);               
                lastPos = mainWordProcessor.Document.InsertDocumentContent(lastPos, bodyRange, InsertOptions.KeepSourceFormatting).End;

                foreach (String field in jo.GetValue("fields")) {
                    this.replaceTextWithNewTextLast(field, node[field].InnerText);
                }
            }


            // Copy footer
            DocumentRange footerRange = getRowsRange(table, headerCount + rowCount, footerCount);
           
            mainWordProcessor.Document.InsertDocumentContent(lastPos, footerRange, InsertOptions.KeepSourceFormatting);

            this.mainWordProcessor.Document.Delete(comment.Range);
            this.mainWordProcessor.Document.Delete(table.Range);

            
        }


        public void replaceTextWithNewTable(string text, int rows, int cols) {
            //this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(text);
            this.mainWordProcessor.Document.Tables.Create(this.targetRange.Start, rows, cols);
            this.delete();
        }
        public void replaceRangeWithNewText(DocumentRange sourceRange, string targetText) {
            //this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = sourceRange;
            if (this.targetRange != null)
                this.mainWordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithNewText(string sourceText, string targetText) {
            //this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(sourceText);
            if (this.targetRange != null)
                this.mainWordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithNewTextLast(string sourceText, string targetText) {
            //this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRangeLast(sourceText);
            if (this.targetRange != null)
                this.mainWordProcessor.Document.Replace(targetRange, targetText);
        }
        public void replaceTextWithImage(DocumentRange sourceRange, string targetText, string id) {
            //this.mainWordProcessor.Document.BeginUpdate();
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
                        this.parseCommentTypes(c, id, foreightKey);
                    }
                    doc.EndUpdate();
                }

                //get the main wordprocessor back
                this.mainWordProcessor = this.tempWordProcessor;
                this.mainWordProcessor.Document.InsertDocumentContent(comment.Range.End, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
            }
        }


        /// <summary>
        /// Gets the row range
        /// </summary>
        /// <param name="table">the table object</param>
        /// <param name="rowIndex">the rowIndex that we will get the range</param>
        /// <param name="rowCount">how many rows</param>
        /// <returns></returns>
        protected DocumentRange getRowsRange(Table table, int rowIndex, int rowCount) {
            DocumentPosition start = table.Rows[rowIndex].Range.Start;
            DocumentPosition end = table.Rows[rowIndex + rowCount - 1].Range.End;
            int length = end.ToInt() - start.ToInt();

            return this.mainWordProcessor.Document.CreateRange(start, length);
        }
        /// <summary>
        /// get the table of the document by the index
        /// </summary>
        /// <param name="position">the index position</param>
        /// <returns></returns>
        public Table getTable(int position) {
            return this.mainWordProcessor.Document.Tables[position];
        }
        /// <summary>
        /// gets the paragraph by index
        /// </summary>
        /// <param name="position">the index the document</param>
        /// <returns></returns>
        public Paragraph getParagraph(int position) {
            return this.mainWordProcessor.Document.Paragraphs[position];
        }
        /// <summary>
        /// Searches for a text in a document and gets its range. The search happens from top to bottom
        /// </summary>
        /// <param name="search">the text to search</param>
        /// <returns></returns>
        public DocumentRange getTextRange(string search) {
            try {
                Regex myRegEx = new Regex(search);
                return this.mainWordProcessor.Document.FindAll(myRegEx).First();
            } catch (Exception ex) {
                return null;
            }
        }
        /// <summary>
        /// Searches for a text in a document and gets its range. The search happens from bottom to top
        /// </summary>
        /// <param name="search">the text to search</param>
        /// <returns></returns>
        public DocumentRange getTextRangeLast(string search) {
            try {
                Regex myRegEx = new Regex(search);
                return this.mainWordProcessor.Document.FindAll(myRegEx).Last();
            } catch (Exception ex) {
                return null;
            }
        }
        /// <summary>
        /// gets comment text
        /// </summary>
        /// <param name="comment">the comment object</param>
        /// <returns></returns>
        public string getCommentText(Comment comment) {
            SubDocument doc = comment.BeginUpdate();
            string commentText = doc.GetText(doc.Range).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");            
            comment.EndUpdate(doc);
            
            return commentText;
        }
    }
}
