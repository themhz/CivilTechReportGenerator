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
            
            List<Comment> lcomments = new List<Comment>();

            Comment lastComment = null;
            foreach (Comment comment in comments.ToList()) {
                //this block of code is used to get the comments and save them in a string variable
                //all the block is needed according to devexpress
                SubDocument doc = comment.BeginUpdate();
                String field = doc.GetText(doc.Range).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");
                comment.EndUpdate(doc);
                lcomments.Add(comment);
                lastComment = comment;
            }           
            this.scanDocument(lcomments, lastComment);
            Regex r = new Regex("{PBR}");
            wordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
            
            //Console.ReadLine();       
        }

        private void scanDocument(List<Comment> comments, Comment lastComment) {

            //DocumentRange lastCommentRange = lastComment.Range;

            //foreach(Comment comment in comments) {
            //    //TableCell tableCell = this.wordProcessor.Document.Tables.GetTableCell(comment.Range.Start);
            //    wordProcessor.Document.InsertDocumentContent(lastComment.Range.End, comment.Range, InsertOptions.KeepSourceFormatting);
            //}
            TableCell tableCell = this.wordProcessor.Document.Tables.GetTableCell(lastComment.Range.End);
            
            
            this.SplitTable(this.wordProcessor.Document, tableCell.Table, tableCell.Row.Index);

            //wordProcessor.Document.InsertDocumentContent(tableCell.Table.Range.End, tableCell.Row.Range, InsertOptions.KeepSourceFormatting);            

            //var test = tableCell.Row.Index;
            ////tableCell.Table


            //Regex r = new Regex("{{.*?}}");
            //var result = this.wordProcessor.Document.FindAll(r).GetAsFrozen() as DocumentRange[];
            //Dictionary<String, JObject> fields = new Dictionary<String, JObject>();
            //for (int i = 0; i < result.Length; i++) {
            //    //var data = this.wordProcessor.Document.GetText(result[i]);
            //    String field = this.wordProcessor.Document.GetText(result[i]).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");

            //    JObject o1 = JObject.Parse(field);
            //    fields.Add("{{" + i.ToString() + "}}", o1);
            //}

            //return fields;
        }

        public void SplitTable(Document document, Table table, int rowIndex) {
            Paragraph newParagraph = document.Paragraphs.Insert(table.Range.End);
            // copy original table  
            string tableContent = document.GetRtfText(table.Range);
            DocumentRange newTableRange = document.InsertRtfText(newParagraph.Range.End, tableContent);
            try {
                Table newTable = document.Tables.Get(newTableRange)[0];
                // remove rows in original table  
                int rowsCount = table.Rows.Count;
                for (int i = rowIndex; i < rowsCount; i++) {
                    table.Rows.RemoveAt(table.Rows.Count - 1);
                }
                // remove rows in new table  
                for (int i = 0; i < rowIndex; i++) {
                    newTable.Rows.RemoveAt(0);
                }
            } catch(Exception ex) {

            }
            
        }
        public void delete() {
            throw new NotImplementedException();
        }
    }
}
