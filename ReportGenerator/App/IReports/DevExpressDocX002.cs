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

            foreach (Comment comment in comments.ToList()) {
                //this block of code is used to get the comments and save them in a string variable
                //all the block is needed according to devexpress
                SubDocument doc = comment.BeginUpdate();
                comment.EndUpdate(doc);
                lcomments.Add(comment);
            }
            //this.scanDocumentv2(lcomments, lastComment);
            this.scanDocument(lcomments);
            Regex r = new Regex("{PBR}");
            wordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
            
            //Console.ReadLine();       
        }


        private void scanDocumentv2(List<Comment> comments, Comment lastComment) {


            this.wordProcessor.Document.AppendText("newTable");
            var newTableRange = this.getTextRange("newTable");


            int counter = 0;
            TableCell tableCell = null;


            var test = this.wordProcessor.Document.Tables[0].Rows[1].Cells[0];


            //for (int i = 0; i < 10; i++) {
                foreach (Comment comment in comments) {
                    tableCell = this.wordProcessor.Document.Tables.GetTableCell(comment.Range.End);
                    Table table = tableCell.Table;
                    this.copyRow(table, tableCell.Row.Index, 1);
                }
            //}

            
            #region OldCode

            //DocumentRange lastCommentRange = lastComment.Range;


            //TableCell tableCell = this.wordProcessor.Document.Tables.GetTableCell(lastComment.Range.End);


            //this.SplitTable(this.wordProcessor.Document, tableCell.Table, tableCell.Row.Index);

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

            #endregion
        }

        private void scanDocumentv3(List<Comment> comments) {


            this.wordProcessor.Document.AppendText("newTable");
            var newTableRange = this.getTextRange("newTable");


            int counter = 0;
            TableCell tableCell = null;

            //var test = this.wordProcessor.Document.Tables[0].Rows[1].Cells[0];

            // Create header
            tableCell = this.wordProcessor.Document.Tables.GetTableCell(comments[0].Range.Start);
            var headerRange = this.wordProcessor.Document.CreateRange(tableCell.Table.Range.Start.ToInt(), tableCell.Row.Range.Start.ToInt());
            wordProcessor.Document.InsertDocumentContent(newTableRange.End, headerRange, InsertOptions.KeepSourceFormatting);

            // Keep last comment
            Comment lastComment = null;
            if (comments.Count > 0) {
                lastComment = comments[comments.Count - 1];
            }

            // Keep footer range
            tableCell = this.wordProcessor.Document.Tables.GetTableCell(lastComment.Range.End);
            var footerRange = this.wordProcessor.Document.CreateRange(tableCell.Row.Next.Range.Start.ToInt(), tableCell.Table.LastRow.Range.End.ToInt());

            // Do row repetition
            int count = comments.Count;
            Comment comment;
            for (int i = 0; i < 2; i++) {
                for (int indexComment = 0; indexComment < count; indexComment++) {
                    comment = comments[indexComment];
                    tableCell = this.wordProcessor.Document.Tables.GetTableCell(comment.Range.Start);
                    wordProcessor.Document.InsertDocumentContent(newTableRange.End, tableCell.Row.Range, InsertOptions.KeepSourceFormatting);
                }
            }

            // Insert footer
            wordProcessor.Document.InsertDocumentContent(newTableRange.End, footerRange, InsertOptions.KeepSourceFormatting);


            //asdsasavar hashc2 = lastComment.Range.End.GetHashCode();
            //var currentRow = tableCell.Row.Index;
            //var nextRow = tableCell.Row.Next.Index;

            //var footerRange = this.wordProcessor.Document.CreateRange(tableCell.Row.Range.End.ToInt(), tableCell.Table.Range.End.ToInt());




            //tableCell = this.wordProcessor.Document.Tables.GetTableCell(lastComment.Range.End);
            //var hashc2 = lastComment.Range.End.GetHashCode();
            //Console.WriteLine(hashc2);
            ////var currentRow = tableCell.Row.Index;
            ////var nextRow = tableCell.Row.Next.Index;

            ////var footerRange = this.wordProcessor.Document.CreateRange(tableCell.Row.Range.End.ToInt(), tableCell.Table.Range.End.ToInt());
            //var footerRange = this.wordProcessor.Document.CreateRange(tableCell.Row.Next.Range.Start.ToInt(), tableCell.Table.Range.End.ToInt());
            //wordProcessor.Document.InsertDocumentContent(newTableRange.End, footerRange, InsertOptions.KeepSourceFormatting);


            #region OldCode

            //DocumentRange lastCommentRange = lastComment.Range;


            //TableCell tableCell = this.wordProcessor.Document.Tables.GetTableCell(lastComment.Range.End);


            //this.SplitTable(this.wordProcessor.Document, tableCell.Table, tableCell.Row.Index);

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

            #endregion
        }

        private void scanDocument(List<Comment> comments) {

            this.wordProcessor.Document.AppendText("newTable");
            var newTableRange = this.getTextRange("newTable");
            

            TableCell tableCell = null;
            Table table;
            
            int headerCount = 2;
            int rowCount = 2;
            int footerCount = 2;

            // Copy header
            tableCell = this.wordProcessor.Document.Tables.GetTableCell(comments[0].Range.Start);
            table = tableCell.Table;
            var headerRange = getRowsRange(table, 0, headerCount);
            wordProcessor.Document.InsertDocumentContent(newTableRange.End, headerRange, InsertOptions.KeepSourceFormatting);

            // Do row repetition
            DocumentPosition lastPos = newTableRange.End;
            for (int i = 0; i < 4; i++) {
                var bodyRange = getRowsRange(table, headerCount, rowCount);
                lastPos = wordProcessor.Document.InsertDocumentContent(lastPos, bodyRange, InsertOptions.KeepSourceFormatting).End;
            }

            // Copy footer
            var footerRange = getRowsRange(table, headerCount + rowCount, footerCount);
            wordProcessor.Document.InsertDocumentContent(lastPos, footerRange, InsertOptions.KeepSourceFormatting);


            this.delete(this.getTextRange("newTable"));
            this.delete(this.wordProcessor.Document.Tables[0].Range);
            
        }

        protected DocumentRange getRowsRange(Table table, int rowIndex, int rowCount) {
            DocumentPosition start = table.Rows[rowIndex].Range.Start;
            DocumentPosition end = table.Rows[rowIndex + rowCount - 1].Range.End;
            int length = end.ToInt() - start.ToInt();

            return this.wordProcessor.Document.CreateRange(start, length);
        }

        public void delete(DocumentRange element) {
            this.wordProcessor.Document.Delete(element);
        }
        public DocumentRange getTextRange(string search) {
            try {
                Regex myRegEx = new Regex(search);
                return this.wordProcessor.Document.FindAll(myRegEx).First();
            } catch (Exception ex) {
                return null;
            }

        }

        public void copyRow(Table table, int rowIndex, int newRowIndex) {
            //Table table = this.wordProcessor.Document.Tables[tableIndex];
           
            table.Rows.InsertAfter(newRowIndex);
            
            //table.Rows[newRowIndex+1].Cells[0];
            //this.wordProcessor.Document.InsertSingleLineText(table[newRowIndex + 1, 0].Range.Start, "aaaa");

            //this.wordProcessor.Document.InsertSingleLineText(table[0, 0].Range.Start, "themhz");


            //for (int i = 0; i < table.Rows[rowIndex].Cells.Count; i++) {
            //    String text = this.wordProcessor.Document.GetText(table.Rows[rowIndex].Cells[i].Range);
            //    this.wordProcessor.Document.InsertText(table.Rows[newRowIndex + 1].Cells[i].Range.Start, text);
            //}

        }
        public void delete() {
            throw new NotImplementedException();
        }
    }
}
