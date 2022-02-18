using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.API.Native.Implementation;
using ReportGenerator;
using ReportGenerator.Interfaces.Elements;
using ReportGenerator.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ReportGenerator_v1.System {

    class DevExpressDocX : IReport {

        public RichEditDocumentServer wordProcessor;
        public String template { set; get; }
        public String generatedfile { set; get; }
        private DocumentRange sourceRange { set; get; }
        private DocumentRange targetRange { set; get; }
        

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
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.LoadDocument(this.template);
        }

        public void save() {
            this.wordProcessor.Document.EndUpdate();
            Console.WriteLine("Saving file report");            
            this.wordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);            
            Console.WriteLine("Report save in :"+ this.template);            
        }

        //This function is under construction. It will be used to parse the word template document and conscrtruct the report
        //However some commans are implemented.. more to come..
        public void parse() {

            //#Replace text with new table
            //this.replaceTextWithNewTable("{table1}", 2,3);

            //#Copy any element, 
            //this.sourceRange = this.getTable(0).Range;
            //this.targetRange = this.getTable(1).Range;
            //this.copy();

            //#Move any element, 
            //this.sourceRange = this.getTextRange("{table2}"); 
            //this.targetRange = this.getTable(0).Range;
            //this.move();

            //#Delete any element, 
            //this.targetRange = this.getTextRange("{table2}");
            //this.delete();


            //#Copy any element tha corresponds to comment, 
            //this.sourceRange = this.wordProcessor.Document.Comments[0].Range;            
            ////String test = ((NativeComment)this.wordProcessor.Document.Comments[0]).Comment.Content.PieceTable.TextBuffer.ToString();
            //this.targetRange = this.getTable(2).Range;
            //this.copy();

            //#Delete any element related to a specific comment. 
            //this.targetRange = this.wordProcessor.Document.Comments[0].Range;
            //delete();

            //#populate Table, this uses a dummy datasource at the moment
            //this.populateTable(this.getTable(1));

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
        protected DocumentRange getElementRange() {
            return null;
        }
        public Table getTable(int position) {
           return this.wordProcessor.Document.Tables[position];            
        }
        public Paragraph getParagraph(int position) {
            return this.wordProcessor.Document.Paragraphs[position];
        }
        public DocumentRange getTextRange(String search) {
            Regex myRegEx = new Regex(search);
            return this.wordProcessor.Document.FindAll(myRegEx).First();
        }
        public void delete() {
            this.wordProcessor.Document.Delete(this.targetRange);            
        }

        private void populateTable(Table table) {
            TableData tabledata = new TableData();
            List<List<string>> rows = new List<List<string>>();

            List<string> row1 = new List<string> { "col1", "col2", "col3", "col4" };
            List<string> row2 = new List<string> { "col1", "col2", "col3", "col4" };
            List<string> row3 = new List<string> { "col1", "col2", "col3", "col4" };
            List<string> row4 = new List<string> { "col1", "col2", "col3", "col4" };

            //tabledata.TableKey = "1";
            tabledata.Rows.Add(row1);
            tabledata.Rows.Add(row2);
            tabledata.Rows.Add(row3);
            tabledata.Rows.Add(row4);            
                        
            List<TableData> tableDatas = new List<TableData>();
            tableDatas.Add(tabledata);

            foreach (TableData td in tableDatas) {
                //int tableKey = int.Parse(td.TableKey);
                var tbl = this.wordProcessor.Document.Tables;

                table.BeginUpdate();
                foreach (List<string> row in td.Rows) {
                    int rowcount = table.Rows.Count() - 1;
                    table.Rows.InsertAfter(rowcount);
                    for (int i = 0; i < row.Count; i++) {
                        this.wordProcessor.Document.InsertSingleLineText(table[rowcount, i].Range.Start, row[i]);
                    }

                }
                table.EndUpdate();
            }

        }
    }
}
