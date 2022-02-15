using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CivilTechReportGenerator.Types;
using System.Text.RegularExpressions;
using CivilTechReportGenerator.Interfaces;

namespace CivilTechReportGenerator.Handlers {
    
    public class TableItem : DocumentXItem, ITableItem {        
        public Regex myRegEx { set; get; }
        public DocumentRange dr { set; get; }
        public DocumentPosition dpos { set; get; }
        public int rows { set; get; }
        public int cols { set; get; }

        public TableItem setRows(int val) {
            rows = val;
            return this;
        }

        public TableItem setCols(int val) {
            cols = val;
            return this;
        }

        public TableItem(RichEditDocumentServer _wordProcessor){
            this.wordProcessor = _wordProcessor;            
        }

        //Counts how many tables are in the document
        public override int count() {
            return this.wordProcessor.Document.Tables.Count;
        }

        //Simply counts the rows of a table, just give the index of the table
        public void countTableRows(int index) {
            MessageBox.Show(this.wordProcessor.Document.Tables[index].Rows.Count.ToString());
        }

        //Creates a table in the document
        public override void create() {

            //Create a new table and specify its layout type
            DocumentPosition position = this.wordProcessor.Document.CreatePosition(documentPosition);
            base.createSpace(documentPosition);
            Table table = wordProcessor.Document.Tables.Create(position, rows, cols);
        }

        public Table findTable(int pos) {
            return this.wordProcessor.Document.Tables[pos];
        }

        //Copies table in a specific position
        //as suggested https://supportcenter.devexpress.com/ticket/details/t293243/copy-paste-paragraph-or-table
        public void copy(int tableIndex, int posTarget) {

            base.createSpace(posTarget);            
            this.wordProcessor.Document.InsertDocumentContent(this.dpos, this.wordProcessor.Document.Tables[tableIndex].Range);
        }

        public override void delete(int index) {
            this.wordProcessor.Document.Delete(this.wordProcessor.Document.Tables[index].Range);
        }

        public void copyRow(int tableIndex, int rowIndex, int newRowIndex) {
            Table table = this.wordProcessor.Document.Tables[tableIndex];
            table.BeginUpdate();
            table.Rows.InsertAfter(newRowIndex);

            for (int i = 0; i < table.Rows[rowIndex].Cells.Count; i++) {
                String text = this.wordProcessor.Document.GetText(table.Rows[rowIndex].Cells[i].Range);
                this.wordProcessor.Document.InsertText(table.Rows[newRowIndex + 1].Cells[i].Range.Start, text);
            }
            table.EndUpdate();
        }


        // gets table items type TableData. Check folder Types for see the structurwe of TableData
        //You also need to pass the target file which will be the generatedfile. It is a string with the path of the file that will be
        //generated
        public void populateTable(List<TableData> tableItems) {

            foreach (TableData td in tableItems) {
                int tableKey = int.Parse(td.TableKey);
                var table = this.wordProcessor.Document.Tables;

                table[tableKey].BeginUpdate();
                foreach (List<string> row in td.Rows) {
                    int rowcount = table[tableKey].Rows.Count() - 1;
                    table[tableKey].Rows.InsertAfter(rowcount);
                    for (int i = 0; i < row.Count; i++) {
                        this.wordProcessor.Document.InsertSingleLineText(table[tableKey][rowcount, i].Range.Start, row[i]);
                    }

                }
                this.wordProcessor.Document.Tables[tableKey].EndUpdate();
            }

        }

        public void replace(String generatedfile, int pos, Regex _myRegEx) {

            this.myRegEx = _myRegEx;
            this.dr = this.wordProcessor.Document.FindAll(myRegEx).First();
            this.dpos = this.wordProcessor.Document.CreatePosition(dr.Start.ToInt());
            this.wordProcessor.Document.InsertText(dpos, " ");
            this.wordProcessor.Document.InsertDocumentContent(dpos, this.wordProcessor.Document.Tables[pos].Range);
            this.wordProcessor.Document.Delete(dr);
        }

    }
}
