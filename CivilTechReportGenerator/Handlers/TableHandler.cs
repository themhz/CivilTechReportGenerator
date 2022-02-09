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

namespace CivilTechReportGenerator.Handlers {
    class TableHandler: CivilDocumentX {
        
        public TableHandler(RichEditDocumentServer wordProcessor) : base() {
            base.srv = wordProcessor;
        }

        //Counts how many tables are in the document
        public int countTables() {            
            return base.document.Tables.Count;            
                
        }

        //Simply counts the rows of a table, just give the index of the table
        public void countTableRows(int index) {
            MessageBox.Show(base.document.Tables[index].Rows.Count.ToString());
        }


        //Creates a table in the document
        public override void create() {

            //Create a new table and specify its layout type
            DocumentPosition position = base.document.CreatePosition(this.pos);
            Table table = document.Tables.Create(position, x, y);
            base.saveDocument();

        }

        //Copies table in a specific position
        //as suggestion https://supportcenter.devexpress.com/ticket/details/t293243/copy-paste-paragraph-or-table
        public void copy(int tableIndex, int posTarget, String generatedfile) {

            DocumentRange myRange = document.Tables[tableIndex].Range;
            DocumentPosition dpos = document.CreatePosition(posTarget);
            document.InsertText(dpos, " ");            
            base.document.InsertDocumentContent(dpos, base.document.Tables[pos].Range);
            base.saveDocument(generatedfile);
        }

        public void copyRow(int tableIndex, int rowIndex, int newRowIndex, String generatedfile) {
            Table table = document.Tables[tableIndex];
            table.BeginUpdate();
            table.Rows.InsertBefore(newRowIndex);

            //table.Rows.InsertAfter(newRowIndex);
         

            for (int i = 0; i < table.Rows[1].Cells.Count; i++) {
                String text = document.GetText(table.Rows[rowIndex].Cells[i].Range);
                document.InsertText(table.Rows[newRowIndex + 1].Cells[i].Range.Start, text);                
            }

            //table.Rows[1].Cells.Append();
            table.EndUpdate();

            base.saveDocument(generatedfile);
        }


        // gets table items type TableData. Check folder Types for see the structurwe of TableData
        //You also need to pass the target file which will be the generatedfile. It is a string with the path of the file that will be
        //generated
        public void populateTable(List<TableData> tableItems) {            
            
            foreach (TableData td in tableItems) {
                int tableKey = int.Parse(td.TableKey);
                var table = base.document.Tables;

                table[tableKey].BeginUpdate();
                foreach (List<string> row in td.Rows) {                    
                    int rowcount = table[tableKey].Rows.Count()-1;
                    table[tableKey].Rows.InsertAfter(rowcount);
                    for (int i = 0; i < row.Count; i++) {
                        base.document.InsertSingleLineText(table[tableKey][rowcount, i].Range.Start, row[i]);
                    }

                }
                base.document.Tables[tableKey].EndUpdate();
            }
            
        }
      
    }
}
