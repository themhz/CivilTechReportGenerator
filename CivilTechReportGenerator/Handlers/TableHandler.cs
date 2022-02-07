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
//using CivilTechReportGenerator.Types;


namespace CivilTechReportGenerator.Handlers {
    class TableHandler:CivilDocumentX {                

        public TableHandler(): base() {            
        }


        public int countTables() {            
            return base.document.Tables.Count;            
                
        }


        public void createTable(int x, int y) {
            
            using (base.srv) {
                Document document = base.srv.Document;
                //Create a new table and specify its layout type


                
                //Table table = document.Tables.Create(document.Range.End, x, y);

                document.InsertSection(document.Tables.Create(document.Range.End, x, y).Range.End);

       

                //Add new rows to the table
                //TableRow newRowBefore = table.Rows.InsertBefore(0);
                //TableRow newRowAfter = table.Rows.InsertAfter(0);

                ////Add a new column to the table
                //TableCell newLastColumn = table.Rows[0].Cells.Append();

                base.saveDocument();
            }

        }
        


    }
}
