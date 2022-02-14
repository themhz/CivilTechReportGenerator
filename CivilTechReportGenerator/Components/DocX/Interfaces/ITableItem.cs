using CivilTechReportGenerator.Types;
using DevExpress.XtraRichEdit.API.Native;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CivilTechReportGenerator.Handlers {
    interface ITableItem {
        void copy(int tableIndex, int posTarget, string generatedfile);
        void copyRow(int tableIndex, int rowIndex, int newRowIndex);
        int count();
        void countTableRows(int index);
        void create();
        void delete(int index);
        Table findTable(int pos);
        void populateTable(List<TableData> tableItems);
        void replace(string generatedfile, int pos, Regex _myRegEx);
        TableItem setCols(int val);
        TableItem setRows(int val);
    }
}