using ReportGenerator.DataSources;
using ReportGenerator.Types;

using System.Collections.Generic;

namespace ReportGenerator_v1.System {
    class Database : IDataSource {

        List<List<string>> rows;
        List<TableData> listOfTables;
        TableData tabledata;
        public Database(List<List<string>> _rows, List<TableData> _listOfTables, TableData _tabledata) {
            this.rows = _rows;
            this.listOfTables = _listOfTables;
            this.tabledata = _tabledata;
        }
        public List<List<string>> getData() {

            return null;
        }
        public List<TableData> getTableData() {
                                    
            this.tabledata.Rows.Add(new List<string> { "col1", "col2", "col3", "col4" });
            this.tabledata.Rows.Add(new List<string> { "col1", "col2", "col3", "col4" });
            this.tabledata.Rows.Add(new List<string> { "col1", "col2", "col3", "col4", "col5" });
            this.tabledata.Rows.Add(new List<string> { "col1", "col2", "col3", "col4" });

            this.listOfTables.Add(tabledata);
            return this.listOfTables;
        }        

    }
}
