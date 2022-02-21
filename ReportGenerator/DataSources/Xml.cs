

using ReportGenerator.DataSources;
using ReportGenerator.Types;
using System.Collections.Generic;

namespace ReportGenerator_v1.DataSources {
    class Xml : IDataSource {
        List<List<string>> rows;
        List<TableData> listOfTables;
        TableData tabledata;
        public List<List<string>> getData() {
            throw new global::System.NotImplementedException();
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
