

using ReportGenerator.DataSources;
using ReportGenerator.Types;
using System;
using System.Collections.Generic;
using System.Xml;
using System.Data;

namespace ReportGenerator_v1.DataSources {
    class Xml : IDataSource {
        private DataSet _dataSet;
        private Dictionary<String, DataColumn> binder;

        public Xml() {
            this.binder = getDictionary();
        }

        public object GetValue(string field, int index = 0) {
            DataColumn column;

            if (binder.TryGetValue(field, out column)) {
                return column.Table.Rows[index][column.Ordinal];
            }
            else {
                return null;
            }
        }

        private Dictionary<String, DataColumn> getDictionary() {
            var dictionary = new Dictionary<String, DataColumn>();

            _dataSet = new DataSet();
            _dataSet.ReadXml("d:\\Trans\\databases\\testReport.xml");

            foreach (DataTable table in _dataSet.Tables) {
                foreach (DataColumn column in table.Columns) {
                    dictionary.Add(table.TableName + "." + column.ColumnName, column);
                }
            }

            return dictionary;
        }
    }
}
