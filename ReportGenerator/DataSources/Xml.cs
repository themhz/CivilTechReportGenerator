

using ReportGenerator.DataSources;
using ReportGenerator.Types;
using System;
using System.Collections.Generic;
using System.Xml;
using System.Data;

namespace ReportGenerator_v1.DataSources {
    public class Xml : IDataSource {
        private DataSet _dataSet;
        private Dictionary<String, DataColumn> binder;
        private String xmlPath = "C:\\Users\\themis\\source\\repos\\CivilTechReportGenerator\\ReportGenerator\\DataSources\\files\\testReport.xml";

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
            _dataSet.ReadXmlSchema(this.xmlPath);
            _dataSet.ReadXml(this.xmlPath, XmlReadMode.ReadSchema);

            foreach (DataTable table in _dataSet.Tables) {
                foreach (DataColumn column in table.Columns) {
                    dictionary.Add(table.TableName + "." + column.ColumnName, column);
                }
            }

            return dictionary;
        }

        public XmlNodeList getList(String field) {

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(this.xmlPath);
            var nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("ns", "http://www.civilteam.gr/dsBuildingHeatInsulation.xsd");
            XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes($"//ns:dsBuildingHeatInsulation//ns:{field}", nsmgr);

            return nodeList;
        }

        public Dictionary<String, DataColumn> getXmlDictionary() {
            XmlTextReader reader = new XmlTextReader(this.xmlPath);
            while (reader.Read()) {
                // Do some work here on the data.
                Console.WriteLine(reader.Name);
            }
            Console.ReadLine();
            return null;
        }


    }
}
