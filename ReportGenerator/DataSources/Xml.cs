﻿

using ReportGenerator.DataSources;
using ReportGenerator.Types;
using System;
using System.Collections.Generic;
using System.Xml;
using System.Data;
using System.Linq;
using System.Configuration;

namespace ReportGenerator_v1.DataSources {
    public class Xml : IDataSource {
        private DataSet _dataSet;
        private Dictionary<String, DataColumn> binder;
        private String xmlPath = ConfigurationManager.AppSettings["xmlPath"]+ "testReport.xml";

        public Xml() {
            this.binder = getDictionary();
            
        }

        public object GetValue(string field, int index = 0) {
            DataColumn column;

            if (binder.TryGetValue(field, out column)) {
                return column.Table.Rows[index][column.Ordinal];
            }
            else {
                return field;
            }
        }

        //Function to get any field by id. You need to specify parentnode.childnode in the field parameter, and the primary key ID
        public String GetValueByID(string field, string id="") {

            string[] fields = field.Split('.');
            //var element = this.getList(fields[0] + "[ns:ID='" + id + "']");
            var element = this.getList(fields[0], "ID", id);
            return element[0][fields[1]].InnerText;
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

        public XmlNodeList getList(string node, string field, string value) {

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(this.xmlPath);
            var nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            string selector = node;
            if (value == "") {
                selector += $"[ns:{field}]";
            } else {
                selector += $"[ns:{field}='" + value + "']";
            }
            
            nsmgr.AddNamespace("ns", "http://www.civilteam.gr/dsBuildingHeatInsulation.xsd");
            XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes($"//ns:dsBuildingHeatInsulation//ns:{selector}", nsmgr);

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


        // TODO Need to do some work
        public object GetValueWithLinq() {
            DataTable PageA = _dataSet.Tables["PageADetails"];
            IEnumerable<DataRow> query =
            from pageA in PageA.AsEnumerable()
            select pageA;
            foreach (DataRow p in query) {
                Console.WriteLine(p.Field<string>("ID"));
            }

            return null;
        }

    }
}
