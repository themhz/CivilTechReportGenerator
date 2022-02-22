

using ReportGenerator.DataSources;
using ReportGenerator.Types;
using System;
using System.Collections.Generic;
using System.Xml;

namespace ReportGenerator_v1.DataSources {
    class Xml : IDataSource {
        List<List<string>> rows;
        List<TableData> listOfTables;
        TableData tabledata;
        public List<List<string>> getData() {

            //XmlDocument doc = new XmlDocument();
            ////d:\Trans\databases\testReport.xml
            //doc.Load("d:\\Trans\\databases\\testReport.xml");

            //XmlNode node = doc.DocumentElement.SelectSingleNode("/PageA");
            XmlTextReader reader = new XmlTextReader("d:\\Trans\\databases\\testReport.xml");

            int counter = 0;
            while (reader.Read()) {
                // Do some work here on the data.
                switch (reader.NodeType) {
                    case XmlNodeType.Element: // The node is an element.
                        Console.Write("<" + reader.Name);
                        while (reader.MoveToNextAttribute()) // Read the attributes.
                            Console.Write(" " + reader.Name + "='" + reader.Value + "'");                        
                            Console.WriteLine(">");
                        break;
                    case XmlNodeType.Text: //Display the text in each element.
                        Console.WriteLine("text is :"+ reader.Value);
                        break;
                    case XmlNodeType.EndElement: //Display the end of the element.
                        Console.Write("</" + reader.Name);
                        Console.WriteLine(">");
                        break;                        
                }


                if (counter > 40) {
                    break;
                }
                counter++;
            }
            Console.ReadLine();
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
