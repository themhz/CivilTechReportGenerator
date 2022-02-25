﻿using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.API.Native.Implementation;
using ReportGenerator;
using ReportGenerator.DataSources;
using ReportGenerator.Helpers;
using ReportGenerator.Types;
using ReportGenerator_v1.DataSources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace ReportGenerator_v1.System {

    class DevExpressDocX : IReport {

        public RichEditDocumentServer wordProcessor { set; get; }
        public IDataSource datasource { set; get; }
        public String template { set; get; }
        public String generatedfile { set; get; }
        private DocumentRange sourceRange { set; get; }
        private DocumentRange targetRange { set; get; }


        public DevExpressDocX(RichEditDocumentServer _wordProcessor, IDataSource _datasource) {
            this.wordProcessor = _wordProcessor;
            this.datasource = _datasource;
        }
        public IReport create() {
            Console.WriteLine("creating file report");
            using (this.wordProcessor) {
                this.load();
                this.parse();
                this.save();
            }
            Console.WriteLine("file report created");
            return this;
        }
        public void load() {
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.LoadDocument(this.template);
        }
        public void save() {
            this.wordProcessor.Document.EndUpdate();
            Console.WriteLine("Saving file report");
            this.wordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);
            Console.WriteLine("Report save in :" + this.template);
        }
                //This function is under construction. It will be used to parse the word template document and conscrtruct the report
        //However some commans are implemented.. more to come..
        public void parse() {

            
            //#Replace text with new table
            //this.replaceTextWithNewTable("{table1}", 2,3);

            //#Copy any element, 
            //this.sourceRange = this.getTable(0).Range;
            //this.targetRange = this.getTable(1).Range;
            //this.copy();

            //#Move any element, 
            //this.sourceRange = this.getTextRange("{table2}"); 
            //this.targetRange = this.getTable(0).Range;
            //this.move();

            //#Delete any element, 
            //this.targetRange = this.getTextRange("{table2}");
            //this.delete();


            //#Copy any element tha corresponds to comment, 
            //this.sourceRange = this.wordProcessor.Document.Comments[0].Range;            
            ////String test = ((NativeComment)this.wordProcessor.Document.Comments[0]).Comment.Content.PieceTable.TextBuffer.ToString();
            //this.targetRange = this.getTable(2).Range;
            //this.copy();

            //#Delete any element related to a specific comment. 
            //this.targetRange = this.wordProcessor.Document.Comments[0].Range;
            //delete();

            //#populate Table, this uses a dummy datasource at the moment
            //this.populateTable(this.getTable(1));
                                    

            //#Replace text with new text
            this.replaceTextWithNewText("{{Projects.ProjectName}}", datasource.GetValue("Projects.ProjectName").ToString());
            this.replaceTextWithNewText("{{Projects.Address1}}", datasource.GetValue("Projects.Address1").ToString());
            this.replaceTextWithNewText("{{Projects.SolutionEngineersSynopsis}}", datasource.GetValue("Projects.SolutionEngineersSynopsis").ToString());
            this.replaceTextWithNewText("{{Projects.SolutionPrintedYear}}", datasource.GetValue("Projects.SolutionPrintedYear").ToString());
            this.replaceTextWithNewText("{{Projects.TEECurrentVersion}}", datasource.GetValue("Projects.TEECurrentVersion").ToString());
            this.replaceTextWithNewText("{{Projects.TEESN}}", datasource.GetValue("Projects.TEESN").ToString());
            this.replaceTextWithNewText("{{Projects.SoftwareName}}", datasource.GetValue("Projects.SoftwareName").ToString());
            this.replaceTextWithNewText("{{Projects.EnergyBuildingRegistrationNumber}}", datasource.GetValue("Projects.EnergyBuildingRegistrationNumber").ToString());
            this.replaceTextWithNewText("{{Projects.EnergyBuildingVersion}}", datasource.GetValue("Projects.EnergyBuildingVersion").ToString());
            this.replaceTextWithNewText("{{Projects.EnergyBuildingSN}}", datasource.GetValue("Projects.EnergyBuildingSN").ToString());


            this.replaceTextWithNewText("{{BuildingsGeneral.CityID}}", datasource.GetValue("BuildingsGeneral.CityID").ToString());
            this.replaceTextWithNewText("{{BuildingsGeneral.Elevation}}", datasource.GetValue("BuildingsGeneral.Elevation").ToString());
            this.replaceTextWithNewText("{{BuildingsGeneral.ClimaticZoneName}}", datasource.GetValue("BuildingsGeneral.ClimaticZoneName").ToString());
            this.replaceTextWithNewText("{{PageCBuildings.RecNumber}}", datasource.GetValue("PageCBuildings.RecNumber").ToString());
            this.replaceTextWithNewText("{{PageCBuildings.Name}}", datasource.GetValue("PageCBuildings.Name").ToString());


            this.replaceTextWithNewText("{{SpecialAttributes.FT}}", MathOperations.formatTwoDecimalWithoutRound(datasource.GetValue("SpecialAttributes.FT").ToString()));
            this.replaceTextWithNewText("{{SpecialAttributes.FW}}", datasource.GetValue("SpecialAttributes.FW").ToString());
            this.replaceTextWithNewText("{{SpecialAttributes.FR}}", datasource.GetValue("SpecialAttributes.FW").ToString());
            this.replaceTextWithNewText("{{SpecialAttributes.FFB}}", datasource.GetValue("SpecialAttributes.FFB").ToString());
            this.replaceTextWithNewText("{{SpecialAttributes.FFU}}", datasource.GetValue("SpecialAttributes.FFU").ToString());
            this.replaceTextWithNewText("{{SpecialAttributes.FFA}}", datasource.GetValue("SpecialAttributes.FFA").ToString());
            this.replaceTextWithNewText("{{SpecialAttributes.FTU}}", datasource.GetValue("SpecialAttributes.FTU").ToString());
            this.replaceTextWithNewText("{{SpecialAttributes.FTB}}", datasource.GetValue("SpecialAttributes.FTB").ToString());
            this.replaceTextWithNewText("{{SpecialAttributes.FGF}}", datasource.GetValue("SpecialAttributes.FGF").ToString());

            this.replaceTextWithNewText("{{SpecialAttributes.F}}", MathOperations.formatTwoDecimalWithoutRound(datasource.GetValue("SpecialAttributes.F").ToString()));
            this.replaceTextWithNewText("{{SpecialAttributes.BuildingVolume}}", MathOperations.formatTwoDecimalWithoutRound(datasource.GetValue("SpecialAttributes.BuildingVolume").ToString()));
            this.replaceTextWithNewText("{{SpecialAttributes.FV}}", MathOperations.formatTwoDecimalWithoutRound(datasource.GetValue("SpecialAttributes.FV").ToString()));
            this.replaceTextWithNewText("{{SpecialAttributes.Umax}}", MathOperations.formatTwoDecimalWithoutRound(datasource.GetValue("SpecialAttributes.Umax").ToString()));

            String[,] contents = new String[11, 2] {
                { "Υπολογισμός συντελεστών θερμοπερατότητας αδιαφανών δομικών στοιχείων", "3"},
                { "1. Υπολογισμός συντελεστών θερμοπερατότητας αδιαφανών δομικών στοιχείων", "4"},
                { "2. Υπολογισμός ισοδύναμων συντελεστών θερμοπερατότητας αδιαφανών δομικών στοιχείων σε επαφή με το έδαφος", "11"},
                { "3. Υπολογισμός συντελεστών θερμοπερατότητας και συντελεστών ηλιακών κερδών  διαφανών δομικών στοιχείων", "12"},
                { "4. Κατακόρυφα αδιαφανή δομικά στοιχεία", "13"},
                { "5. Οριζόντια αδιαφανή δομικά στοιχεία", "21"},
                { "6. Διαφανή δομικά στοιχεία", "24"},
                { "7. Μη θερμαινόμενοι χώροι", "25"},
                { "8. Θερμογέφυρες", "26"},
                { "9. Υπολογισμός μέγιστου επιτρεπτού και πραγματοποιήσιμου Um του κτηρίου", "36"},
                { "10. Υπολογισμός αθέλητου αερισμού", "38"}            };

            //Report Type 1
            XmlNodeList list = ((Xml)datasource).getList("PageA");

            
            var test = ((Xml)datasource).getXmlDictionary();


            //this.replaceTextWithNewText("{{}}", datasource.GetValue("").ToString());

        }
        //Τhe only way to copy and paste something is via InsertDocumentContent method
        //https://supportcenter.devexpress.com/ticket/details/t725837/richeditdocumentserver-copy-paste-problem
        public void copy() {
            wordProcessor.Document.InsertDocumentContent(this.targetRange.End, this.sourceRange);
        }
        //Moves any table 
        public void move() {
            wordProcessor.Document.InsertDocumentContent(this.targetRange.End, this.sourceRange);
            this.targetRange = this.sourceRange;
            this.delete();
        }
        public void replaceTextWithNewTable(String text, int rows, int cols) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(text);
            this.wordProcessor.Document.Tables.Create(this.targetRange.Start, rows, cols);
            this.delete();
        }

        public void replaceTextWithNewText(String sourceText, String targetText) {
            this.wordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(sourceText);
            this.wordProcessor.Document.Replace(targetRange, targetText);
            //this.delete();
        }
        protected DocumentRange getElementRange() {
            return null;
        }
        public Table getTable(int position) {
            return this.wordProcessor.Document.Tables[position];
        }
        public Paragraph getParagraph(int position) {
            return this.wordProcessor.Document.Paragraphs[position];
        }
        public DocumentRange getTextRange(String search) {
            Regex myRegEx = new Regex(search);
            return this.wordProcessor.Document.FindAll(myRegEx).First();
        }
        public void delete() {
            this.wordProcessor.Document.Delete(this.targetRange);
        }
        private void populateTable(Table targetTable) {
            //foreach (TableData tabledata in this.datasource.getTableData()) {
            //    targetTable.BeginUpdate();
            //    this.addTableRows(targetTable, tabledata);
            //    targetTable.EndUpdate();
            //}
        }
        private void addTableRows(Table targetTable, TableData tabledata) {
            foreach (List<string> row in tabledata.Rows) {
                addTableRow(targetTable, row);
            }
        }
        private void addTableRow(Table targetTable, List<string> row) {
            int rowcount = targetTable.Rows.Count() - 1;
            targetTable.Rows.InsertAfter(rowcount);
            for (int i = 0; i < row.Count; i++) {
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, i].Range.Start, row[i]);
            }
        }


    }
}
