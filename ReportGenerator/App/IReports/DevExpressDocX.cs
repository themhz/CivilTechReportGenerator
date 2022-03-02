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
using System.IO;
using System.Drawing;

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

            this.reportTemplate1();
            //this.replaceTextWithNewText("{{}}", datasource.GetValue("").ToString());
        }

        public void populatePageADetails(XmlNodeList DetailList, Table table) {
            DocumentRange r = this.getTextRange("{{PageADetails}}");

            int tableIndex = 0;
            foreach (XmlNode node in DetailList) {
                //List<string> rows = new List<string>();
                Dictionary<String, String> cols = new Dictionary<string, string>();
                foreach (XmlNode row in node) {
                    cols.Add(row.Name, row.InnerText);
                }                             
                this.addTableRow(table, cols);
                tableIndex++;                               
            }
         

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
            if(this.targetRange != null)
                this.wordProcessor.Document.Replace(targetRange, targetText);            
        }
        public void replaceTextWithImage(String sourceText, String targetText) {
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.Document.Unit = DevExpress.Office.DocumentUnit.Inch;
            this.targetRange = this.getTextRange(sourceText);
            if (this.targetRange != null) {                
                byte[] bytes = Convert.FromBase64String(targetText);                
                bytes = ImageResizer.resize(bytes, 700, 700);
                using (MemoryStream ms = new MemoryStream(bytes)) {                    
                    DocumentImageSource image = DocumentImageSource.FromStream(ms);                    
                    this.wordProcessor.Document.Images.Insert(this.targetRange.Start, image);
                }
                this.delete();
            }
            //this.wordProcessor.Document.Replace(targetRange, targetText);
        }        
        public Table getTable(int position) {
            return this.wordProcessor.Document.Tables[position];
        }
        public Paragraph getParagraph(int position) {
            return this.wordProcessor.Document.Paragraphs[position];
        }
        public DocumentRange getTextRange(String search) {
            try {
                Regex myRegEx = new Regex(search);
                return this.wordProcessor.Document.FindAll(myRegEx).First();
            } catch (Exception ex) {
                return null;
            }
            
        }
        public void delete() {
            this.wordProcessor.Document.Delete(this.targetRange);
        }
        private void addText(DocumentRange range, String value) {
            this.wordProcessor.Document.InsertSingleLineText(range.Start, value);
        }
        private void addTableRow(Table targetTable, Dictionary<String, String> cols) {
            int rowcount = targetTable.Rows.Count() - 1;
            targetTable.Rows.InsertAfter(rowcount);
            try {
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 0].Range.Start, cols["Index"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 1].Range.Start, cols["Name"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 2].Range.Start, cols["Density"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 3].Range.Start, cols["d"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 4].Range.Start, cols["λ"].ToString());
                this.wordProcessor.Document.InsertSingleLineText(targetTable[rowcount, 5].Range.Start, MathOperations.formatTwoDecimalWithoutRound(cols["dλ"].ToString(), 4));
            }catch(Exception ex) {

            }

        }
      

        private void reportTemplate1() {
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


           

            ////Report Type 1
            XmlNodeList PageAList = ((Xml)datasource).getList("PageA");
            //foreach PageA
            int counter = 0;
            foreach (XmlNode page in PageAList) {
                //Foreach child nodes of page A get the fields if ID, BuildingID, TypeID, RecNumber etc..                


                //Code that creates sections and clones templates
                DocumentRange startRange = this.getTextRange("{{START}}");
                DocumentRange endRange = this.getTextRange("{{END}}");
                DocumentRange totalRange = wordProcessor.Document.CreateRange(startRange.Start.ToInt(), endRange.End.ToInt());

                //Copy table to a new line

                wordProcessor.Document.InsertDocumentContent(wordProcessor.Document.InsertSection(endRange.End).Range.Start, totalRange);

                XmlNodeList DetailList = ((Xml)datasource).getList("PageADetails[ns:PageADetailID='" + page["ID"].InnerText + "']");
                this.populatePageADetails(DetailList, this.wordProcessor.Document.Tables[counter]);


                this.replaceTextWithImage("{{PageA." + page["Image"].Name + "}}", page["Image"].InnerText);
                this.replaceTextWithNewText("{{PageA.Name}}", page["Name"].InnerText);
                this.replaceTextWithNewText("{{PageA.ElementTypeCase}}", page["ElementTypeCase"].InnerText);
                this.replaceTextWithNewText("{{PageA.Ri}}", MathOperations.formatTwoDecimalWithoutRound(page["Ri"].InnerText,2));
                this.replaceTextWithNewText("{{PageA.R}}", MathOperations.formatTwoDecimalWithoutRound(page["R"].InnerText,2));
                this.replaceTextWithNewText("{{PageA.Ra}}", MathOperations.formatTwoDecimalWithoutRound(page["Ra"].InnerText,2));
                this.replaceTextWithNewText("{{PageA.Rall}}", MathOperations.formatTwoDecimalWithoutRound(page["Rall"].InnerText,2));



                //this.wordProcessor.Document.Delete(startRange);
                //this.wordProcessor.Document.Delete(endRange);



                counter++;
                if (counter == 4) {                    
                    this.wordProcessor.Document.Delete(totalRange);

                    for(int i = 0; i < counter; i++) {
                        startRange = this.getTextRange("{{START}}");
                        endRange = this.getTextRange("{{END}}");
                        this.wordProcessor.Document.Delete(startRange);
                        this.wordProcessor.Document.Delete(endRange);
                    }                    
                    break;
                }
            }
        }
    }
}
