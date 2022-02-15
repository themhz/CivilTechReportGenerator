using CivilTechReportGenerator.Handlers;
using CivilTechReportGenerator.tests;
using CivilTechReportGenerator.Types;
using DevExpress.XtraRichEdit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace CivilTechReportGenerator {
    public class App : IApp {

        public RichEditDocumentServer wordProcessor { set; get; }
        public TableItem tableItem { set; get; }
        public DocumentHandler documentHandler { set; get; }
        public TableData tableData { set; get; }
        public App(RichEditDocumentServer _wordProcessor, 
            TableItem _tableItem, 
            DocumentHandler _documentHandler, 
            TableData _tableData) {
            this.wordProcessor = _wordProcessor;
            this.tableItem = _tableItem;
            this.documentHandler = _documentHandler;
            this.tableData = _tableData;
        }

        public void run() {
            //RichEditDocumentServer wordProcessor = new RichEditDocumentServer();
            using (wordProcessor) {

                test_populateTable(wordProcessor);
            }
        }


        public void test_CreateTableAfterAnElementOnTheDocument(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";

            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.wordProcessor.LoadDocument(template);
            //int itemPosition = wordProcessor.Document.Paragraphs[3].Range.End.ToInt();
            int itemPosition = this.documentHandler.wordProcessor.Document.Tables[1].Range.End.ToInt();
            //int itemPosition = wordProcessor.Document.Sections[3].Range.End.ToInt();
            //int itemPosition = wordProcessor.Document.NumberingLists[3].Range.End.ToInt();
            
            this.documentHandler
                .setDocumentItem(this.tableItem
                .setCols(2)
                .setRows(3)
                );

            this.documentHandler.beginUpdate();
            this.documentHandler.getDocumentItem().setDocumentPosition(itemPosition).create();
            this.documentHandler.saveDocument(generatedfile);          
        }
        public void test_deleteElement(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";
            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.loadTemplate(template);
            this.documentHandler.deleteElement(this.tableItem, 1);
            this.documentHandler.saveDocument(generatedfile);

        }
        public void test_CopyRow(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";
            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.loadTemplate(template);
            this.tableItem.copyRow(1, 0, 1);           
            this.documentHandler.saveDocument(generatedfile);

        }
        private void test_countTables(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test.docx";
            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.loadTemplate(generatedfile);
            int i = this.documentHandler.setDocumentItem(this.tableItem).getDocumentItem().count();
            MessageBox.Show(i.ToString());

        }
        private void test_populateTable(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";

            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.setDocumentItem(this.tableItem);
            this.documentHandler.loadTemplate(template);


            this.tableData.TableKey = "1";
            List<string> row1 = new List<string> { "col1", "col2", "col3", "col4" };
            List<string> row2 = new List<string> { "col1", "col2", "col3", "col4" };
            List<string> row3 = new List<string> { "col1", "col2", "col3", "col4" };
            List<string> row4 = new List<string> { "col1", "col2", "col3", "col4" };
            this.tableData.Rows.Add(row1);
            this.tableData.Rows.Add(row2);
            this.tableData.Rows.Add(row3);
            this.tableData.Rows.Add(row4);

            List<TableData> tds = new List<TableData>();
            tds.Add(this.tableData);



            this.tableItem.populateTable(tds);            

            this.documentHandler.saveDocument(generatedfile);

        }


        public void scanDocumentV2(RichEditDocumentServer wordProcessor) {
            //String template = "c://Users//themis//Documents/Test.docx";
            //String generatedfile = "c://Users//themis//Documents/Test_copy.docx";

            //DocumentHandler dh = new DocumentHandler(wordProcessor);
            //dh.loadTemplate(template);

            //List<String[]> countElements = dh.countElements();

            //foreach(var element in countElements) {
            //    MessageBox.Show(element[0] + element[1]);
            //}



        }
        public void scanDocument(RichEditDocumentServer wordProcessor) {
            //String template = "c://Users//themis//Documents/Test.docx";
            //String generatedfile = "c://Users//themis//Documents/Test_copy.docx";

            //DocumentHandler dh = new DocumentHandler(wordProcessor);
            //dh.loadTemplate(template);

            //String text = dh.scanDocument();
            //memoEdit1.Text = text;
        }
        public void testCopyElement(RichEditDocumentServer wordProcessor) {
            //String template = "c://Users//themis//Documents/Test-2.docx";
            //String generatedfile = "c://Users//themis//Documents/Test-2_copy.docx";

            //TableItem th = new TableItem(wordProcessor);
            //th.loadTemplate(template);
            //th.copy(0, 14, generatedfile);

        }
        private void parseDocument() {
            parseDocument pd = new parseDocument();
            pd.OpenDocument("c://Users//themis//Documents/Test.docx");
        }


        private void createSection(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/test3.docx";
            SectionItem sh = new SectionItem(wordProcessor);
            sh.loadTemplate(template);
            sh.create();

        }
        private void createParagraph(RichEditDocumentServer wordProcessor) {
            //String template = "c://Users//themis//Documents/test3.docx";
            //ParagraphItem ph = new ParagraphItem(wordProcessor);
            //ph.text = "dasdsadsa";
            //ph.x = 0;
            //ph.y = 1;

            //ph.loadTemplate(template);

            //ph.create();

        }
        private void createTable(RichEditDocumentServer wordProcessor) {
            //String template = "c://Users//themis//Documents/test3.docx";
            //TableItem tw = new TableItem(wordProcessor);
            //tw.loadTemplate(template);

            //tw.create();
            //MessageBox.Show("Tables :" + tw.count().ToString() + " at position " + a);
            //a = a + 20;
        }
        private void testDevExpressReplaceKeys() {
            String template = "c://Users//themis//Documents/ΠαράρτημαVI_Template.docx";
            String generatedfile = "c://Users//themis//Documents/ΠαράρτημαVI_Template2.docx";
            DocxDevExpressHandler dh = new DocxDevExpressHandler(template, generatedfile);
            Dictionary<string, string> fieldItems = new Dictionary<string, string>();
            fieldItems.Add("#ΠΑΡΑΣΤΑΤΙΚΑ#", "test1");
            fieldItems.Add("#ΠΑΡΕΜΒΑΣΕΙΣ1#", "test 2");
            fieldItems.Add("#ΠΕΑ#", "test 3");
            fieldItems.Add("#ΠΑΡΕΜΒΑΣΕΙΣ2#", " τεστ 4");
            fieldItems.Add("#ΣΤ_Γ#", "τεστ 5");
            fieldItems.Add("#ΣΤ_Δ1#", "τεστ 6 ");
            fieldItems.Add("#ΣΤ_Δ2#", "τεστ 7");
            fieldItems.Add("#ΣΤ_Δ3#", "test8");
            fieldItems.Add("#ΣΤ_Ε1#", "τεστ 9");

            dh.setFieldItems(fieldItems);
            dh.loadTemplate();
            dh.startReplaceKeysInDoc();
        }
        private void testDevExpressLoadWord() {
            String file = "c://Users//themis//Documents/Test.docx";
            DocxDevExpressHandler dh = new DocxDevExpressHandler(file, file);
            dh.loadDocument();
            dh.startEditTableTest();
        }
        private void testDevExpressWord() {
            String file = "c://Users//themis//Documents/Test.docx";
            DocxDevExpressHandler dh = new DocxDevExpressHandler(file, file);
            dh.setCharacterPropertiesInDocument();
            dh.setParagraphPropertiesInDocument();
            dh.insertSection();
            dh.appendText("hello world");
            dh.insertSection();
            dh.appendText("hello world2");
            dh.insertSection();
            dh.appendText("hello world3");
            dh.createDocument();
            //dh.showDocument();
        }

        private void testXml() {
            //d:\ProjNet2022\applications\Building.Project\Building.UI\data.el\databases\building\

            String path = "d:\\ProjNet2022\\applications\\Building.EnergyProject\\EnergyBuilding.UI\\data\\databases\\building\\richDocumentsData.EnergyBuilding131.ctdatabase";
            XmlDocument _doc = new XmlDocument();
            XmlHandler xh = new XmlHandler(path, _doc);
            String xml = xh.getXml();
        }
    }
}
