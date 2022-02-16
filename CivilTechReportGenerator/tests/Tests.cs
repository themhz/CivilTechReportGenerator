using ReportGenerator.Handlers;
using ReportGenerator.tests;
using ReportGenerator.Types;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ReportGenerator {
    public class Tests : ITests {

        public String template { set; get; }
        public String generatedfile { set; get; }

        public RichEditDocumentServer wordProcessor { set; get; }
        public TableItem tableItem { set; get; }
        public SectionItem sectionItem { set; get; }
        public DocumentHandler documentHandler { set; get; }
        public TableData tableData { set; get; }
        public Tests( RichEditDocumentServer _wordProcessor, TableItem _tableItem, SectionItem _sectionItem, 
            DocumentHandler _documentHandler, TableData _tableData) {

            this.wordProcessor = _wordProcessor;
            this.tableItem = _tableItem;
            this.documentHandler = _documentHandler;
            this.tableData = _tableData;
            this.sectionItem = _sectionItem;

            this.template = "c://Users//themis//Documents/Test.docx";
            this.generatedfile = "c://Users//themis//Documents/Test_copy.docx";
    }

        public void run() {
            using (wordProcessor) {
                test_CopyElement(wordProcessor);
            }
        }


        public void test_CreateTableAfterAnElementOnTheDocument(RichEditDocumentServer wordProcessor) {
            
            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.wordProcessor.LoadDocument(this.template);
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
            this.documentHandler.saveDocument(this.generatedfile);          
        }
        public void test_deleteElement(RichEditDocumentServer wordProcessor) {
            
            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.loadTemplate(this.template);
            this.documentHandler.deleteElement(this.tableItem, 1);
            this.documentHandler.saveDocument(this.generatedfile);

        }
        public void test_CopyRow(RichEditDocumentServer wordProcessor) {
            
            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.loadTemplate(this.template);
            this.tableItem.copyRow(1, 0, 1);           
            this.documentHandler.saveDocument(this.generatedfile);

        }
        private void test_countTables(RichEditDocumentServer wordProcessor) {
            
            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.loadTemplate(this.generatedfile);
            int i = this.documentHandler.setDocumentItem(this.tableItem).getDocumentItem().count();
            MessageBox.Show(i.ToString());

        }
        private void test_populateTable(RichEditDocumentServer wordProcessor) {            

            this.documentHandler.wordProcessor = wordProcessor;
            this.tableItem.wordProcessor = wordProcessor;

            this.documentHandler.setDocumentItem(this.tableItem);
            this.documentHandler.loadTemplate(this.template);


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

            this.documentHandler.saveDocument(this.generatedfile);

        }

        private void test_createSection(RichEditDocumentServer wordProcessor) {            

            this.documentHandler.wordProcessor = wordProcessor;
            this.sectionItem.wordProcessor = wordProcessor;

            this.sectionItem.loadTemplate(this.template);
            this.sectionItem.create();

            this.documentHandler.saveDocument(this.generatedfile);

        }


        public void test_CopyElement(RichEditDocumentServer wordProcessor) {

            //this.documentHandler.loadTemplate(this.template);
            //this.tableItem.dpos = this.documentHandler.wordProcessor.Document.CreatePosition(2);            
            //this.tableItem.copy(1, 2);
            //this.documentHandler.saveDocument(this.generatedfile);


            wordProcessor.Document.LoadDocument(this.template, DocumentFormat.OpenXml);
            DocumentRange myRange = wordProcessor.Document.CreateRange(0, 120);

            //DocumentRange NewRange = wordProcessor.Document.CreateRange(10, 120);
            //DocumentRange myRange = wordProcessor.Document.Selection;
            wordProcessor.Document.Copy(myRange);
            wordProcessor.Document.Paste(DocumentFormat.PlainText);

            this.documentHandler.wordProcessor = wordProcessor;
            this.documentHandler.saveDocument(this.generatedfile);
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
        private void parseDocument() {
            parseDocument pd = new parseDocument();
            pd.OpenDocument("c://Users//themis//Documents/Test.docx");            
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

    
        
    }
}
