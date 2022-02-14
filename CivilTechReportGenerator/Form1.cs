using CivilTechReportGenerator.Handlers;
using CivilTechReportGenerator.tests;
using CivilTechReportGenerator.Types;
using DevExpress.XtraRichEdit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace CivilTechReportGenerator {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }
        int a = 0;
        private void button1_Click(object sender, EventArgs e) {
            RichEditDocumentServer wordProcessor = new RichEditDocumentServer();
            using (wordProcessor) {
                test_populateTable(wordProcessor);              
            }
        }

        public void test_CreateTableAfterAnElementOnTheDocument(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";

            wordProcessor.LoadDocument(template);
            //int itemPosition = wordProcessor.Document.Paragraphs[3].Range.End.ToInt();
            int itemPosition = wordProcessor.Document.Tables[2].Range.End.ToInt();
            //int itemPosition = wordProcessor.Document.Sections[3].Range.End.ToInt();
            //int itemPosition = wordProcessor.Document.NumberingLists[3].Range.End.ToInt();

            DocumentHandler dh = new DocumentHandler(wordProcessor)
                .setDocumentItem(new TableItem(wordProcessor)
                .setCols(2)
                .setRows(3)
                );
            
            dh.beginUpdate();
            dh.getDocumentItem().setDocumentPosition(itemPosition).create();
            dh.saveDocument(generatedfile);
        }
        public void test_deleteElement(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";
            DocumentHandler dh = new DocumentHandler(wordProcessor);
            
            dh.loadTemplate(template);
            dh.deleteElement(new TableItem(wordProcessor), 1);
            dh.saveDocument(generatedfile);

        }
        public void test_CopyRow(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";
            DocumentHandler dh = new DocumentHandler(wordProcessor);
            dh.loadTemplate(template);

            TableItem th = new TableItem(wordProcessor);
            //TableIndex, RowIndex, NewRowIndex
            th.copyRow(1, 0, 2);
            //th.copyRow(1, 1, 3);

            dh.saveDocument(generatedfile);

        }
        private void test_countTables(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";
            DocumentHandler dh = new DocumentHandler(wordProcessor);
            dh.loadTemplate(generatedfile);
            int i = dh.setDocumentItem(new TableItem(wordProcessor)).getDocumentItem().count();            
            MessageBox.Show(i.ToString());
            
        }
        private void test_populateTable(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";
            DocumentHandler dh = new DocumentHandler(wordProcessor);
            
            dh.setDocumentItem(new TableItem(wordProcessor));
            dh.loadTemplate(template);

            //TableItem th = new TableItem(wordProcessor);
            //th.loadTemplate(template);


            TableData td = new TableData();
            td.TableKey = "2";
            List<string> row1 = new List<string> { "col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8", "col9", "col10", "col11" };
            List<string> row2 = new List<string> { "col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8", "col9", "col10", "col11" };
            List<string> row3 = new List<string> { "col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8", "col9", "col10", "col11" };
            List<string> row4 = new List<string> { "col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8", "col9", "col10", "col11" };
            td.Rows.Add(row1);
            td.Rows.Add(row2);
            td.Rows.Add(row3);
            td.Rows.Add(row4);
            List<TableData> tds = new List<TableData>();
            tds.Add(td);


            //dh.getDocumentItem().beginUpdate();
            //th.populateTable(tds);
            //th.saveDocument(generatedfile);

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
            String template = "c://Users//themis//Documents/Test.docx";
            String generatedfile = "c://Users//themis//Documents/Test_copy.docx";

            DocumentHandler dh = new DocumentHandler(wordProcessor);
            dh.loadTemplate(template);
            
            String text = dh.scanDocument();
            memoEdit1.Text = text;            
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
