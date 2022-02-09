using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;


using DevExpress.Spreadsheet;
using System.IO;
using CivilTechReportGenerator.Types;
using CivilTechReportGenerator.Handlers;
using DevExpress.XtraRichEdit;
using CivilTechReportGenerator.tests;

namespace CivilTechReportGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int a = 0;
        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            RichEditDocumentServer wordProcessor = new RichEditDocumentServer();
            using (wordProcessor) {
                parseDocument();


            }
            
        }
        public void testCopyRow(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test-2.docx";
            String generatedfile = "c://Users//themis//Documents/Test-2_copy.docx";

            TableHandler th = new TableHandler(wordProcessor);
            th.loadTemplate(template);
            th.copyRow(1, 1, 3, generatedfile);
            th.copyRow(1, 1, 4, generatedfile);


        }


        public void testCopyElement(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/Test-2.docx";
            String generatedfile = "c://Users//themis//Documents/Test-2_copy.docx";

            TableHandler th = new TableHandler(wordProcessor);
            th.loadTemplate(template);
            th.copy(0, 14, generatedfile);

        }

        private void parseDocument() {
            parseDocument pd = new parseDocument();
            pd.OpenDocument("c://Users//themis//Documents/test1.docx");
        }

        private void countTables(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/ΠαράρτημαVI_Template.docx";
            String generatedfile = "c://Users//themis//Documents/ΠαράρτημαVI_Template2.docx";

            TableHandler th = new TableHandler(wordProcessor);
            th.loadTemplate(template);
            
            th.countTables();
        }

        private void testPopulateTable(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/ΠαράρτημαVI_Template.docx";
            String generatedfile = "c://Users//themis//Documents/ΠαράρτημαVI_Template2.docx";

            TableHandler th = new TableHandler(wordProcessor);
            th.loadTemplate(template);


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

            th.loadTemplate(template);
            th.beginUpdate();
            th.populateTable(tds);
            th.saveDocument(generatedfile);

        }

        private void createSection(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/test3.docx";
            SectionHandler sh = new SectionHandler(wordProcessor);
            sh.loadTemplate(template);
            sh.create();

        }
        private void createParagraph(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/test3.docx";
            ParagraphHandler ph = new ParagraphHandler(wordProcessor);
            ph.text = "dasdsadsa";
            ph.x = 0;
            ph.y = 1;
            
            ph.loadTemplate(template);

            ph.create();
            
       }
        private void createTable(RichEditDocumentServer wordProcessor) {
            String template = "c://Users//themis//Documents/test3.docx";
            TableHandler tw = new TableHandler(wordProcessor);
            tw.loadTemplate(template);
            
            tw.create();            
            MessageBox.Show("Tables :" + tw.countTables().ToString() + " at position "+ a);
            a = a + 20;
        }       

        private void testDevExpressReplaceKeys()
        {        
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
        private void testDevExpressLoadWord()
        {
            String file = "c://Users//themis//Documents/Test.docx";
            DocxDevExpressHandler dh = new DocxDevExpressHandler(file, file);
            dh.loadDocument();
            dh.startEditTableTest();
        }
        private void testDevExpressWord()
        {
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

        private void testXml()
        {
            //d:\ProjNet2022\applications\Building.Project\Building.UI\data.el\databases\building\


            String path = "d:\\ProjNet2022\\applications\\Building.EnergyProject\\EnergyBuilding.UI\\data\\databases\\building\\richDocumentsData.EnergyBuilding131.ctdatabase";

            XmlHandler xh = new XmlHandler(path);
            String xml = xh.getXml();
        }
      
    }
}
