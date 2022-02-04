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


namespace CivilTechReportGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            testDevExpressReplaceKeys();
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
