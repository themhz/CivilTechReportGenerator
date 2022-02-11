using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;
using CivilTechReportGenerator.Types;

namespace CivilTechReportGenerator {
    class DocxDevExpressHandler {
        public String template;
        public String filePath;
        public RichEditDocumentServer srv;
        public Document doc;
        public CharacterProperties cp;
        public ParagraphProperties pp;
        public Dictionary<string, string> fieldItems;        

        public DocxDevExpressHandler(String _template, String _filePath) {
            this.template = _template;
            this.filePath = _filePath;
            this.srv = new DevExpress.XtraRichEdit.RichEditDocumentServer();
            this.doc = srv.Document;
        }

        public void setFieldItems(Dictionary<string, string> items) {
            this.fieldItems = items;
        }


        
        //Populates Tables based on table key. It uses a custom type like List<TableData> to populate the table tha corresponds to the particular table key
        public void populateTable(List<TableData> tableItems) {
            var document = this.srv.Document;
            document.BeginUpdate();
            
            foreach(TableData td in tableItems) {
                //var table = document.Tables[int.Parse(td.TableKey)];

                document.Tables[int.Parse(td.TableKey)].BeginUpdate();
                foreach (List<string> row in td.Rows) {
                    for(int i=0; i < row.Count; i++) {

                        document.InsertSingleLineText(document.Tables[int.Parse(td.TableKey)][0, i].Range.Start, row[i]);
                    }

                }
                document.Tables[int.Parse(td.TableKey)].EndUpdate();
            }

            
            document.EndUpdate();
            document.SaveDocument(this.filePath, DocumentFormat.OpenXml);


        }
        public void startEditTableTest() {
            this.srv.Document.BeginUpdate();
            var document = this.srv.Document;

            document.Tables[1].BeginUpdate();
            document.Tables[1].Rows.InsertBefore(1);
            document.Tables[1].Rows.InsertAfter(1);
            document.Tables[1].Rows[1].Cells.Append();
            document.Tables[1].EndUpdate();

            document.EndUpdate();
            this.srv.SaveDocument(this.filePath, DocumentFormat.OpenXml);
            //MessageBox.Show(document.Tables.Count.ToString());
        }

        //
        public void startReplaceKeysInDoc() {                      
            this.srv.Document.BeginUpdate();
            foreach (KeyValuePair<string, string> fieldItem in this.fieldItems) {                
                System.Text.RegularExpressions.Regex myRegEx = new System.Text.RegularExpressions.Regex(fieldItem.Key);
                this.srv.Document.ReplaceAll(myRegEx, fieldItem.Value);                
            }
            this.srv.Document.EndUpdate();
            this.srv.SaveDocument(this.filePath, DocumentFormat.OpenXml);
        }

        //Loads the template to be used to create the new document repoert with the replaced key value pairs
        public void loadTemplate() {
            this.srv.LoadDocument(this.template);
        }

        //Loads and already existed docx document to it can be parsed.
        public void loadDocument() {
            this.srv.LoadDocument(this.filePath);
        }

        //Creates a new document on disk after memory manipulation of this.srv 
        public void createDocument() {
            this.srv.SaveDocument(this.filePath, DocumentFormat.OpenXml);
        }

        //You can set some properties for the entire docx document aka global properties
        public void setCharacterPropertiesInDocument() {
            this.cp = this.doc.BeginUpdateCharacters(this.doc.Paragraphs[0].Range);
            this.cp.ForeColor = Color.FromArgb(0x83, 0x92, 0x96);
            this.cp.Italic = true;
            this.doc.EndUpdateCharacters(this.cp);

        }

        //This is a test function that sets some paragraph properties. {To be updated or removed}
        public void setParagraphPropertiesInDocument() {
            this.pp = this.doc.BeginUpdateParagraphs(this.doc.Paragraphs[0].Range);
            this.pp.Alignment = ParagraphAlignment.Left;
            this.doc.EndUpdateParagraphs(this.pp);
        }

        //Appends some text on the bottom of document
        public void appendText(String text) {
            this.doc.AppendText(text);
        }

        //Inserts secion on the document
        public void insertSection() {
            this.doc.InsertSection(this.doc.Paragraphs[this.doc.Paragraphs.Count - 1].Range.End);
        }

        //Opens the document directly to the user by using word program
        public void showDocument() {
            System.Diagnostics.Process.Start(this.filePath);
        }

    }
}
