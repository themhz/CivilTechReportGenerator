using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CivilTechReportGenerator.Handlers {
    public class SectionItem : DocumentX, ISectionItem {


        private String _text;

        public String text {
            get { return _text; }
            set { _text = value; }
        }

        public SectionItem(RichEditDocumentServer wordProcessor) : base(wordProcessor) {
        }

        public override void create() {

            Document document = this.wordProcessor.Document;


            MessageBox.Show(document.Sections.Count().ToString());
            document.InsertSection(document.Range.End);

            //more samples here https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.API.Native.Section#related-examples            
        }

        public void delete(int index) {
            this.wordProcessor.Document.Delete(this.wordProcessor.Document.Sections[index].Range);
        }

        public void replace(String generatedfile, int pos) {

            System.Text.RegularExpressions.Regex myRegEx = new System.Text.RegularExpressions.Regex("{{SECTION}}");
            DocumentRange dr = this.wordProcessor.Document.FindAll(myRegEx).First();
            DocumentPosition dpos = this.wordProcessor.Document.CreatePosition(dr.Start.ToInt());
            this.wordProcessor.Document.InsertText(dpos, " ");
            this.wordProcessor.Document.InsertDocumentContent(dpos, this.wordProcessor.Document.Sections[pos].Range);
            this.wordProcessor.Document.Delete(dr);
            base.saveDocument(generatedfile);
        }

        public override int count() {
            return this.wordProcessor.Document.Sections.Count;
        }
    }
}
