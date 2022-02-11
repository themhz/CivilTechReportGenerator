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
    class SectionHandler : CivilDocumentX {


        private String _text;

        public String text {
            get { return _text; }
            set { _text = value; }
        }






        public SectionHandler(RichEditDocumentServer wordProcessor) {
            base.srv = wordProcessor;
        }

        public override void create() {

            Document document = this.srv.Document;

            
            MessageBox.Show(document.Sections.Count().ToString());
            document.InsertSection(document.Range.End);

            //more samples here https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.API.Native.Section#related-examples



            base.saveDocument();

        }


        public void delete(int index, String generatedfile) {

            base.document.Delete(base.document.Sections[index].Range);
            base.saveDocument(generatedfile);
        }

        public void replace(String generatedfile, int pos) {

            System.Text.RegularExpressions.Regex myRegEx = new System.Text.RegularExpressions.Regex("{{SECTION}}");
            DocumentRange dr = this.srv.Document.FindAll(myRegEx).First();
            DocumentPosition dpos = document.CreatePosition(dr.Start.ToInt());
            document.InsertText(dpos, " ");
            base.document.InsertDocumentContent(dpos, base.document.Sections[pos].Range);
            base.document.Delete(dr);
            base.saveDocument(generatedfile);
        }

        public override int count() {
            return base.document.Sections.Count;
        }
    }
}
