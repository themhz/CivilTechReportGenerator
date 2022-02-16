using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using ReportGenerator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator_v1.System {

    class DevExpressDocX : IReport {

        public RichEditDocumentServer wordProcessor;

        public DevExpressDocX(RichEditDocumentServer _wordProcessor) {
            this.wordProcessor = _wordProcessor;
        }
        public IReport create() {
            Console.WriteLine("creating report");
            using (this.wordProcessor) {


            }
            return null;
        }

        public void delete() {
            throw new NotImplementedException();
        }

        public IImage image() {
            throw new NotImplementedException();
        }

        public IListElement list() {
            throw new NotImplementedException();
        }

        public IParagraphElement paragraph() {
            throw new NotImplementedException();
        }

        public ITableElement table() {
            throw new NotImplementedException();
        }

        public ITitleElement title() {
            throw new NotImplementedException();
        }
    }
}
