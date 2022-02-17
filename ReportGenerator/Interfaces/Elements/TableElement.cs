using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Text.RegularExpressions;

namespace ReportGenerator.Interfaces.Elements {
    class TableElement : ITableElement {
        public RichEditDocumentServer wordProcessor { set; get; }
        public Regex regex { set; get; }
        public DocumentRange documentRange { set; get; }
        public DocumentRange targetDocumentRange { set; get; }
        public DocumentPosition documentPosition { set; get; }
        public int tableIndex { set; get; }
        public ITableElement setTableIndex(int val) {
            tableIndex = val;
            return this;
        }
        public int rows { set; get; }
        public TableElement setRows(int val) {
            rows = val;
            return this;
        }
        public int cols { set; get; }
        public TableElement setCols(int val) {
            cols = val;
            return this;
        }

        public TableElement(RichEditDocumentServer _wordProcessor) {
            this.wordProcessor = _wordProcessor;
        }        

        public ITableElement copy(int index) {
            this.documentRange = this.wordProcessor.Document.Tables[index].Range;
            return this;
        }

        public ITableElement paste() {
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.Document.Tables.Create(this.targetDocumentRange.End, this.rows, this.cols);

            return this;
        }


        public ITableElement create() {
            throw new NotImplementedException();
        }

        public ITableElement delete() {
            throw new NotImplementedException();
        }

       
        public ITableElement tableElementItem() {
            throw new NotImplementedException();
        }

        public int count() {
            return this.wordProcessor.Document.Tables.Count;
        }
    }
}
