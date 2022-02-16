using DevExpress.XtraRichEdit;

namespace ReportGenerator {
    public interface ITests {
        void run();
        void scanDocument(RichEditDocumentServer wordProcessor);
        void scanDocumentV2(RichEditDocumentServer wordProcessor);
        void test_CopyElement(RichEditDocumentServer wordProcessor);
        void test_CopyRow(RichEditDocumentServer wordProcessor);
        void test_CreateTableAfterAnElementOnTheDocument(RichEditDocumentServer wordProcessor);
        void test_deleteElement(RichEditDocumentServer wordProcessor);
    }
}