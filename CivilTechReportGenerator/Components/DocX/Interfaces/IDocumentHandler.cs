using System.Collections.Generic;

namespace CivilTechReportGenerator.Handlers {
    public interface IDocumentHandler {
        int count();
        List<string[]> countElements();
        void create();
        void deleteElement(DocumentXItem item, int index);
        DocumentXItem getDocumentItem();
        string scanDocument();
        DocumentHandler setDocumentItem(DocumentXItem _item);
    }
}