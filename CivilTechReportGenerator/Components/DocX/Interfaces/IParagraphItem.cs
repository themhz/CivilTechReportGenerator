using DevExpress.XtraRichEdit.API.Native;
using System.Text.RegularExpressions;

namespace CivilTechReportGenerator.Handlers {
    public interface IParagraphItem {
        int count();
        void create();
        void delete(int index);
        Table findParagraph(int pos);
        void replace(int pos, Regex _myRegEx);
    }
}