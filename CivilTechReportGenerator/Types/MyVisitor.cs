using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Types {
    public class MyVisitor : DocumentVisitorBase {
        readonly StringBuilder buffer;
        public MyVisitor() { this.buffer = new StringBuilder(); }
        public StringBuilder Buffer { get { return buffer; } }
        public string Text { get { return Buffer.ToString(); } }

        public override void Visit(DocumentText text) {
            string prefix = (text.TextProperties.FontBold) ? "**" : "";
            Buffer.Append(prefix);
            Buffer.AppendLine(text.Text);
            Buffer.Append(prefix);
        }
        public override void Visit(DocumentParagraphEnd paragraphEnd) {
            Buffer.AppendLine();
        }
    }
}
