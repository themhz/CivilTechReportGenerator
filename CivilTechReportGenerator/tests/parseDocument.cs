using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CivilTechReportGenerator.tests {
    class parseDocument {
       
        public void OpenDocument(string fileName) {
            if (System.IO.File.Exists(fileName)) {
                using (DevExpress.XtraRichEdit.RichEditDocumentServer srv = new DevExpress.XtraRichEdit.RichEditDocumentServer()) {
                    if (srv.LoadDocument(fileName)) {
                        Document doc = srv.Document;
                        DocumentIterator iterator = new DocumentIterator(doc, true);

                        string log = "";

                        while (iterator.MoveNext()) {
                            IDocumentElement element = iterator.Current;

                            var type = element.Type;
                            string txt = getElementName(element).Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");

                            log += string.Format("{0}: {1}\r\n", type.ToString(), txt);
                        }

                        System.IO.File.WriteAllText("c://Users//themis//Documents/Log.txt", log);
                    }
                }
            }
        }

        protected string getElementName(IDocumentElement element) {
            switch (element.Type) {
                case DocumentElementType.BookmarkStart: return ((DocumentBookmarkStart)element).Name;
                case DocumentElementType.BookmarkEnd: return ((DocumentBookmarkEnd)element).Name;
                case DocumentElementType.CheckBox: return ((DocumentBookmarkEnd)element).Name;
                case DocumentElementType.CommentStart: return ((DocumentCommentStart)element).Name;
                case DocumentElementType.CommentEnd: return ((DocumentCommentEnd)element).Name;
                case DocumentElementType.EndnoteCustomMark: return ((DocumentEndnoteCustomMark)element).Text;
                case DocumentElementType.EndnoteEmptyReference: return string.Empty;
                case DocumentElementType.EndnoteReference: return string.Empty;
                case DocumentElementType.FieldCodeStart: return string.Empty;
                case DocumentElementType.FieldCodeEnd: return string.Empty;
                case DocumentElementType.FieldResultEnd: return string.Empty;
                case DocumentElementType.FootnoteCustomMark: return ((DocumentFootnoteCustomMark)element).Text;
                case DocumentElementType.FootnoteEmptyReference: return string.Empty;
                case DocumentElementType.FootnoteReference: return string.Empty;
                case DocumentElementType.HyperlinkStart: return ((DocumentHyperlinkStart)element).NavigateUri;
                case DocumentElementType.HyperlinkEnd: return ((DocumentHyperlinkEnd)element).NavigateUri;
                case DocumentElementType.InlinePicture: return ((DocumentInlinePicture)element).Uri;
                case DocumentElementType.ParagraphStart: return string.Empty;
                case DocumentElementType.ParagraphEnd: return ((DocumentParagraphEnd)element).Text;
                case DocumentElementType.Picture: return ((DocumentPicture)element).Uri;
                case DocumentElementType.RangePermissionStart: return string.Empty;
                case DocumentElementType.RangePermissionEnd: return string.Empty;
                case DocumentElementType.SectionStart: return string.Empty;
                case DocumentElementType.SectionEnd: return ((DocumentSectionEnd)element).Text;
                case DocumentElementType.Shape: return ((Shape)element).AltText;
                case DocumentElementType.TableCellBorder: return String.Empty;
                case DocumentElementType.Text: return ((DocumentText)element).Text;
                case DocumentElementType.TextBox: return String.Empty;                
                default: return string.Empty;
            }
        }
       
    }
}
