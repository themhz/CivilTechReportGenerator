using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CivilTechReportGenerator.Interfaces;
using CivilTechReportGenerator.Handlers;

namespace CivilTechReportGenerator {
    abstract class DocumentXItem : IDocumentItem {

        public int x;
        public DocumentXItem setX(int val) {
            x = val;
            return this;
        }

        public int y;
        public DocumentXItem setY(int val) {
            y = val;
            return this;
        }

        public int pos;
        public DocumentXItem setPos(int val) {
            pos = val;
            return this;
        }

    }
}
