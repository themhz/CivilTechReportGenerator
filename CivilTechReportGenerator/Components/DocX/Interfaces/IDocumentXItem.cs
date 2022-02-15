using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CivilTechReportGenerator.Interfaces {
    public abstract class IDocumentXItem {

        public abstract void create();
        public abstract int count();
        public abstract void delete(int index);
    }
}
