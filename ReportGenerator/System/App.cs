using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator_v1.System {
    class App {
        public ReportGenerator reportGenerator { set; get; }
        public App(ReportGenerator _reportGenerator) {
            this.reportGenerator = _reportGenerator;
        }

        public void start() {
            this.reportGenerator.CreateDocX();
        }
    }
}
