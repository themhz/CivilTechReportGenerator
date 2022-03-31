using ReportGenerator;
using System;

namespace ReportGenerator_v1.System {
    class App {
        public ProcessorDocx reportGenerator { set; get; }
        public App(ProcessorDocx _reportGenerator) {
            this.reportGenerator = _reportGenerator;
        }

        //Starts the application by using the report type that was passed by
        //program.cs
        //You pass in reportgenerator the class type of IReport
        public void start(IDocument reportType) {

            this.reportGenerator.CreateDocX(reportType);            
        }
    }
}
