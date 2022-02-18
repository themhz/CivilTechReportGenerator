using ReportGenerator;
using System;

namespace ReportGenerator_v1.System {
    class App {
        public ReportGenerator reportGenerator { set; get; }
        public App(ReportGenerator _reportGenerator) {
            this.reportGenerator = _reportGenerator;
        }

        //Starts the application by using the report type that was passed by
        //program.cs
        //You pass in reportgenerator the class type of IReport
        public void start(IReport reportType) {

            this.reportGenerator.CreateDocX(reportType);
            //Console.ReadLine();
        }
    }
}
