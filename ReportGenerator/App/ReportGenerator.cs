using ReportGenerator;

namespace ReportGenerator_v1.System {
    class ReportGenerator {
        public IReport DocXReport { set; get; }

        public ReportGenerator(IReport _docXReport) {
            this.DocXReport = _docXReport;
        }

        public void CreateDocX(IReport reportType) {
            string reportPath = "C:\\Users\\themis\\source\\repos\\CivilTechReportGenerator\\ReportGenerator\\DataSources\\files\\";
            this.DocXReport = reportType;
            this.DocXReport.template = reportPath + "report_template2.docx";
            this.DocXReport.generatedfile = reportPath+ "report_template_generated2.docx";
            this.DocXReport = this.DocXReport.create();
        }

        

    }
}
