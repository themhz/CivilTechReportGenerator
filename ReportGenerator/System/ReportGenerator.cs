using ReportGenerator;

namespace ReportGenerator_v1.System {
    class ReportGenerator {
        public IReport DocXReport { set; get; }

        public ReportGenerator(IReport _docXReport) {
            this.DocXReport = _docXReport;
        }

        public void CreateDocX(IReport reportType) {
            this.DocXReport = reportType;
            this.DocXReport.template = "C:\\Users\\themis\\source\\repos\\CivilTechReportGenerator\\ReportGenerator\\DataSources\\files\\report_template.docx";
            this.DocXReport.generatedfile = "C:\\Users\\themis\\source\\repos\\CivilTechReportGenerator\\ReportGenerator\\DataSources\\files\\report_template_generated.docx";
            this.DocXReport = this.DocXReport.create();
        }

        

    }
}
