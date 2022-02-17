using ReportGenerator;

namespace ReportGenerator_v1.System {
    class ReportGenerator {
        public IReport DocXReport { set; get; }

        public ReportGenerator(IReport _docXReport) {
            this.DocXReport = _docXReport;
        }

        public void CreateDocX(IReport reportType) {
            this.DocXReport = reportType;
            this.DocXReport.template = "c://Users//themis//Documents/Test.docx";
            this.DocXReport.generatedfile = "c://Users//themis//Documents/Test_Generated.docx";
            this.DocXReport = this.DocXReport.create();
        }

        

    }
}
