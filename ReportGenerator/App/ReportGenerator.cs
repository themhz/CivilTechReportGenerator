using ReportGenerator;
using System.Configuration;

namespace ReportGenerator_v1.System {
    class ReportGenerator {
        public IReport DocXReport { set; get; }

        public ReportGenerator(IReport _docXReport) {
            this.DocXReport = _docXReport;
        }

        public void CreateDocX(IReport reportType) {            
            this.DocXReport = reportType;
            this.DocXReport.template = ConfigurationManager.AppSettings["reportPath"] + "Main.docx";
            this.DocXReport.generatedFile = ConfigurationManager.AppSettings["reportPath"] + "report_generated.docx";
            this.DocXReport.fieldsFile = ConfigurationManager.AppSettings["fieldsPath"] + "fields.json";
            this.DocXReport.includesFile = ConfigurationManager.AppSettings["includesPath"] + "includes.json";
            this.DocXReport = this.DocXReport.create();
        }
    }
}
