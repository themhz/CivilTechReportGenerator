using ReportGenerator;

namespace ReportGenerator_v1.System {
    class ProcessorFactory {
        public IDocument DocXReport { set; get; }

        public ProcessorFactory(DocType docType, DocLib docLib) {
            //this.DocXReport = _docXReport;
        }

        //public void CreateDocX(IDocument reportType) {
        //    this.DocXReport = reportType;
        //    this.DocXReport.template = "c://Users//themis//Documents/report_template.docx";
        //    this.DocXReport.generatedfile = "c://Users//themis//Documents/report_template_generated.docx";
        //    this.DocXReport = this.DocXReport.create();
        //}
    }
}
