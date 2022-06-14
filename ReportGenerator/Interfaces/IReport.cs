using System;

namespace ReportGenerator {
    public interface IReport {
        String template { set; get; }
        String generatedFile { set; get; }
        String fieldsFile { set; get; }
        String includesFile { set; get; }
        IReport create();
        void save();
        void start();
        void delete();        
    }
}