using System;

namespace ReportGenerator {
    public interface IReport {
        String template { set; get; }
        String generatedfile { set; get; }
        IReport create();
        void save();
        void parse();
        void delete();        
    }
}