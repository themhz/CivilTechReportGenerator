using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator {
    class ReportGenerator {
        private String _templatePath;
        public String templatePath {
            get { return _templatePath; }
            set { _templatePath = value; }
        }

        private String _reportPath;
        public String reportPath {
            get { return _reportPath; }
            set { _reportPath = value; }
        }

        public ReportGenerator() {

        }


        public IReport create() {
            
            return null;
        }


    }
}
