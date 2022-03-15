using ReportGenerator.Types;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.DataSources {
    interface IDataSource {
        object GetValue(string field, int index = 0);
        string GetValueByID(string field, string id = "");
    }
}
