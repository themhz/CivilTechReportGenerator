using ReportGenerator.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.DataSources {
    interface IDataSource {
        List<List<string>> getData();
        List<TableData> getTableData();
    }
}
