using System.Collections.Generic;
using System.IO;
using FormattedExcelExport.Style;

namespace FormattedExcelExport.TableWriters {
    public interface ITableWriterSimple {
        void WriteHeader(IEnumerable<string> cells);
        void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells);
        void AutosizeColumns();
        MemoryStream GetStream();
    }
}
