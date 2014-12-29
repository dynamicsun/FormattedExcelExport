using System.Collections.Generic;
using System.IO;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public interface ITableWriterComplex {
		void WriteHeader(IEnumerable<string> cells);
        void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells, bool prependDelimeter = false);
		void WriteChildHeader(params string[] cells);
        void WriteChildRow(IEnumerable<KeyValuePair<dynamic, TableWriterStyle>> cells, bool prependDelimeter = false);
		void AutosizeColumns();
		MemoryStream GetStream();
	}
}