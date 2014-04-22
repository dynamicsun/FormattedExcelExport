using System.Collections.Generic;
using System.IO;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public interface ITableWriterComplex {
		void WriteHeader(params string[] cells);
		void WriteRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false);
		void WriteChildHeader(params string[] cells);
		void WriteChildRow(IEnumerable<KeyValuePair<string, TableWriterStyle>> cells, bool prependDelimeter = false);
		void AutosizeColumns();
		MemoryStream GetStream();
	}
}