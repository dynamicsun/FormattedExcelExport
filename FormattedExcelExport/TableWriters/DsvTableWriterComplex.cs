using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public sealed class DsvTableWriterComplex : ITableWriterComplex {
		private readonly StringBuilder _stringBuilder = new StringBuilder();
		private readonly string _delimeter;
		public DsvTableWriterComplex(string delimeter = "\t") {
			_delimeter = delimeter;
		}

	    public void WriteHeader(params string[] cells) {
			WriteRow(false, cells);
		}

	    public void WriteRow(List<KeyValuePair<dynamic, TableWriterStyle>> cells, bool prependDelimeter = false) {
			int cellsCount = cells.Count() - 1;
			int i = 0;
			if (prependDelimeter) _stringBuilder.Append(_delimeter);
			foreach (KeyValuePair<dynamic, TableWriterStyle> cell in cells) {
				_stringBuilder.Append(cell.Key);

				if (i < cellsCount)
					_stringBuilder.Append(_delimeter);
				i++;
			}
			_stringBuilder.AppendLine();
		}
		public void WriteRow(bool prependDelimeter, params string[] cells) {
			int cellsCount = cells.Length - 1;
			int i = 0;
			if (prependDelimeter) _stringBuilder.Append(_delimeter);
			foreach (string cell in cells) {
				_stringBuilder.Append(cell);

				if (i < cellsCount)
					_stringBuilder.Append(_delimeter);
				i++;
			}
			_stringBuilder.AppendLine();
		}
		public void WriteChildHeader(params string[] cells) {
			WriteHeader(cells);
		}
		public void WriteChildRow(IEnumerable<KeyValuePair<dynamic, TableWriterStyle>> cells, bool prependDelimeter = false) {
			WriteRow(cells.ToList(), prependDelimeter);
		}
		public void AutosizeColumns() { }
		public MemoryStream GetStream() {
			MemoryStream memoryStream = new MemoryStream();
			StreamWriter streamWriter = new StreamWriter(memoryStream, Encoding.UTF8);
			streamWriter.WriteLine(_stringBuilder.ToString());
			streamWriter.Flush();
			memoryStream.Position = 0;
			return memoryStream;
		}
	}
}