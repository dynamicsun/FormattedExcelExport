using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using FormattedExcelExport.Style;


namespace FormattedExcelExport.TableWriters {
	public sealed class DsvTableWriterSimple : ITableWriterSimple {
		private readonly StringBuilder _stringBuilder = new StringBuilder();
		private readonly string _delimeter;
		public DsvTableWriterSimple(string delimeter = "\t") {
			_delimeter = delimeter;
		}

		public void WriteHeader(List<string> cells) {
			WriteRow(cells.ConvertAll(x => new KeyValuePair<string, TableWriterStyle>(x, null)));
		}

		public void WriteRow(List<KeyValuePair<string, TableWriterStyle>> cells) {
			int cellsCount = cells.Count() - 1;
			int i = 0;
			foreach (KeyValuePair<string, TableWriterStyle> cell in cells) {
				if (cell.Key != null)
					_stringBuilder.Append(cell.Key);

				if (i < cellsCount)
					_stringBuilder.Append(_delimeter);
				i++;
			}
			_stringBuilder.AppendLine();
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