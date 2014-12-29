using System.IO;
using System.Text;

namespace FormattedExcelExport.TableWriters {
    public abstract class DsvTableWriteBase {
        protected readonly StringBuilder _stringBuilder = new StringBuilder();
        public MemoryStream GetStream() {
            var memoryStream = new MemoryStream();
            var streamWriter = new StreamWriter(memoryStream, Encoding.UTF8);
            streamWriter.WriteLine(_stringBuilder.ToString());
            streamWriter.Flush();
            memoryStream.Position = 0;
            return memoryStream;
        }
    }
}
