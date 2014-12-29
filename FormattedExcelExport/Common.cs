namespace FormattedExcelExport {
	internal static class Common {
		public static string GetHeader(string propertyName, string attributePropertyName) {
			return string.IsNullOrWhiteSpace(attributePropertyName) ? propertyName : attributePropertyName;
		}
	}
}
