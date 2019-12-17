namespace DataImport
{
    public class Document : Item
    {
        public const string FN_FILEREF = "FileRef";
        public const string FN_DOCICON = "DocIcon";
        public const string FN_FILELEAFREF = "FileLeafRef";
        public string FileRef { get; set; }
        public string DocIcon { get; set; }
        public string FileLeafRef { get; set; }
    }
}
