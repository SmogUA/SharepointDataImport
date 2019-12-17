namespace DataImport
{
    public class DataImportDocument : Document
    {
        public const string C_PROCESS_TYPE_VALIDATE = "Validate";
        public const string C_PROCESS_TYPE_IMPORT = "Import";

        public const string FN_SECTION_ITEM = "SectionItem";
        public const string FN_IS_IN_PROCESS = "IsInProcess";
        public const string FN_DATE_FORMAT = "DateFormat";
        public const string FN_PROCESS_TYPE = "ProcessType";
        public const string FN_SELECTED_SHEET = "SelectedSheet";
        public const string FN_IS_GENERATED = "IsGenerated";
        public const string FN_RESULTS = "Results";

        public Microsoft.SharePoint.SPFieldLookupValue SectionItem { get; set; }
        public bool IsInProcess { get; set; }
        public string DateFormat { get; set; }
        public string ProcessType { get; set; }
        public string SelectedSheet { get; set; }
        public bool IsGenerated { get; set; }
        public string Results { get; set; }
    }
}
