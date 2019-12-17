using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using Microsoft.SharePoint;
using System.Xml.Serialization;
using Newtonsoft.Json;
using LHR.Types;

namespace DataImport
{
    [JsonObject(MemberSerialization.OptIn)]
    [Serializable()]
    public class Item
    {
        public const string C_LIST_FIELDS_PROPERTY_NAME = "ListFields";
        private List<ESFieldValue> _listFields = new List<ESFieldValue>();
        public const string FN_ID = "ID";
        public const string FN_TITLE = "Title";
        public const string FN_VERSION = "_UIVersionString";
        public const string FN_PATH = "FileDirRef";
        public const string FN_MODIFIED = "Modified";
        public const string FN_CREATED = "Created";
        public const string FN_MODIFIED_BY = "Editor";
        public const string FN_ATTACHMENT = "Attachments";
        public const string FN_CREATED_BY = "Author";

        [XmlIgnore]
        public List<string> ItemColumns
        {
            get
            {
                var _knownColumns = new List<string>();
                _knownColumns.Add(FN_ID);
                _knownColumns.Add(FN_TITLE);
                _knownColumns.Add(FN_VERSION);
                _knownColumns.Add(FN_PATH);
                _knownColumns.Add(FN_MODIFIED);
                _knownColumns.Add(FN_CREATED);
                _knownColumns.Add(FN_CREATED_BY);
                return _knownColumns;
            }
        }

        [XmlIgnore]
        public virtual List<string> KnownColumns
        {
            get
            {
                return ItemColumns;
            }
        }

        [JsonProperty("I1")]
        public int ID { get; set; }
        [JsonProperty("I2")]
        public string Title { get; set; }
        [JsonProperty("I3")]
        public string Version { get; set; }
        [JsonProperty("I4")]
        public DateTime Modified { get; set; }
        [JsonProperty("I5")]
        public DateTime Created { get; set; }
        [JsonProperty("I6")]
        public string ModifiedBy { get; set; }
        [XmlIgnore]
        public object Tag { get; set; }

        public SPFieldLookupValue LookupValue
        {
            get
            {
                return new SPFieldLookupValue(ID, Title);
            }
        }

        [XmlElement("ListFields")]
        [JsonProperty("LF")]
        public List<ESFieldValue> ListFields
        {
            get
            {
                return _listFields;
            }
        }

        public string ListFieldValueString(string fieldName)
        {
            var fld = ListFields.FirstOrDefault(f => (f.InternalFieldName ?? "") == (fieldName ?? ""));

            if (fld != null && fld.DisplayValue != null)
                return fld.DisplayValue;
            else
                return string.Empty;
        }

        public object ListFieldValue(string fieldName)
        {
            var fld = ListFields.FirstOrDefault(f => (f.InternalFieldName ?? "") == (fieldName ?? ""));

            if (fld != null)
                return fld.Value;
            else
                return null;
        }

        public static int GetChoiceCode(string sValue)
        {
            if (!string.IsNullOrEmpty(sValue) && sValue.Length >= 1 && System.Text.RegularExpressions.Regex.IsMatch(sValue.Substring(0, 1), "^[0-9]"))
                return Convert.ToInt32(sValue.Substring(0, 1));
            else
                return -1;
        }

        public string GetChoiceValueByCode(int iCode, System.Collections.Specialized.StringCollection choices)
        {
            return choices.Cast<string>().FirstOrDefault(s => !string.IsNullOrEmpty(s) && s.StartsWith(iCode.ToString()));
        }

        public string GetChoiceValueByCode(int iCode, IEnumerable<string> choices)
        {
            return choices.FirstOrDefault(f => !string.IsNullOrEmpty(f) && f.StartsWith(iCode.ToString()));
        }

        public SPFieldLookupValue CreateLookupValue(object val)
        {
            if (val == null)
                return null;
            var result = new SPFieldLookupValue(val.ToString());
            return result.LookupId > -1 && result.LookupValue != null ? result : null;
        }

        public string GetStringFromLookupValue(SPFieldLookupValue val)
        {
            string result = string.Empty;

            if (val != null)
                result = val.ToString();

            return result;
        }

        public string GetStringFromLookupValueCollection(SPFieldLookupValueCollection val)
        {
            string result = string.Empty;

            if (val != null && val.Count > 0)
                result = val.ToString();

            return result;
        }

        public SerializableLookupValue CreateSerializableLookup(SPFieldLookupValue lookup)
        {
            if (lookup != null)
            {
                var srLkp = new SerializableLookupValue();
                srLkp.LookupId = lookup.LookupId;
                srLkp.LookupValue = lookup.LookupValue;
                return srLkp;
            }
            else
                return null;
        }

        public static string EncodeFieldNameSpecialSymbols(string fieldName)
        {
            return fieldName.Replace(" ", "_x0020_").Replace(":", "_x003a_").Substring(0, 30);
        }
    }
}
