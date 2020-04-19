using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Text;
using Microsoft.SharePoint;
using System.Collections;
using System.Globalization;

namespace DataImport
{
    public class BLDataImport
    {
        private const string ErrEmpty = " is empty;";
        private const string ErrNotFound = " can not be found;";
        private const string ErrWrongType = " has wrong type;";
        private const string ErrSave = " adding/updating item error;";
        private const string ErrTheSame = " {0} and {1} cannot be the same;";
        private const string multivalueDelimiter = ";";
        private List<string> standardFieldsCollection = new List<string>();
        private List<SPField> nonStandardListFields = new List<SPField>();
        private Dictionary<string, object> Relations = new Dictionary<string, object>();
        private readonly SharePoint SP = new SharePoint();
        private Guid _employeesListGUID;
        public const string FN_GLOBAL_ID = "GlobalEmployeeID";
        private System.Globalization.DateTimeFormatInfo df = new System.Globalization.DateTimeFormatInfo();
        private string err;
        private Dictionary<int, string> Errors = new Dictionary<int, string>();
        private int NumberOfItems = 0;
        private bool CheckBoolean(string val)
        {
            val = val.ToLower();
            if ((val ?? "") == "yes" || (val ?? "") == "y" || (val ?? "") == "true" || (val ?? "") == "1")
                return true;
            else
                return false;
        }
        private DateTime ParseDate(string val, string format)
        {
            DateTime dt;
            DateTime.TryParseExact(val, format, CultureInfo.InvariantCulture,   DateTimeStyles.None, out dt);
            if (dt.Year <= 1900 || dt.Year >= 2100)
                return default(DateTime);
            return dt;
        }
        private SPList GetEmpoloyeesList(string url)
        {
            return SP.GetListByInternalName(url, "Employees");
        }

        private void FillNonStandardFieldsRelations(List<DIMapping> mapping, SPWeb web, SPList lst)
        {
            Relations.Clear();
            var nonStandardFields = mapping.Where(m => standardFieldsCollection.All(sfc => (sfc ?? "") != (m.Name ?? ""))).ToList();
            if (nonStandardFields.Count > 0)
            {
                var nonStandardListFieldsCollection = lst.Fields;
                foreach (var map in nonStandardFields)
                {
                    var fld = nonStandardListFields.Where(sf => (sf.InternalName ?? "") == (map.Name ?? "")).FirstOrDefault();
                    if (fld == null)
                    {
                        fld = nonStandardListFieldsCollection.GetFieldByInternalName(map.Name);
                        nonStandardListFields.Add(fld);
                    }
                    if (fld != null)
                    {
                        switch (fld.Type)
                        {
                            case SPFieldType.Lookup:
                                {
                                    //AddLookupRelation(fld, map.Name, web);
                                    break;
                                }

                            case SPFieldType.Choice:
                                {
                                    var field = lst.Fields.GetFieldByInternalName(map.Name);
                                    System.Collections.Specialized.StringCollection choices;
                                    if (field != null)
                                    {
                                        choices = ((SPFieldChoice)field).Choices;
                                        Relations.Add(map.Name, choices);
                                    }

                                    break;
                                }

                            case SPFieldType.Invalid:
                                {
                                    //if (fld.FieldValueType == typeof(SPFieldLookupValue) | fld.FieldValueType == typeof(SPFieldLookupValueCollection))
                                       // AddLookupRelation(fld, map.Name, web);
                                    break;
                                }
                        }
                    }
                }
            }
        }
        private void HashDefineLookupValue(ref SPItem itm, string rowCell, DIMapping map, SPField fld, Dictionary<string, Hashtable> HashDictionary, Dictionary<string, SPList> LookupRelations)
        {
            var field = (SPFieldLookup)fld;
            SPList list = LookupRelations[map.Value];
            string Lookuplistname = list.Title + "LIST";
            Hashtable table = HashDictionary[Lookuplistname];

            //LookupRelations
            if (field.AllowMultipleValues)
            {

                var values = rowCell.ToLower().Split(multivalueDelimiter.ToArray(), StringSplitOptions.RemoveEmptyEntries);
                var result = new SPFieldLookupValueCollection();
                foreach (string val in values)
                {
                    bool isExists = false;

                    List<int> ItemID = (List<int>)table[val];
                    SPListItem Item = list.GetItemById(ItemID.FirstOrDefault());
                    result.Add(new SPFieldLookupValue(Item.ID, Item[field.LookupField].ToString()));
                    isExists = true;

                    if (!isExists)
                    {
                        if (!string.IsNullOrEmpty(val))
                            err += val + ": " + val + ErrNotFound;
                        else
                            err += val + ErrNotFound;
                    }
                }

                itm[map.Name] = result;
            }
            else
            {

                SPFieldLookupValue val = null;
                List<int> ItemID = (List<int>)table[rowCell.ToLower().Trim()];
                if (ItemID != null)
                {
                    SPListItem Item = list.GetItemById(ItemID.FirstOrDefault());
                    val = new SPFieldLookupValue(Item.ID, Item[field.LookupField].ToString());
                };

                if (val == null)
                {
                    if (!string.IsNullOrEmpty(rowCell))
                        err += map.Value + ": " + rowCell + ErrNotFound;
                    else
                        err += map.Value + ErrNotFound;
                }
                else
                    itm[map.Name] = val;

            };

        }
            
      
        private void FillNonStandardFields(List<SPListItem> items, string rowCell, DIMapping map, Dictionary<string, Hashtable> HashDictionary, Dictionary<string, SPList> LookupRelations, string dateFormat)
        {
            foreach (var itm in items)
            {
                
           
            var fld = nonStandardListFields.Where(sf => (sf.InternalName ?? "") == (map.Name ?? "")).FirstOrDefault();
            if (fld != null)
            {
                if (!string.IsNullOrEmpty(rowCell))
                {
                    switch (fld.Type)
                    {
                        case SPFieldType.Text:
                        case SPFieldType.Note:
                            {
                                itm[map.Name] = rowCell;
                                break;
                            }

                        case SPFieldType.Number:
                            {
                                try
                                {
                                    itm[map.Name] = Convert.ToDecimal(rowCell);
                                }
                                catch (Exception ex)
                                {
                                    err += map.Value + ErrWrongType;
                                }

                                break;
                            }

                        case SPFieldType.Boolean:
                            {
                                itm[map.Name] = CheckBoolean(rowCell);
                                break;
                            }

                        case SPFieldType.DateTime:
                            {
                                try
                                {
                                    var parsedDate = ParseDate(rowCell, dateFormat);
                                    if (parsedDate != default(DateTime))
                                        itm[map.Name] = parsedDate;
                                }
                                catch (Exception ex)
                                {
                                    err += map.Value + ErrWrongType;
                                }

                                break;
                            }

                        case SPFieldType.Lookup:
                            {
                                SPItem argitm = itm;
                                HashDefineLookupValue(ref argitm, rowCell, map, fld, HashDictionary, LookupRelations);
                                break;
                            }

                        case SPFieldType.Choice:
                            {
                                System.Collections.Specialized.StringCollection choices;
                                try
                                {
                                    choices = (System.Collections.Specialized.StringCollection)Relations[map.Name];
                                }
                                catch (Exception ex)
                                {
                                    err += "choices for " + map.Value + ErrNotFound;
                                    return;
                                }
                                string choice = string.Empty;
                                foreach (var ch in choices)
                                {
                                    if ((ch.Trim().ToLower() ?? "") == (rowCell.ToLower() ?? ""))
                                    {
                                        choice = ch;
                                        break;
                                    }
                                }
                                if (!string.IsNullOrEmpty(choice))
                                    itm[map.Name] = choice;
                                else
                                    err += map.Value + ErrNotFound;
                                break;
                            }

                        case SPFieldType.Invalid:
                            {
                                if (fld.FieldValueType == typeof(SPFieldLookupValue) | fld.FieldValueType == typeof(SPFieldLookupValueCollection))
                                {
                                    SPItem argitm1 = itm;
                                        HashDefineLookupValue(ref argitm1, rowCell, map, fld, HashDictionary, LookupRelations);
                                   // DefineLookupValue(ref argitm1, rowCell, map, fld);
                                }

                                break;
                            }
                    }
                }
                else
                    switch (fld.Type)
                    {
                        case SPFieldType.Text:
                        case SPFieldType.Note:
                        case SPFieldType.Choice:
                            {
                                itm[map.Name] = string.Empty;
                                break;
                            }

                        case SPFieldType.Number:
                            {
                                itm[map.Name] = 0;
                                break;
                            }

                        default:
                            {
                                itm[map.Name] = null;
                                break;
                            }
                    }
            }
           }
        }
        public void ProcessDataImport(IProgress<int> progress, SPList lst, List<DIMapping> mapping, DataTable dataTable, string dateFormat, SPWeb web, Dictionary<string, SPList> LookupRelations,string TimeFormat, Dictionary<string, Hashtable> HashDictionary =null, Dictionary<string, int> SelectedKeysRows =null )
        {
            Errors.Clear();
            NumberOfItems = 0;
            _employeesListGUID = GetEmpoloyeesList(web.Url).ID;
            df.ShortDatePattern = dateFormat;
            FillNonStandardFieldsRelations(mapping, web, lst);
            var fields = mapping.Where(m => standardFieldsCollection.All(sfc => (sfc ?? "") != (m.Name ?? ""))).Select(m => m.Name).ToArray();

            for (int index = 1, loopTo = dataTable.Rows.Count - 1; index <= loopTo; index++)
            {
                progress?.Report(index);
                List<SPListItem> itm = new List<SPListItem>();
                try
                {
                    err = string.Empty;
                    if (HashDictionary != null && SelectedKeysRows != null)
                    {
                        DataRow rowCell = dataTable.Rows[index];
                        List<SPListItem>  UpdateItem = FindItemToUpdate(HashDictionary, SelectedKeysRows, rowCell, lst);
                        if (UpdateItem == null)
                        {
                            //no item found, ltesc created a new one
                            SPListItem NewItemToCreate = lst.AddItem();
                            itm.Add(NewItemToCreate);

                        }
                        else {
                            itm = itm.Concat(UpdateItem).ToList();
                        };
                    }
                    else
                    {
                        SPListItem NewItemToCreate = lst.AddItem();
                        itm.Add(NewItemToCreate);
                    }

                    for (int i = 0, loopTo1 = dataTable.Columns.Count - 1; i <= loopTo1; i++)
                    {
                        string rowCell = dataTable.Rows[index][i].ToString().Trim();
                        string colNameByIndex = dataTable.Rows[0][i].ToString().Trim();
                        var map = mapping.Where(m => (m.Value ?? "") == (colNameByIndex ?? "")).FirstOrDefault();

                        if (map != null)
                            FillNonStandardFields(itm, rowCell, map, HashDictionary, LookupRelations, dateFormat);
                    }
                }
                catch (Exception ex)
                {
                    err += "Data fail " + ex.Message + ex.StackTrace;
                }
                if (!string.IsNullOrEmpty(err))
                    Errors.Add(index, err);
                else
                    UpdateSaveItems(dataTable, itm, index);
            }
        }

        private List<SPListItem> FindItemToUpdate(Dictionary<string, Hashtable> HashDictionary, Dictionary<string, int> SelectedKeysRows, DataRow rowCell, SPList DestanationList)
        {
            List<int> result = null;
            List<SPListItem> ItemsToUpdate = null;
            bool firstlap = true;

            foreach (KeyValuePair<string, int> KeyValue in SelectedKeysRows)
            {
                string FieldName = KeyValue.Key;
                int RowId = KeyValue.Value;

                string SearchValue = rowCell[RowId].ToString().Trim().ToLower();
                Hashtable table = HashDictionary[FieldName];
                List<int> tmp = null;
                if (firstlap == true)
                {   //fist lap
                    if (table.ContainsKey(SearchValue))
                    {
                        tmp = (List<int>)table[SearchValue];
                        result = tmp;
                    }

                }
                else
                {//not first lap
                    if (table.ContainsKey(SearchValue))
                    {
                        tmp = (List<int>)table[SearchValue];
                        result = result.Intersect(tmp).ToList();
                    }

                };
                if (result == null || result.Count < 0) return null;

                firstlap = false;
                
            }
            if (result != null)
            {

                foreach (int resultID in result)
                {
                    try
                    {
                        SPListItem SPItem = DestanationList.GetItemById(resultID);
                            if (ItemsToUpdate==null && SPItem!=null)
                            {
                                ItemsToUpdate = new List<SPListItem>();
                            };
                        ItemsToUpdate.Add(SPItem);
                    }
                    catch (Exception ex)
                    {
                        err += "Data fail " + ex.Message + ex.StackTrace;
                    };
                };
            };

            return ItemsToUpdate;
        }

        private void UpdateSaveItems(DataTable dataTable, List<SPListItem> itemsToUpdate, int index)
        {
            using (var scope = new DisabledItemEventsScope())
            {
                foreach (var item in itemsToUpdate)
                {
                    try
                    {
                        item.Update();
                        NumberOfItems = NumberOfItems + 1;
                    }
                    catch
                    {
                        err += ErrSave;
                    }
                }
            }
            if (!string.IsNullOrEmpty(err))
            {
                if (Errors.Keys.Contains(index))
                    Errors[index] += err;
                else
                    Errors.Add(index, err);
            }
        }
        public string GetErrors()
        {
            var builder = new StringBuilder();
            foreach (var pair in Errors)
                builder.Append(pair.Key).Append(" : ").Append(pair.Value).Append(Environment.NewLine);
            string result = builder.ToString();
            return result;
        }
        public int GetNumberOfItems()
        {
            return NumberOfItems;
        }
    }

    [Serializable()]
    public class DIMapping
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
}
