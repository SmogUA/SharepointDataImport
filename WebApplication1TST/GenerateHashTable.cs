using DataImport;
using Microsoft.SharePoint;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

namespace WebApplication1TST
{
    static class GenerateHashTable
    {

        public static ConcurrentDictionary<string, Hashtable> HashTableForListField(SPList listIn, List<string> SelectedKeys, List<DIMapping> mapping, string WebURL, string TimeFormat)
        {

           ConcurrentDictionary<string, Hashtable> HashDictionary = new ConcurrentDictionary<string, Hashtable>();
           Dictionary<string, SPFieldType> FieldTypes= new Dictionary<string, SPFieldType>();
           string listTitle = listIn.Title;
          //Createting Key fields HashTables
            for (var i = 0; i < SelectedKeys.Count; i++)
            {
                HashDictionary.TryAdd(SelectedKeys[i], new Hashtable());
                DIMapping mp = mapping.Find(x => x.Value == SelectedKeys[i]);

                SPField field = listIn.Fields.GetField(mp.Name);
                SPFieldType fieldType = field.Type;
                FieldTypes.Add(SelectedKeys[i], fieldType);


            }
            //CreatefingLookup fields HashTables
            //1
            Parallel.ForEach(mapping, (mp) =>
            {
             SPSecurity.RunWithElevatedPrivileges(() =>
                {
                    SharePoint SP = new SharePoint();
                    SPWeb Web = SP.GetWeb(WebURL);
                    SPList list = SP.GetListByDisplayName(WebURL, listTitle);

                    SPField field = list.Fields.GetField(mp.Name);
                SPFieldType fieldType = field.Type;
                if (fieldType == SPFieldType.Lookup)
                {
                    SPList LookupList = null;
                    SPFieldLookup lookupField = (SPFieldLookup)list.Fields.GetField(mp.Name);

                    if (!String.IsNullOrEmpty(lookupField.LookupList) && !String.IsNullOrEmpty(lookupField.LookupField))
                    {
                        // Is this the primary or secondary field for a list relationship?
                        //string strRelationship = lookupField.IsRelationship ? "Primary" : "Secondary";


                        // Is this a secondary field in a list relationship?
                        if (lookupField.IsDependentLookup)
                        {
                            SPField primaryField = list.Fields[new Guid(lookupField.PrimaryFieldId)];

                        }
                        // Get the site where the target list is located.

                        // Get the name of the list where this field gets information.
                        LookupList = Web.Lists[new Guid(lookupField.LookupList)];
                        SPField targetField = LookupList.Fields.GetFieldByInternalName(lookupField.LookupField);
                        string TargetInternalName = targetField.InternalName;
                        string Lookuplistname = LookupList.Title + "LIST";

                        if (!HashDictionary.ContainsKey(Lookuplistname))
                        {
                            Hashtable HT = GetLookupHash(LookupList, TargetInternalName);
                            HashDictionary.TryAdd(Lookuplistname, HT);
                            
                        }
                    }
                };


                });
            });
            //2
            /* foreach (DIMapping mp in mapping)
             {    
              SPField field = list.Fields.GetField(mp.Name);
              SPFieldType fieldType = field.Type;

                 if (fieldType == SPFieldType.Lookup)
                 {
                     SPList LookupList = null;
                     SPFieldLookup lookupField = (SPFieldLookup)list.Fields.GetField(mp.Name);

                     if (!String.IsNullOrEmpty(lookupField.LookupList) && !String.IsNullOrEmpty(lookupField.LookupField))
                     {
                         // Is this the primary or secondary field for a list relationship?
                         string strRelationship = lookupField.IsRelationship ? "Primary" : "Secondary";


                         // Is this a secondary field in a list relationship?
                         if (lookupField.IsDependentLookup)
                         {
                             SPField primaryField = list.Fields[new Guid(lookupField.PrimaryFieldId)];

                         }
                         // Get the site where the target list is located.

                         // Get the name of the list where this field gets information.
                         LookupList = Web.Lists[new Guid(lookupField.LookupList)];
                         SPField targetField = LookupList.Fields.GetFieldByInternalName(lookupField.LookupField);
                         string TargetInternalName = targetField.InternalName;
                         string Lookuplistname = LookupList.Title + "LIST";

                         if (!HashDictionary.ContainsKey(Lookuplistname))
                         {
                             Hashtable HT = GetLookupHash(LookupList, TargetInternalName);
                             HashDictionary.TryAdd(Lookuplistname, HT);
                             LookupRelations.Add(mp.Value, LookupList);
                         }
                     }             
                 };
             }
             */


            foreach (SPListItem listItem in listIn.Items)
            {
                for (var i = 0; i < SelectedKeys.Count; i++)
                {
                    Hashtable table = HashDictionary[SelectedKeys[i]];
                    SPFieldType FieldType = FieldTypes[SelectedKeys[i]];
                    DIMapping mp = mapping.Find(x => x.Value == SelectedKeys[i]);
                    string Val = null;
                    switch (FieldType)
                    {
                        case SPFieldType.Lookup:
                            if (listItem[mp.Name] != null)
                            {
                                Val = listItem[mp.Name].ToString();
                                SPFieldLookupValue value = new SPFieldLookupValue(Val);
                                Val = value.LookupValue.ToLower();
                            }
                            break;
                        case SPFieldType.Text:
                            if (listItem[mp.Name] != null)
                            {
                                Val = listItem[mp.Name].ToString().ToLower();
                            }
                            break;
                        case SPFieldType.DateTime:
                            if(listItem[mp.Name] != null)
                            { 
                            string TMPdate = listItem[mp.Name].ToString().ToLower();
                            DateTime dt= DateTime.Parse(TMPdate);                            
                            Val = dt.ToString(TimeFormat);
                            }
                             break;
                        case SPFieldType.Boolean:
                            if (listItem[mp.Name] != null)
                            {
                                Val = listItem[mp.Name].ToString().ToLower();
                            }
                            break;
                        case SPFieldType.Calculated:
                            if (listItem[mp.Name] != null)
                            {
                                Val = listItem[mp.Name].ToString().ToLower();
                            }
                            break;
                        case SPFieldType.Choice:
                            if (listItem[mp.Name] != null)
                            {
                                Val = listItem[mp.Name].ToString().ToLower();
                            }
                            break;
                        default:
                            if (listItem[mp.Name] != null)
                            {
                                Val = listItem[mp.Name].ToString().ToLower();
                            }
                            break;

                    }

                    if (Val != null && table.ContainsKey(Val))
                    {
                       List<int> tmp = (List<int>)table[Val];
                       tmp.Add(listItem.ID);

                    }
                    else if (Val != null)
                    {
                       List<int> ItemID = new List<int>();
                        ItemID.Add(listItem.ID);
                        table.Add(Val, ItemID);
                    };

                }

            }
            return HashDictionary;
        }

        private static Hashtable GetLookupHash(SPList list, string TargetInternalName)
        {
            var result = new Hashtable();
            foreach (SPListItem listItem in list.Items)
            {
                string Val = listItem[TargetInternalName].ToString().ToLower();

                if (String.IsNullOrEmpty(Val))
                {
                    Val = "";
                };

                    if (result.ContainsKey(Val))
                    {
                        List<int> tmp = (List<int>)result[Val];
                        tmp.Add(listItem.ID);

                    }
                    else
                    {
                        List<int> ItemID = new List<int>();
                        ItemID.Add(listItem.ID);
                        result.Add(Val, ItemID);
                    };
                

            }

            return result;
        }
    }
}
