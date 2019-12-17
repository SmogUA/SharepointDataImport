using System.Data;
using System.Linq;
using System.Collections.Generic;
using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace DataImport
{
    public class SharePoint
    {
        private readonly SPFarm farm = SPFarm.Local;
        public List<SPWeb> GetAllWebs()
        {
            var AllWebs = new List<SPWeb>();
            SPSecurity.RunWithElevatedPrivileges(() =>
            {
                var services = farm.Services;
                foreach (SPService curService in services)
                {
                    if (curService is SPWebService)
                    {
                        var webService = (SPWebService)curService;
                        if (curService.TypeName.Equals("Microsoft SharePoint Foundation Web Application"))
                        {
                            webService = (SPWebService)curService;
                            var webApplications = webService.WebApplications;
                            foreach (SPWebApplication webApplication in webApplications)
                            {
                                try
                                {
                                    if (webApplication != null)
                                    {
                                        foreach (SPSite site in webApplication.Sites)
                                        {
                                            foreach (SPWeb web in site.AllWebs)
                                            {
                                                if (web.Url.Contains("/es"))
                                                    AllWebs.Add(web);
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                }
                            }
                        }
                    }
                }
            });
            return AllWebs;
        }
        public List<SPList> GetAllLists(string url)
        {
            SPWeb web;
            using (var siteCollection = new SPSite(url))
            {
                web = siteCollection.OpenWeb();
            }
            var AllLists = new List<SPList>();
            SPSecurity.RunWithElevatedPrivileges(() =>
            {
                foreach (SPList list in web.Lists)
                    AllLists.Add(list);
            });
            return AllLists;
        }
        public List<SPField> GetListFields(SPList list)
        {
            var fields = new List<SPField>();
            if (list != null)
            {
                foreach (SPField fld in list.Fields)
                {
                    if ((fld.ReadOnlyField || !fld.CanBeDeleted) && !((fld.InternalName ?? "") == "Title") && !((fld.InternalName ?? "") == "Employee"))
                        continue;
                    fields.Add(fld);
                }
            }
            return fields;
        }
        public SPList GetListByDisplayName(string url, string displayListName)
        {
            SPWeb web;
            using (var siteCollection = new SPSite(url))
            {
                web = siteCollection.OpenWeb();
            }
            SPList list = null;         
            list = web.Lists[displayListName];
            return list;
        }
        public SPList GetListByInternalName(string url, string internalListName)
        {
            SPWeb _web;
            using (var siteCollection = new SPSite(url))
            {
                _web = siteCollection.OpenWeb();
            }
            string path = _web.ServerRelativeUrl;
            if (!path.EndsWith("/", StringComparison.OrdinalIgnoreCase))
                path += "/";
            path += "Lists/" + internalListName;
            var list = _web.GetList(path);
            return list;
        }
        public SPList GetLibraryByInternalName(string url, string internalLibraryName)
        {
            SPWeb _web;
            using (var siteCollection = new SPSite(url))
            {
                _web = siteCollection.OpenWeb();
            }
            string path = _web.ServerRelativeUrl;
            if (!path.EndsWith("/", StringComparison.OrdinalIgnoreCase))
                path += "/";
            path += internalLibraryName;
            var list = _web.GetList(path);
            return list;
        }
        public SPWeb GetWeb(string url)
        {
            SPWeb web;
            using (var siteCollection = new SPSite(url))
            {
                web = siteCollection.OpenWeb();
            }
            return web;
        }
    }
}
