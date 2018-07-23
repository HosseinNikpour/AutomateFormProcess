using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.DirectoryServices.AccountManagement;
using Newtonsoft.Json;
using System.Web.Services;
using System.Collections.Generic;
using AutomateFormProcess.Classes;
using System.Linq;
using System.Xml;

namespace AutomateFormProcess.Layouts.AutomateFormProcess
{
    public partial class Services : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }
        [WebMethod]
        public static object getSpData(string spName)
        {
            SPQuery query;
            SPWeb web = SPContext.Current.Web;
            string[] strArray = spName.Split(',');
            object[] parameterValues = new object[strArray.Length - 1];
            string sPName = strArray[0].ToString();

            SPUser currentUser = web.CurrentUser;
            string str = "";
            //var domainName = "jnasr";
            //var queryUser = "nasr\\SPS_Farm_Prd";
            //var queryUserPassword = "rnk@Qk7fqH";
            var domainName = "nasr2";
            var queryUser = "nasr2\\spadmin";
            var queryUserPassword = "Nsr!dm$n!Sp";
            var principalContext = new PrincipalContext(ContextType.Domain, domainName, queryUser, queryUserPassword);
            GroupPrincipal managerPrincipal = GroupPrincipal.FindByIdentity(principalContext, "pmis_managers");
            GroupPrincipal directorPrincipal = GroupPrincipal.FindByIdentity(principalContext, "pmis_directors");
            GroupPrincipal contractorPrincipal = GroupPrincipal.FindByIdentity(principalContext, "pmis_contractors");
            GroupPrincipal engineerPrincipal = GroupPrincipal.FindByIdentity(principalContext, "pmis_engineers");
            //GroupPrincipal managerPrincipal = GroupPrincipal.FindByIdentity(principalContext, "Epm-managers");
            //  GroupPrincipal directorPrincipal = GroupPrincipal.FindByIdentity(principalContext, "Epm-Directors_pmis");
            // GroupPrincipal contractorPrincipal = GroupPrincipal.FindByIdentity(principalContext, "Epm-contractors");
            // GroupPrincipal engineerPrincipal = GroupPrincipal.FindByIdentity(principalContext, "Epm-engineers");
            UserPrincipal user = UserPrincipal.FindByIdentity(principalContext, currentUser.LoginName);
            string userdomain = currentUser.LoginName != "SHAREPOINT\\system" ? currentUser.LoginName.Split('|')[1].Split('\\')[0] : "nasr";

            if (strArray[1].ToString().ToUpper() == "NULL")
            {
                if (userdomain != "nasr" && (user.IsMemberOf(managerPrincipal) || user.IsMemberOf(directorPrincipal)))
                {
                    query = new SPQuery();
                    SPListItem item = SPContext.Current.Web.GetList("/Lists/Areas").Items[0];
                    parameterValues[0] = item.ID;
                }
                else
                {
                    parameterValues[0] = System.DBNull.Value;
                }
            }
            else
            {
                parameterValues[0] = strArray[1].ToString();
            }
            if (strArray[2].ToString().ToUpper() == "NULL")
            {
                if (userdomain != "nasr" && (user.IsMemberOf(contractorPrincipal) || user.IsMemberOf(engineerPrincipal)))
                {

                    query = new SPQuery();
                    SPListItem item2 = SPContext.Current.Web.GetList("/Lists/Contracts").Items[0];
                    parameterValues[1] = item2.ID;
                }
                else
                {
                    parameterValues[1] = System.DBNull.Value;
                }
            }
            else
            {
                parameterValues[1] = strArray[2].ToString();
            }

            if (strArray.Length > 3)
            {
                for (int i = 3; i < strArray.Length; i++)
                {
                    if (strArray[i].ToString().ToUpper() == "NULL")
                    {
                        parameterValues[i - 1] = System.DBNull.Value;
                    }
                    else
                    {
                        int result = 0;
                        if (int.TryParse(strArray[i].ToString(), out result))
                        {
                            parameterValues[i - 1] = result;
                        }
                        else
                        {
                            parameterValues[i - 1] = strArray[i].ToString();
                        }
                    }
                }
            }
            DataAccessBase base2 = new DataAccessBase();
            return JsonConvert.SerializeObject(base2.ReaderSp(sPName, parameterValues));
        }

        [WebMethod]
        public  static object GetDisplayFields(string listId, string formType, string contentTypeId)
        {
            SPWeb web = SPContext.Current.Web;

            SPList currentList = web.Lists[new Guid(listId)];
            string listUrl = currentList.RootFolder.ServerRelativeUrl.ToString();
            string listInternalName = listUrl.Substring(listUrl.LastIndexOf("/") + 1);

            return new { fields = getFieldsFromList(currentList, contentTypeId) };
        }


        private static List<SPFieldGeneral> getFieldsFromList(SPList list, string contentTypeId)
        {
            SPContentType type = list.ContentTypes[contentTypeId];
            List<SPFieldGeneral> listFields = new List<SPFieldGeneral>();
            SPFieldCollection fields = (contentTypeId != "") ? list.ContentTypes[new SPContentTypeId(contentTypeId)].Fields : list.ContentTypes[0].Fields;
            foreach (SPField field in fields)
            {
                if ((field.Hidden || (field.FromBaseType && (field.InternalName != "Title"))) || ((field.InternalName == "Status") || (field.InternalName == "CurrentUser")))
                {
                    continue;
                }
                SPFieldGeneral general = new SPFieldGeneral();
                string str = field.Type.ToString();
                general.Guid = field.Id;
                general.InternalName = field.InternalName;
                XmlDocument document = new XmlDocument();
                document.LoadXml(field.SchemaXml);
                XmlElement documentElement = document.DocumentElement;
                if (documentElement.HasAttribute("DisplayName"))
                {
                    general.Title = documentElement.GetAttribute("DisplayName");
                }
                else
                {
                    general.Title = field.Title;
                }
                general.Title = field.GetProperty("DisplayName");
                general.Disable = field.GetProperty("Disable");
                general.NotShow = field.GetProperty("NotShow");
                string[] notShowArr =general.NotShow!=null? general.NotShow.Split(','):new string[]{};
                general.DefaultValue = field.DefaultValue;
                general.IsRequire = field.Required;
                general.Type = field.TypeAsString;
                general.Description = field.Description;
                switch (field.TypeAsString)
                {
                    case "Text":
                        general.MaxLength = ((SPFieldText)field).MaxLength;
                        break;

                    case "Number":
                        general.MaxValue = ((SPFieldNumber)field).MaximumValue;
                        general.MinValue = ((SPFieldNumber)field).MinimumValue;
                        general.ShowAsPercentage = ((SPFieldNumber)field).ShowAsPercentage;
                        break;

                    case "Lookup":
                        general.LookupList = ((SPFieldLookup)field).LookupList.Replace("{", "").Replace("}", "");
                        general.LookupTitleField = "Title";
                        general.LookupValueField = "ID";
                        general.AllowMultipleValue = ((SPFieldLookup)field).AllowMultipleValues;
                        break;

                    case "LookupMulti":
                        general.LookupList = ((SPFieldLookup)field).LookupList.Replace("{", "").Replace("}", "");
                        general.LookupTitleField = "Title";
                        general.LookupValueField = "ID";
                        general.AllowMultipleValue = ((SPFieldLookup)field).AllowMultipleValues;
                        break;

                    case "RelatedCustomLookupQuery":
                        general.LookupList = field.GetCustomProperty("ListNameLookup").ToString().Replace("{", "").Replace("}", "");
                        general.LookupTitleField = field.GetCustomProperty("FieldTitleLookup").ToString();
                        general.LookupValueField = field.GetCustomProperty("FieldValueLookup").ToString();
                        general.AllowMultipleValue = ((SPFieldLookup)field).AllowMultipleValues;
                        general.RelatedFields = field.GetCustomProperty("RelatedFields").ToString().Split(new char[] { '|' });
                        general.Query = field.GetCustomProperty("QueryLookup").ToString();
                        if ((field.GetCustomProperty("IsFile") != null) ? bool.Parse(field.GetCustomProperty("IsFile").ToString()) : false)
                        {
                            general.Type = "File";
                            general.TypeFile = field.GetCustomProperty("TypeFile").ToString();
                            general.VolumeFile = field.GetCustomProperty("VolumeFile").ToString();
                        }
                        break;

                    case "MasterDetail":
                        general.LookupList = field.GetCustomProperty("ListNameLookup").ToString();
                        general.RelatedFields = field.GetCustomProperty("RelatedFields").ToString().Split(new char[] { '|' });
                        general.MasterLookupName = field.GetCustomProperty("MasterFieldNameLookup").ToString();
                        break;

                    case "CustomComputedField":
                        general.LookupList = field.GetCustomProperty("ListNameQuery").ToString();
                        general.LookupTitleField = field.GetCustomProperty("FieldNameQuery").ToString();
                        general.Query = field.GetCustomProperty("TextQuery").ToString();
                        general.AggregationFunction = field.GetCustomProperty("AggregatorFunction").ToString();
                        break;

                    case "Choice":
                        {
                            SPFieldChoice choice = (SPFieldChoice)field;
                            general.options = new List<string>();
                            foreach (string str2 in choice.Choices)
                            {
                                general.options.Add(str2);
                            }
                            general.DefaultValue = ((SPFieldChoice)field).DefaultValue;
                            general.AllowMultipleValue = ((SPFieldChoice)field).ListItemMenu;
                            break;
                        }
                    case "MultiChoice":
                        {
                            SPFieldMultiChoice choice2 = (SPFieldMultiChoice)field;
                            general.options = new List<string>();
                            foreach (string str2 in choice2.Choices)
                            {
                                general.options.Add(str2);
                            }
                            general.AllowMultipleValue = ((SPFieldMultiChoice)field).ListItemMenu;
                            break;
                        }
                }
                if(!notShowArr.Contains("New"))
                    listFields.Add(general);
            }
            SPField fieldByInternalName = list.Fields.GetFieldByInternalName("ID");
            SPFieldGeneral item = new SPFieldGeneral
            {
                Guid = fieldByInternalName.Id,
                InternalName = fieldByInternalName.InternalName,
                Title = fieldByInternalName.Title,
                DefaultValue = fieldByInternalName.DefaultValue,
                IsRequire = fieldByInternalName.Required,
                Type = fieldByInternalName.TypeAsString,
                Description = fieldByInternalName.Description
            };
            listFields.Add(item);
            return listFields;
        }

       


        //save
        [WebMethod]
        public static string SaveFieldItems(string guid, List<SPFieldValue> fields, List<SPListItemDelete> deletedItems, List<Attachment> addFiles, List<Attachment> deleteFiles)
        {
            List<ErrorMessage> list = new List<ErrorMessage>();
            string s = "ok";
           
            List<SPItemSave> saveItems = new List<SPItemSave>();
            SPWeb web = SPContext.Current.Web;
            SPList masterList = web.Lists[new Guid(guid)];
            string url = masterList.RootFolder.Url;
            string listname = url.Substring(url.LastIndexOf('/') + 1);
            foreach (SPListItemDelete delete in deletedItems)
            {
                SPList delList = web.Lists[new Guid(delete.ListId)];
                s = Utility.DeleteItemFromList(delList, delete.ItemId);
                if (s != "ok")
                {
                    return s;
                }
            }
            SPFieldValue idValue = fields.FirstOrDefault<SPFieldValue>(a => a.InternalName == "ID");
            SPListItem item = (idValue.value != "0") ? masterList.GetItemById(int.Parse(idValue.value)) : null;
          
           
          
         
         
            if (s == "ok")
            {
                List<SPFieldValue> masterDetailFields = fields.Where(a => a.Type == "MasterDetail").ToList();
               // SPListItemCollection items = Utility.FindValidation(web, list4.ID);




                int result = 0;
                s = Utility.UpdateFiles(web,"", fields, addFiles, deleteFiles);
                if (s != "ok")
                {
                    return s;
                }
                SPListItem itemSave = (int.Parse(idValue.value) > 0) ? masterList.GetItemById(int.Parse(idValue.value)) : masterList.AddItem(masterList.RootFolder.ServerRelativeUrl, SPFileSystemObjectType.File);
                SPListItem itemSP= setItemFields(fields, itemSave);
                SPItemSave save = new SPItemSave
                {
                    ListId = guid,
                    Item = itemSP,
                    Folder = web.GetFolder( masterList.RootFolder.ServerRelativeUrl)
                };
                s = SaveItem(web, save.ListId, save.Item, save.Folder,"");
                int.TryParse(s, out result);
                if (result > 0)
                {
                    s = "ok";
                    using (List<SPFieldValue>.Enumerator masterDetaileEnum = masterDetailFields.GetEnumerator())
                    {
                        while (masterDetaileEnum.MoveNext())
                        {
                            SPFieldValue detailField = masterDetaileEnum.Current;
                            SPFieldValue lk = new SPFieldValue
                            {
                                InternalName = detailField.value,
                                Type = "Lookup",
                                value = result.ToString()
                            };
                            Guid siteID = web.Site.ID;
                            SPSecurity.RunWithElevatedPrivileges(delegate
                            {
                                using (SPSite site = new SPSite(siteID))
                                {
                                    using (SPWeb web1 = site.OpenWeb())
                                    {
                                        SPList detailList = web1.Lists[new Guid(detailField.LookupList)];
                                        //SPListItemCollection detailItems = Utility.FindValidation(web, detailList.ID);
                                        SPFolder folder = web1.GetFolder(detailList.RootFolder.ServerRelativeUrl);
                                        foreach (List<SPFieldValue> details in detailField.rows)
                                        {
                                            details.Add(lk);
                                            SPFieldValue detailIdValue = details.FirstOrDefault<SPFieldValue>(a => a.InternalName == "ID");
                                            SPListItem detailItem = setItemFields(details, (int.Parse(detailIdValue.value) > 0) ? detailList.GetItemById(int.Parse(detailIdValue.value)) : detailList.AddItem(folder.Url, SPFileSystemObjectType.File));
                                            SPItemSave detailsave = new SPItemSave
                                            {
                                                ListId = detailField.LookupList,
                                                Item = detailItem,
                                                Folder = folder
                                            };
                                            saveItems.Add(detailsave);
                                        }
                                    }
                                }
                            });
                        }
                    }
                }
                string msgResult = "";
                int num10 = 0;
                foreach (SPItemSave saveItem in saveItems)
                {
                    msgResult = SaveItem(web, saveItem.ListId, saveItem.Item, saveItem.Folder,"");
                }
                int.TryParse(msgResult, out num10);
               
            }
            return s;
        }

        private static string SaveItem(SPWeb web, string guid, SPListItem item, SPFolder spFolder,string folderTitle)
        {
            
            string str = "";
            Guid siteID = web.Site.ID;

            SPSecurity.RunWithElevatedPrivileges(delegate
            {

                using (SPSite site = new SPSite(siteID))
                {
                    using (SPWeb Web = site.OpenWeb())
                    {
                        SPList list = Web.Lists[new Guid(guid)];
                        if (!spFolder.Exists)
                        {
                            SPListItem folderItem = list.Items.Add(list.RootFolder.ServerRelativeUrl, SPFileSystemObjectType.Folder);
                            folderItem["Title"] = folderTitle;
                            Web.AllowUnsafeUpdates = true;
                            folderItem.Update();

                            spFolder = web.GetFolder(list.RootFolder.Url + "/" + folderTitle);
                        }



                    }
                }
            });
                            


               
                        
               
                SPWeb webItem = item.Web;
                webItem.AllowUnsafeUpdates = true;
                item.Update();
                
               
                str = item.ID.ToString();
                webItem.AllowUnsafeUpdates = false;
          
          
            return str;
        }

        private static void SetFieldValue(SPListItem item, SPFieldValue fieldValue)
        {
            int num;
            switch (fieldValue.Type)
            {
                case "Text":
                    item[fieldValue.InternalName] = fieldValue.value;
                    break;

                case "Note":
                    item[fieldValue.InternalName] = fieldValue.value;
                    break;

                case "Number":
                    item[fieldValue.InternalName] = decimal.Parse(fieldValue.value);
                    break;

                case "DateTime":
                    item[fieldValue.InternalName] = Convert.ToDateTime(fieldValue.value);
                    break;

                case "Lookup":
                    item[fieldValue.InternalName] = new SPFieldLookupValue(int.Parse(fieldValue.value), "");
                    break;

                case "LookupMulti":
                    {
                        string[] strArray = fieldValue.value.Split(new char[] { ',' });
                        SPFieldLookupValueCollection values = new SPFieldLookupValueCollection();
                        num = 0;
                        while (num < strArray.Length)
                        {
                            values.Add(new SPFieldLookupValue(int.Parse(strArray[num]), ""));
                            num++;
                        }
                        item[fieldValue.InternalName] = values;
                        break;
                    }
                case "RelatedCustomLookupQuery":
                    item[fieldValue.InternalName] = new SPFieldLookupValue(int.Parse(fieldValue.value), "");
                    break;
                case "CustomComputedField":
                    item[fieldValue.InternalName] = new SPFieldLookupValue(int.Parse(fieldValue.value), "");
                    break;

                case "Choice":
                    item[fieldValue.InternalName] = fieldValue.value;
                    break;

                case "MultiChoice":
                    {
                        string[] strArray2 = fieldValue.value.Split(new char[] { ',' });
                        SPFieldMultiChoiceValue value2 = new SPFieldMultiChoiceValue();
                        for (num = 0; num < strArray2.Length; num++)
                        {
                            value2.Add(strArray2[num]);
                        }
                        item[fieldValue.InternalName] = value2;
                        break;
                    }
                case "Boolean":
                    item[fieldValue.InternalName] = fieldValue.value;
                    break;
            }
        }

        private static SPListItem setItemFields(List<SPFieldValue> fields, SPListItem item)
        {
            foreach (SPFieldValue fieldValue in fields)
            {
                string[] strArray;
                SPFieldLookupValueCollection values;
                int num;
                switch (fieldValue.Type)
                {
                    case "Text":
                        {
                            item[fieldValue.InternalName] = fieldValue.value;
                            continue;
                        }
                    case "Note":
                        {
                            item[fieldValue.InternalName] = fieldValue.value;
                            continue;
                        }
                    case "Number":
                        {
                            item[fieldValue.InternalName] = decimal.Parse(fieldValue.value);
                            continue;
                        }
                    case "DateTime":
                        {
                            item[fieldValue.InternalName] = (fieldValue.value != "") ? new DateTime?(Convert.ToDateTime(fieldValue.value)) : null;
                            continue;
                        }
                    case "Lookup":
                        {
                            item[fieldValue.InternalName] = ((fieldValue.value != "") && (fieldValue.value != "0")) ? new SPFieldLookupValue(int.Parse(fieldValue.value), "") : null;
                            continue;
                        }
                    case "LookupMulti":
                        {
                            strArray = fieldValue.value.Split(new char[] { ',' });
                            values = new SPFieldLookupValueCollection();
                            num = 0;


                            if (num < strArray.Length)
                            {
                                values.Add(new SPFieldLookupValue(int.Parse(strArray[num]), ""));
                                num++;
                            }
                            item[fieldValue.InternalName] = values;
                            continue;
                        }
                    case "RelatedCustomLookupQuery":
                        {

                            strArray = fieldValue.value.Split(',');
                            if (strArray.Count() > 0)
                            {
                                values = new SPFieldLookupValueCollection();
                                num = 0;
                                if (num < strArray.Length && strArray[num].Length>0)
                                {
                                    values.Add(new SPFieldLookupValue(int.Parse(strArray[num]), ""));
                                    num++;
                                }
                                item[fieldValue.InternalName] = values;
                            }
                            else
                                item[fieldValue.InternalName] = ((fieldValue.value != "") && (fieldValue.value != "0")) ? new SPFieldLookupValue(int.Parse(fieldValue.value), "") : null;
                            continue;
                        }
                    case "CustomComputedField":
                        {
                            item[fieldValue.InternalName] = fieldValue.value;
                            continue;
                        }
                    case "Choice":
                        {
                            item[fieldValue.InternalName] = fieldValue.value;
                            continue;
                        }
                    case "MultiChoice":
                        {
                            string[] strArray2 = fieldValue.value.Split(new char[] { ',' });
                            SPFieldMultiChoiceValue value3 = new SPFieldMultiChoiceValue();
                            num = 0;
                            while (num < strArray2.Length)
                            {
                                value3.Add(strArray2[num]);
                                num++;
                            }
                            item[fieldValue.InternalName] = value3;
                            continue;
                        }
                    case "Boolean":
                        {
                            item[fieldValue.InternalName] = fieldValue.value;
                            continue;
                        }
                    case "User":
                        {
                            item[fieldValue.InternalName] = ((fieldValue.value != "") && (fieldValue.value != "0")) ? new SPFieldUserValue(SPContext.Current.Web, int.Parse(fieldValue.value), "") : null;
                            continue;
                        }
                    case "File":
                        {
                            item[fieldValue.InternalName] = ((fieldValue.value != "") && (fieldValue.value != "0")) ? new SPFieldLookupValue(int.Parse(fieldValue.value), "") : null;
                            continue;
                        }
                    default:
                        {
                            continue;
                        }
                }
           
          
           
            }
            return item;
        }

    }
}
