using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace AutomateFormProcess.Classes
{
    public static class Utility
    {
        // Methods
        public static List<string> AllIndexesOf(string oldStr, string value, out string str, List<SPFieldValue> item)
        {
            List<string> list = new List<string>();
            str = "";
            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentException("the string to find may not be empty", "value");
            }
            List<int> list2 = new List<int>();
            int num = 0;
            int startIndex = 0;
            while (true)
            {
                startIndex = oldStr.IndexOf(value, startIndex);
                if (startIndex == -1)
                {
                    str = oldStr;
                    return list;
                }
                int index = oldStr.IndexOf("}}", (int)(startIndex + 1));
                string filter = oldStr.Substring(startIndex + 2, (index - startIndex) - 2);
                SPFieldValue value2 = item.FirstOrDefault<SPFieldValue>(a => a.InternalName == filter);
                oldStr = oldStr.Replace(oldStr.Substring(startIndex, (index - startIndex) + 2), value2.value);
                list.Add(filter);
                num++;
                startIndex += value.Length;
            }
        }

      

       

        public static string DeleteItemFromList(SPList list, int itemId)
        {
            string message = "ok";
            try
            {
                Guid siteID = list.ParentWeb.Site.ID;
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    using (SPSite site = new SPSite(siteID))
                    {
                        using (SPWeb web = site.OpenWeb())
                        {
                            web.AllowUnsafeUpdates = true;
                            SPList list1 = web.Lists[list.ID];
                            list1.Items.DeleteItemById(itemId);
                            web.AllowUnsafeUpdates = false;
                        }
                    }
                });
            }
            catch (Exception exception)
            {
                message = exception.Message;
            }
            return message;
        }

      
        


        private static SPListItem setItemFields(List<SPFieldValue> fields, SPListItem item)
        {
            foreach (SPFieldValue value2 in fields)
            {
                string[] strArray;
                SPFieldLookupValueCollection values;
                int num;
                switch (value2.Type)
                {
                    case "Text":
                        {
                            item[value2.InternalName] = value2.value;
                            continue;
                        }
                    case "Note":
                        {
                            item[value2.InternalName] = value2.value;
                            continue;
                        }
                    case "Number":
                        {
                            item[value2.InternalName] = decimal.Parse(value2.value);
                            continue;
                        }
                    case "DateTime":
                        {
                            item[value2.InternalName] = Convert.ToDateTime(value2.value);
                            continue;
                        }
                    case "Lookup":
                        {
                            item[value2.InternalName] = new SPFieldLookupValue(int.Parse(value2.value), "");
                            continue;
                        }
                    case "LookupMulti":
                        strArray = value2.value.Split(new char[] { ',' });
                        values = new SPFieldLookupValueCollection();
                        num = 0;
                        goto Label_0216;

                    case "RelatedCustomLookupQuery":
                        {
                            item[value2.InternalName] = new SPFieldLookupValue(int.Parse(value2.value), "");
                            continue;
                        }
                    case "File":
                        {
                            item[value2.InternalName] = new SPFieldLookupValue(int.Parse(value2.value), "");
                            continue;
                        }
                    case "CustomComputedField":
                        {
                            continue;
                        }
                    case "Choice":
                        {
                            item[value2.InternalName] = value2.value;
                            continue;
                        }
                    case "MultiChoice":
                        {
                            string[] strArray2 = value2.value.Split(new char[] { ',' });
                            SPFieldMultiChoiceValue value3 = new SPFieldMultiChoiceValue();
                            num = 0;
                            while (num < strArray2.Length)
                            {
                                value3.Add(strArray2[num]);
                                num++;
                            }
                            item[value2.InternalName] = value3;
                            continue;
                        }
                    case "Boolean":
                        {
                            item[value2.InternalName] = value2.value;
                            continue;
                        }
                    case "User":
                        {
                            if (value2.value == null)
                            {
                                goto Label_034A;
                            }
                            item[value2.InternalName] = new SPFieldUserValue(SPContext.Current.Web, int.Parse(value2.value), "");
                            continue;
                        }
                    default:
                        {
                            continue;
                        }
                }
            Label_01F7:
                values.Add(new SPFieldLookupValue(int.Parse(strArray[num]), ""));
                num++;
            Label_0216:
                if (num < strArray.Length)
                {
                    goto Label_01F7;
                }
                item[value2.InternalName] = values;
                continue;
            Label_034A:
                item[value2.InternalName] = null;
            }
            return item;
        }

    

        public static string UpdateFiles(SPWeb web,string folderName, List<SPFieldValue> fields, List<Attachment> addFiles, List<Attachment> deleteFiles)
        {
            string message = "ok";
            try
            {
                Guid siteId = web.Site.ID;
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    using (SPSite site = new SPSite(siteId))
                    {
                        using (SPWeb Web = site.OpenWeb())
                        {
                            SPList list;
                            foreach (Attachment attachment in deleteFiles)
                            {
                                list = Web.Lists[new Guid(attachment.LookupList)];
                                Web.AllowUnsafeUpdates = true;
                                SPFolder folder = Web.GetFolder(list.RootFolder + folderName);
                                if (!folder.Exists)
                                {
                                    list.RootFolder.Files[attachment.FileName].Delete();
                                }
                                else
                                {
                                    folder.Files[attachment.FileName].Delete();
                                }
                                list.RootFolder.Update();
                                Web.AllowUnsafeUpdates = false;
                            }
                            foreach (Attachment attachment in addFiles)
                            {
                                list = web.Lists[new Guid(attachment.LookupList)];
                                byte[] buffer = Convert.FromBase64String(attachment.Content);
                                if (!string.IsNullOrEmpty(attachment.FileName))
                                {
                                    web.AllowUnsafeUpdates = true;
                                    SPFile file = null;
                                    if (folderName=="")
                                    {
                                        file = list.RootFolder.Files.Add(attachment.FileName, buffer, false);
                                    }
                                    else
                                    {
                                        SPFolder folderWithName = web.GetFolder(list.RootFolder + "/" + folderName);
                                        if (!folderWithName.Exists)
                                        {
                                            SPListItem item = list.Folders.Add(list.RootFolder.ServerRelativeUrl, SPFileSystemObjectType.Folder, folderName);
                                            item.Update();
                                            file = item.Folder.Files.Add(attachment.FileName, buffer, false);
                                        }
                                        else
                                        {
                                            file = folderWithName.Files.Add(attachment.FileName, buffer, false);
                                        }
                                    }
                                    file.Item["Title"] = attachment.Title;
                                    file.Item.Update();
                                    web.AllowUnsafeUpdates = false;
                                    SPFieldValue fileFieldValue = new SPFieldValue
                                    {
                                        InternalName = attachment.InternalName,
                                        Type = "File",
                                        LookupList = attachment.LookupList,
                                        value = file.Item.ID.ToString()
                                    };
                                    fields.Add(fileFieldValue);
                                }
                            }
                        }
                    }
                });
            }
            catch (Exception exception)
            {
                message = exception.Message;
            }
            return message;
        }
    }

 
}
