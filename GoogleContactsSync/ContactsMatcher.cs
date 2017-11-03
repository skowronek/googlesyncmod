using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.Contacts;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    internal static class ContactsMatcher
    {
        /// <summary>
        /// Time tolerance in seconds - used when comparing date modified.
        /// Less than 60 seconds doesn't make sense, as the lastSync is saved without seconds and if it is compared
        /// with the LastUpdate dates of Google and Outlook, in the worst case you compare e.g. 15:59 with 16:00 and 
        /// after truncating to minutes you compare 15:00 with 16:00
        /// Better take 120 seconds, because when resetting matches the time difference can be up to 2 minutes
        /// </summary>
        public static int TimeTolerance = 120;

        public delegate void NotificationHandler(string message);
        public static event NotificationHandler NotificationReceived;

        /// <summary>
        /// Matches outlook and google contact by a) google id b) properties.
        /// </summary>
        /// <param name="sync">Syncronizer instance</param>
        /// <param name="duplicatesFound">Exception returned, if duplicates have been found (null else)</param>
        /// <returns>Returns a list of m pairs (outlook contact + google contact) for all contact. Those that weren't matche will have it's peer set to null</returns>
        public static List<ContactMatch> MatchContacts(ContactsSynchronizer sync, out DuplicateDataException duplicatesFound)
        {
            Logger.Log("Matching Outlook and Google contacts...", EventType.Information);
            var result = new List<ContactMatch>();

            var duplicateGoogleMatches = string.Empty;
            var duplicateOutlookContacts = string.Empty;
            sync.GoogleContactDuplicates = new Collection<ContactMatch>();
            sync.OutlookContactDuplicates = new Collection<ContactMatch>();

            var skippedOutlookIds = new List<string>();

            //for each outlook contact try to get google contact id from user properties
            //if no m - try to m by properties
            //if no m - create a new m pair without google contact. 
            //foreach (Outlook._ContactItem olc in outlookContacts)
            var outlookContactsWithoutOutlookGoogleId = new Collection<OutlookContactInfo>();
            #region Match first all outlookContacts by sync id
            for (int i = 1; i <= sync.OutlookContacts.Count; i++)
            {
                Outlook.ContactItem olc = null;
                try
                {
                    olc = sync.OutlookContacts[i] as Outlook.ContactItem;
                    if (olc == null)
                    {
                        if (sync.OutlookContacts[i] is Outlook.DistListItem)
                        {
                            Logger.Log("Skipping distribution list", EventType.Debug);
                            sync.TotalCount--;
                        }
                        else
                        {
                            Logger.Log("Empty Outlook contact found. Skipping", EventType.Warning);
                            sync.SkippedCount++;
                            sync.SkippedCountNotMatches++;
                        }
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some contacts throw exceptions
                    Logger.Log("Accessing Outlook contact threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    continue;
                }

                try
                {
                    // sometimes contacts throw Exception when accessing their properties, so we give it a controlled try first.
                    try
                    {
                        string email1Address = olc.Email1Address;
                    }
                    catch (Exception ex)
                    {
                        string message = string.Format("Can't access contact details for outlook contact, got {0} - '{1}'. Skipping", ex.GetType().ToString(), ex.Message);
                        try
                        {
                            message = string.Format("{0} {1}.", message, olc.FileAs);
                            //remember skippedOutlookIds to later not delete them if found on Google side
                            skippedOutlookIds.Add(string.Copy(olc.EntryID));

                        }
                        catch
                        {
                            //e.g. if olc.FileAs also fails, ignore, because messge already set
                            //message = null;
                        }

                        //if (olc != null && message != null) // it's useless to say "we couldn't access some contacts properties
                        //{
                        Logger.Log(message, EventType.Warning);
                        //}
                        sync.SkippedCount++;
                        sync.SkippedCountNotMatches++;
                        continue;
                    }

                    if (!IsContactValid(olc))
                    {
                        Logger.Log(string.Format("Invalid outlook contact ({0}). Skipping", olc.FileAs), EventType.Warning);
                        skippedOutlookIds.Add(string.Copy(olc.EntryID));
                        sync.SkippedCount++;
                        sync.SkippedCountNotMatches++;
                        continue;
                    }

                    if (olc.Body != null && olc.Body.Length > 62000)
                    {
                        // notes field too large                    
                        Logger.Log(string.Format("Skipping outlook contact ({0}). Reduce the notes field to a maximum of 62.000 characters.", olc.FileAs), EventType.Warning);
                        skippedOutlookIds.Add(string.Copy(olc.EntryID));
                        sync.SkippedCount++;
                        sync.SkippedCountNotMatches++;
                        continue;
                    }

                    NotificationReceived?.Invoke(string.Format("Matching contact {0} of {1} by id: {2} ...", i, sync.OutlookContacts.Count, olc.FileAs));

                    // Create our own info object to go into collections/lists, so we can free the Outlook objects and not run out of resources / exceed policy limits.
                    var olci = new OutlookContactInfo(olc, sync);

                    //try to m this contact to one of google contacts
                    var userProperties = olc.UserProperties;
                    var idProp = userProperties[sync.OutlookPropertyNameId];
                    try
                    {
                        if (idProp != null)
                        {
                            var googleContactId = string.Copy((string)idProp.Value);
                            var foundContact = sync.GetGoogleContactById(googleContactId);
                            var match = new ContactMatch(olci, null);

                            //Check first, that this is not a duplicate 
                            //e.g. by copying an existing Outlook contact
                            //or by Outlook checked this as duplicate, but the user selected "Add new"
                            var duplicates = sync.OutlookContactByProperty(sync.OutlookPropertyNameId, googleContactId);
                            if (duplicates.Count > 1)
                            {
                                foreach (OutlookContactInfo duplicate in duplicates)
                                {
                                    if (!string.IsNullOrEmpty(googleContactId))
                                    {
                                        Logger.Log("Duplicate Outlook contact found, resetting match and trying to match again: " + duplicate.FileAs, EventType.Warning);
                                        var item = duplicate.GetOriginalItemFromOutlook();
                                        try
                                        {
                                            ContactPropertiesUtils.ResetOutlookGoogleContactId(sync, item);
                                            item.Save();
                                        }
                                        finally
                                        {
                                            if (item != null)
                                            {
                                                Marshal.ReleaseComObject(item);
                                                item = null;
                                            }
                                        }
                                    }
                                }

                                if (foundContact != null && !foundContact.Deleted)
                                {
                                    ContactPropertiesUtils.ResetGoogleOutlookContactId(sync.SyncProfile, foundContact);
                                }

                                outlookContactsWithoutOutlookGoogleId.Add(olci);
                            }
                            else
                            {

                                if (foundContact != null && !foundContact.Deleted)
                                {
                                    //we found a m by google id, that is not deleted yet
                                    match.AddGoogleContact(foundContact);
                                    result.Add(match);
                                    //Remove the contact from the list to not sync it twice
                                    sync.GoogleContacts.Remove(foundContact);
                                }
                                else
                                {
                                    outlookContactsWithoutOutlookGoogleId.Add(olci);
                                }
                            }
                        }
                        else
                            outlookContactsWithoutOutlookGoogleId.Add(olci);
                    }
                    finally
                    {
                        if (idProp != null)
                            Marshal.ReleaseComObject(idProp);
                        Marshal.ReleaseComObject(userProperties);
                    }
                }

                finally
                {
                    Marshal.ReleaseComObject(olc);
                    olc = null;
                }

            }
            #endregion
            #region Match the remaining contacts by properties

            for (int i = 0; i < outlookContactsWithoutOutlookGoogleId.Count; i++)
            {
                OutlookContactInfo olci = outlookContactsWithoutOutlookGoogleId[i];

                NotificationReceived?.Invoke(string.Format("Matching contact {0} of {1} by unique properties: {2} ...", i + 1, outlookContactsWithoutOutlookGoogleId.Count, olci.FileAs));

                //no m found by id => m by common properties
                //create a default m pair with just outlook contact.
                var match = new ContactMatch(olci, null);

                //foreach google contact try to m and create a m pair if found some m(es)
                for (int j = sync.GoogleContacts.Count - 1; j >= 0; j--)
                {
                    var entry = sync.GoogleContacts[j];
                    if (entry.Deleted)
                        continue;


                    // only m if there is either an email or telephone or else
                    // a matching google contact will be created at each sync
                    //1. try to m by FileAs
                    //1.1 try to m by FullName
                    //2. try to m by primary email
                    //3. try to m by mobile phone number, don't m by home or business bumbers, because several people may share the same home or business number
                    //4. try to math Company, if Google Title is null, i.e. the contact doesn't have a name and title, only a company
                    string entryTitleFirstLastAndSuffix = OutlookContactInfo.GetTitleFirstLastAndSuffix(entry);
                    if (!string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(entry.Title) && olci.FileAs.Equals(entry.Title.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||  //Replace twice to not replace a \r\n by \r\r\n. This is necessary because \r\n are saved as \n only to google
                        !string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(entry.Name.FullName) && olci.FileAs.Equals(entry.Name.FullName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
                        !string.IsNullOrEmpty(olci.FullName) && !string.IsNullOrEmpty(entry.Name.FullName) && olci.FullName.Equals(entry.Name.FullName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
                        !string.IsNullOrEmpty(olci.TitleFirstLastAndSuffix) && !string.IsNullOrEmpty(entryTitleFirstLastAndSuffix) && olci.TitleFirstLastAndSuffix.Equals(entryTitleFirstLastAndSuffix.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
                        !string.IsNullOrEmpty(olci.Email1Address) && entry.Emails.Count > 0 && olci.Email1Address.Equals(entry.Emails[0].Address, StringComparison.InvariantCultureIgnoreCase) ||
                        //!string.IsNullOrEmpty(olci.Email2Address) && FindEmail(olci.Email2Address, entry.Emails) != null ||
                        //!string.IsNullOrEmpty(olci.Email3Address) && FindEmail(olci.Email3Address, entry.Emails) != null ||
                        olci.MobileTelephoneNumber != null && FindPhone(olci.MobileTelephoneNumber, entry.Phonenumbers) != null ||
                        !string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(entry.Title) && entry.Organizations.Count > 0 && olci.FileAs.Equals(entry.Organizations[0].Name, StringComparison.InvariantCultureIgnoreCase)
                        )
                    {
                        match.AddGoogleContact(entry);
                        sync.GoogleContacts.Remove(entry);
                    }

                }

                #region find duplicates not needed now
                //if (m.GoogleContact == null && m.OutlookContact != null)
                //{//If GoogleContact, we have to expect a conflict because of Google insert of duplicates
                //    foreach (Contact entry in sync.GoogleContacts)
                //    {                        
                //        if (!string.IsNullOrEmpty(olc.FullName) && olc.FullName.Equals(entry.Title, StringComparison.InvariantCultureIgnoreCase) ||
                //         !string.IsNullOrEmpty(olc.FileAs) && olc.FileAs.Equals(entry.Title, StringComparison.InvariantCultureIgnoreCase) ||
                //         !string.IsNullOrEmpty(olc.Email1Address) && FindEmail(olc.Email1Address, entry.Emails) != null ||
                //         !string.IsNullOrEmpty(olc.Email2Address) && FindEmail(olc.Email1Address, entry.Emails) != null ||
                //         !string.IsNullOrEmpty(olc.Email3Address) && FindEmail(olc.Email1Address, entry.Emails) != null ||
                //         olc.MobileTelephoneNumber != null && FindPhone(olc.MobileTelephoneNumber, entry.Phonenumbers) != null
                //         )
                //    }
                //// check for each email 1,2 and 3 if a duplicate exists with same email, because Google doesn't like inserting new contacts with same email
                //Collection<Outlook.ContactItem> duplicates1 = new Collection<Outlook.ContactItem>();
                //Collection<Outlook.ContactItem> duplicates2 = new Collection<Outlook.ContactItem>();
                //Collection<Outlook.ContactItem> duplicates3 = new Collection<Outlook.ContactItem>();
                //if (!string.IsNullOrEmpty(olc.Email1Address))
                //    duplicates1 = sync.OutlookContactByEmail(olc.Email1Address);

                //if (!string.IsNullOrEmpty(olc.Email2Address))
                //    duplicates2 = sync.OutlookContactByEmail(olc.Email2Address);

                //if (!string.IsNullOrEmpty(olc.Email3Address))
                //    duplicates3 = sync.OutlookContactByEmail(olc.Email3Address);


                //if (duplicates1.Count > 1 || duplicates2.Count > 1 || duplicates3.Count > 1)
                //{
                //    if (string.IsNullOrEmpty(duplicatesEmailList))
                //        duplicatesEmailList = "Outlook contacts with the same email have been found and cannot be synchronized. Please delete duplicates of:";

                //    if (duplicates1.Count > 1)
                //        foreach (Outlook.ContactItem duplicate in duplicates1)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email1Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    if (duplicates2.Count > 1)
                //        foreach (Outlook.ContactItem duplicate in duplicates2)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email2Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    if (duplicates3.Count > 1)
                //        foreach (Outlook.ContactItem duplicate in duplicates3)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email3Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    continue;
                //}
                //else if (!string.IsNullOrEmpty(olc.Email1Address))
                //{
                //    ContactMatch dup = result.Find(delegate(ContactMatch m)
                //    {
                //        return m.OutlookContact != null && m.OutlookContact.Email1Address == olc.Email1Address;
                //    });
                //    if (dup != null)
                //    {
                //        Logger.Log(string.Format("Duplicate contact found by Email1Address ({0}). Skipping", olc.FileAs), EventType.Information);
                //        continue;
                //    }
                //}

                //// check for unique mobile phone, because this sync tool uses the also the mobile phone to identify matches between Google and Outlook
                //Collection<Outlook.ContactItem> duplicatesMobile = new Collection<Outlook.ContactItem>();
                //if (!string.IsNullOrEmpty(olc.MobileTelephoneNumber))
                //    duplicatesMobile = sync.OutlookContactByProperty("MobileTelephoneNumber", olc.MobileTelephoneNumber);

                //if (duplicatesMobile.Count > 1)
                //{
                //    if (string.IsNullOrEmpty(duplicatesMobileList))
                //        duplicatesMobileList = "Outlook contacts with the same mobile phone have been found and cannot be synchronized. Please delete duplicates of:";

                //    foreach (Outlook.ContactItem duplicate in duplicatesMobile)
                //    {
                //        sync.OutlookContactDuplicates.Add(olc);
                //        string str = olc.FileAs + " (" + olc.MobileTelephoneNumber + ")";
                //        if (!duplicatesMobileList.Contains(str))
                //            duplicatesMobileList += Environment.NewLine + str;
                //    }
                //    continue;
                //}
                //else if (!string.IsNullOrEmpty(olc.MobileTelephoneNumber))
                //{
                //    ContactMatch dup = result.Find(delegate(ContactMatch m)
                //    {
                //        return m.OutlookContact != null && m.OutlookContact.MobileTelephoneNumber == olc.MobileTelephoneNumber;
                //    });
                //    if (dup != null)
                //    {
                //        Logger.Log(string.Format("Duplicate contact found by MobileTelephoneNumber ({0}). Skipping", olc.FileAs), EventType.Information);
                //        continue;
                //    }
                //}

                #endregion

                if (match.AllGoogleContactMatches == null || match.AllGoogleContactMatches.Count == 0)
                {
                    //Check, if this Outlook contact has a m in the google duplicates
                    bool duplicateFound = false;
                    foreach (ContactMatch duplicate in sync.GoogleContactDuplicates)
                    {
                        string entryTitleFirstLastAndSuffix = OutlookContactInfo.GetTitleFirstLastAndSuffix(duplicate.AllGoogleContactMatches[0]);
                        if (duplicate.AllGoogleContactMatches.Count > 0 &&
                            (!string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(duplicate.AllGoogleContactMatches[0].Title) && olci.FileAs.Equals(duplicate.AllGoogleContactMatches[0].Title.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||  //Replace twice to not replace a \r\n by \r\r\n. This is necessary because \r\n are saved as \n only to google
                             !string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(duplicate.AllGoogleContactMatches[0].Name.FullName) && olci.FileAs.Equals(duplicate.AllGoogleContactMatches[0].Name.FullName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
                             !string.IsNullOrEmpty(olci.FullName) && !string.IsNullOrEmpty(duplicate.AllGoogleContactMatches[0].Name.FullName) && olci.FullName.Equals(duplicate.AllGoogleContactMatches[0].Name.FullName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
                             !string.IsNullOrEmpty(olci.TitleFirstLastAndSuffix) && !string.IsNullOrEmpty(entryTitleFirstLastAndSuffix) && olci.TitleFirstLastAndSuffix.Equals(entryTitleFirstLastAndSuffix.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
                             !string.IsNullOrEmpty(olci.Email1Address) && duplicate.AllGoogleContactMatches[0].Emails.Count > 0 && olci.Email1Address.Equals(duplicate.AllGoogleContactMatches[0].Emails[0].Address, StringComparison.InvariantCultureIgnoreCase) ||
                             //!string.IsNullOrEmpty(olci.Email2Address) && FindEmail(olci.Email2Address, duplicate.AllGoogleContactMatches[0].Emails) != null ||
                             //!string.IsNullOrEmpty(olci.Email3Address) && FindEmail(olci.Email3Address, duplicate.AllGoogleContactMatches[0].Emails) != null ||
                             olci.MobileTelephoneNumber != null && FindPhone(olci.MobileTelephoneNumber, duplicate.AllGoogleContactMatches[0].Phonenumbers) != null ||
                             !string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(duplicate.AllGoogleContactMatches[0].Title) && duplicate.AllGoogleContactMatches[0].Organizations.Count > 0 && olci.FileAs.Equals(duplicate.AllGoogleContactMatches[0].Organizations[0].Name, StringComparison.InvariantCultureIgnoreCase)
                            ) ||
                            !string.IsNullOrEmpty(olci.FileAs) && olci.FileAs.Equals(duplicate.OutlookContact.FileAs, StringComparison.InvariantCultureIgnoreCase) ||
                            !string.IsNullOrEmpty(olci.FullName) && olci.FullName.Equals(duplicate.OutlookContact.FullName, StringComparison.InvariantCultureIgnoreCase) ||
                            !string.IsNullOrEmpty(olci.TitleFirstLastAndSuffix) && olci.TitleFirstLastAndSuffix.Equals(duplicate.OutlookContact.TitleFirstLastAndSuffix, StringComparison.InvariantCultureIgnoreCase) ||
                            !string.IsNullOrEmpty(olci.Email1Address) && olci.Email1Address.Equals(duplicate.OutlookContact.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
                            //                                              olci.Email1Address.Equals(duplicate.OutlookContact.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
                            //                                              olci.Email1Address.Equals(duplicate.OutlookContact.Email3Address, StringComparison.InvariantCultureIgnoreCase)
                            //                                              ) ||
                            //!string.IsNullOrEmpty(olci.Email2Address) && (olci.Email2Address.Equals(duplicate.OutlookContact.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
                            //                                              olci.Email2Address.Equals(duplicate.OutlookContact.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
                            //                                              olci.Email2Address.Equals(duplicate.OutlookContact.Email3Address, StringComparison.InvariantCultureIgnoreCase)
                            //                                              ) ||
                            //!string.IsNullOrEmpty(olci.Email3Address) && (olci.Email3Address.Equals(duplicate.OutlookContact.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
                            //                                              olci.Email3Address.Equals(duplicate.OutlookContact.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
                            //                                              olci.Email3Address.Equals(duplicate.OutlookContact.Email3Address, StringComparison.InvariantCultureIgnoreCase)
                            //                                              ) ||
                            olci.MobileTelephoneNumber != null && olci.MobileTelephoneNumber.Equals(duplicate.OutlookContact.MobileTelephoneNumber) ||
                            !string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(duplicate.GoogleContact.Title) && duplicate.GoogleContact.Organizations.Count > 0 && olci.FileAs.Equals(duplicate.GoogleContact.Organizations[0].Name, StringComparison.InvariantCultureIgnoreCase)
                           )
                        {
                            duplicateFound = true;
                            duplicate.AddOutlookContact(olci);
                            sync.OutlookContactDuplicates.Add(match);
                            if (string.IsNullOrEmpty(duplicateOutlookContacts))
                                duplicateOutlookContacts = "Outlook contact found that has been already identified as duplicate Google contact (either same email, Mobile or FullName) and cannot be synchronized. Please delete or resolve duplicates of:";

                            string str = olci.FileAs + " (" + olci.Email1Address + ", " + olci.MobileTelephoneNumber + ")";
                            if (!duplicateOutlookContacts.Contains(str))
                                duplicateOutlookContacts += Environment.NewLine + str;

                            break;
                        }
                    }

                    if (!duplicateFound)
                        Logger.Log(string.Format("No match found for outlook contact ({0}) => {1}", olci.FileAs, (olci.UserProperties.GoogleContactId != null ? "Delete from Outlook" : "Add to Google")), EventType.Information);
                }
                else
                {
                    //Remember Google duplicates to later react to it when resetting matches or syncing
                    //ResetMatches: Also reset the duplicates
                    //Sync: Skip duplicates (don't sync duplicates to be fail safe)
                    if (match.AllGoogleContactMatches.Count > 1)
                    {
                        sync.GoogleContactDuplicates.Add(match);
                        foreach (Contact entry in match.AllGoogleContactMatches)
                        {
                            //Create message for duplicatesFound exception
                            if (string.IsNullOrEmpty(duplicateGoogleMatches))
                                duplicateGoogleMatches = "Outlook contacts matching with multiple Google contacts have been found (either same email, Mobile, FullName or company) and cannot be synchronized. Please delete or resolve duplicates of:";

                            string str = olci.FileAs + " (" + olci.Email1Address + ", " + olci.MobileTelephoneNumber + ")";
                            if (!duplicateGoogleMatches.Contains(str))
                                duplicateGoogleMatches += Environment.NewLine + str;
                        }
                    }
                }

                result.Add(match);
            }
            #endregion

            if (!string.IsNullOrEmpty(duplicateGoogleMatches) || !string.IsNullOrEmpty(duplicateOutlookContacts))
                duplicatesFound = new DuplicateDataException(duplicateGoogleMatches + Environment.NewLine + Environment.NewLine + duplicateOutlookContacts);
            else
                duplicatesFound = null;

            //return result;

            //for each google contact that's left (they will be nonmatched) create a new m pair without outlook contact. 
            for (int i = 0; i < sync.GoogleContacts.Count; i++)
            {
                var entry = sync.GoogleContacts[i];
                NotificationReceived?.Invoke(string.Format("Adding new Google contact {0} of {1} by unique properties: {2} ...", i + 1, sync.GoogleContacts.Count, entry.Title));

                // only m if there is either an email or mobile phone or a name or a company
                // otherwise a matching google contact will be created at each sync
                bool mobileExists = false;
                foreach (PhoneNumber phone in entry.Phonenumbers)
                {
                    if (phone.Rel == ContactsRelationships.IsMobile)
                    {
                        mobileExists = true;
                        break;
                    }
                }

                var googleOutlookId = ContactPropertiesUtils.GetGoogleOutlookContactId(sync.SyncProfile, entry);
                if (!string.IsNullOrEmpty(googleOutlookId) && skippedOutlookIds.Contains(googleOutlookId))
                {
                    Logger.Log("Skipped GoogleContact because Outlook contact couldn't be matched because of previous problem (see log): " + entry.Title, EventType.Warning);
                }
                else if (entry.Emails.Count == 0 && !mobileExists && string.IsNullOrEmpty(entry.Title) && (entry.Organizations.Count == 0 || string.IsNullOrEmpty(entry.Organizations[0].Name)))
                {
                    // no telephone and email

                    //ToDo: For now I use the ResolveDelete function, because it is almost the same, maybe we introduce a separate function for this ans also include DeleteGoogleAlways checkbox
                    using (var r = new ConflictResolver())
                    {
                        var res = r.ResolveDelete(entry);

                        if (res == DeleteResolution.DeleteGoogle || res == DeleteResolution.DeleteGoogleAlways)
                        {
                            ContactPropertiesUtils.SetGoogleOutlookContactId(sync.SyncProfile, entry, "-1"); //just set a dummy Id to delete this entry later on
                            sync.SaveContact(new ContactMatch(null, entry));
                        }
                        else
                        {
                            sync.SkippedCount++;
                            sync.SkippedCountNotMatches++;

                            Logger.Log("Skipped GoogleContact because no unique property found (Email1 or mobile or name or company):" + ContactMatch.GetSummary(entry), EventType.Warning);
                        }
                    }
                }
                else
                {
                    Logger.Log(string.Format("No match found for Google contact ({0}) => {1}", entry.Title, (!string.IsNullOrEmpty(googleOutlookId) ? "Delete from Google" : "Add to Outlook")), EventType.Information);
                    var match = new ContactMatch(null, entry);
                    result.Add(match);
                }
            }
            return result;
        }

        private static bool IsContactValid(Outlook.ContactItem contact)
        {
            /*if (!string.IsNullOrEmpty(contact.FileAs))
				return true;*/

            if (!string.IsNullOrEmpty(contact.Email1Address))
                return true;

            if (!string.IsNullOrEmpty(contact.Email2Address))
                return true;

            if (!string.IsNullOrEmpty(contact.Email3Address))
                return true;

            if (!string.IsNullOrEmpty(contact.HomeTelephoneNumber))
                return true;

            if (!string.IsNullOrEmpty(contact.BusinessTelephoneNumber))
                return true;

            if (!string.IsNullOrEmpty(contact.MobileTelephoneNumber))
                return true;

            if (!string.IsNullOrEmpty(contact.HomeAddress))
                return true;

            if (!string.IsNullOrEmpty(contact.BusinessAddress))
                return true;

            if (!string.IsNullOrEmpty(contact.OtherAddress))
                return true;

            if (!string.IsNullOrEmpty(contact.Body))
                return true;

            if (contact.Birthday != DateTime.MinValue)
                return true;

            return false;
        }

        public static void SyncContacts(ContactsSynchronizer sync)
        {
            for (int i = 0; i < sync.Contacts.Count; i++)
            {
                ContactMatch match = sync.Contacts[i];
                NotificationReceived?.Invoke(string.Format("Syncing contact {0} of {1}: {2} ...", i + 1, sync.Contacts.Count, match.ToString()));
                SyncContact(match, sync);
            }
        }
        public static void SyncContact(ContactMatch m, ContactsSynchronizer sync)
        {
            var olc = m.OutlookContact != null ? m.OutlookContact.GetOriginalItemFromOutlook() : null;

            try
            {
                if (m.GoogleContact == null && m.OutlookContact != null)
                {
                    //no google contact                               
                    var gid = m.OutlookContact.UserProperties.GoogleContactId;
                    if (!string.IsNullOrEmpty(gid))
                    {
                        //Redundant check if exist, but in case an error occurred in MatchContacts
                        var matchingGoogleContact = sync.GetGoogleContactById(gid);
                        if (matchingGoogleContact == null)
                        {
                            if (sync.SyncOption == SyncOption.OutlookToGoogleOnly || !sync.SyncDelete)
                                return;
                            else if (!sync.PromptDelete)
                                sync.DeleteOutlookResolution = DeleteResolution.DeleteOutlookAlways;
                            else if (sync.DeleteOutlookResolution != DeleteResolution.DeleteOutlookAlways &&
                                     sync.DeleteOutlookResolution != DeleteResolution.KeepOutlookAlways)
                            {
                                using (var r = new ConflictResolver())
                                {
                                    sync.DeleteOutlookResolution = r.ResolveDelete(m.OutlookContact);
                                }
                            }
                            switch (sync.DeleteOutlookResolution)
                            {
                                case DeleteResolution.KeepOutlook:
                                case DeleteResolution.KeepOutlookAlways:
                                    ContactPropertiesUtils.ResetOutlookGoogleContactId(sync, olc);
                                    break;
                                case DeleteResolution.DeleteOutlook:
                                case DeleteResolution.DeleteOutlookAlways:
                                    //Avoid recreating a GoogleContact already existing
                                    //==> Delete this outlookContact instead if previous m existed but no m exists anymore
                                    return;
                                default:
                                    throw new ApplicationException("Cancelled");
                            }
                        }
                    }

                    if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        sync.SkippedCount++;
                        Logger.Log(string.Format("Outlook Contact not added to Google, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", m.OutlookContact.FileAs), EventType.Information);
                        return;
                    }

                    //create a Google contact from Outlook contact
                    m.GoogleContact = new Contact();

                    sync.UpdateContact(olc, m.GoogleContact);

                }
                else if (m.OutlookContact == null && m.GoogleContact != null)
                {
                    // no outlook contact
                    var oid = ContactPropertiesUtils.GetGoogleOutlookContactId(sync.SyncProfile, m.GoogleContact);
                    if (oid != null)
                    {
                        if (sync.SyncOption == SyncOption.GoogleToOutlookOnly || !sync.SyncDelete)
                            return;
                        else if (!sync.PromptDelete)
                            sync.DeleteGoogleResolution = DeleteResolution.DeleteGoogleAlways;
                        else if (sync.DeleteGoogleResolution != DeleteResolution.DeleteGoogleAlways &&
                                 sync.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                        {
                            using (var r = new ConflictResolver())
                            {
                                sync.DeleteGoogleResolution = r.ResolveDelete(m.GoogleContact);
                            }
                        }
                        switch (sync.DeleteGoogleResolution)
                        {
                            case DeleteResolution.KeepGoogle:
                            case DeleteResolution.KeepGoogleAlways:
                                ContactPropertiesUtils.ResetGoogleOutlookContactId(sync.SyncProfile, m.GoogleContact);
                                break;
                            case DeleteResolution.DeleteGoogle:
                            case DeleteResolution.DeleteGoogleAlways:
                                //Avoid recreating a OutlookContact already existing
                                //==> Delete this googleContact instead if previous m existed but no m exists anymore                
                                return;
                            default:
                                throw new ApplicationException("Cancelled");
                        }

                    }

                    if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        sync.SkippedCount++;
                        Logger.Log(string.Format("Google Contact not added to Outlook, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", m.GoogleContact.Title), EventType.Information);
                        return;
                    }

                    //create a Outlook contact from Google contact                                                            
                    olc = ContactsSynchronizer.CreateOutlookContactItem(ContactsSynchronizer.SyncContactsFolder);

                    sync.UpdateContact(m.GoogleContact, olc);
                    m.OutlookContact = new OutlookContactInfo(olc, sync);
                }
                else if (m.OutlookContact != null && m.GoogleContact != null)
                {
                    //merge contact details                

                    //determine if this contact pair were synchronized
                    //DateTime? lastUpdated = GetOutlookPropertyValueDateTime(m.OutlookContact, sync.OutlookPropertyNameUpdated);
                    var lastSynced = m.OutlookContact.UserProperties.LastSync;
                    if (lastSynced.HasValue)
                    {
                        //contact pair was syncronysed before.

                        //determine if google contact was updated since last sync

                        //lastSynced is stored without seconds. take that into account.
                        var lastUpdatedOutlook = m.OutlookContact.LastModificationTime.AddSeconds(-m.OutlookContact.LastModificationTime.Second);
                        var lastUpdatedGoogle = m.GoogleContact.Updated.AddSeconds(-m.GoogleContact.Updated.Second);

                        //check if both outlok and google contacts where updated sync last sync
                        if (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance
                            && lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance)
                        {
                            //both contacts were updated.
                            //options: 1) ignore 2) loose one based on SyncOption
                            //throw new Exception("Both contacts were updated!");

                            switch (sync.SyncOption)
                            {
                                case SyncOption.MergeOutlookWins:
                                case SyncOption.OutlookToGoogleOnly:
                                    //overwrite google contact
                                    Logger.Log("Outlook and Google contact have been updated, Outlook contact is overwriting Google because of SyncOption " + sync.SyncOption + ": " + m.OutlookContact.FileAs + ".", EventType.Information);
                                    sync.UpdateContact(olc, m.GoogleContact);
                                    break;
                                case SyncOption.MergeGoogleWins:
                                case SyncOption.GoogleToOutlookOnly:
                                    //overwrite outlook contact
                                    Logger.Log("Outlook and Google contact have been updated, Google contact is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + m.OutlookContact.FileAs + ".", EventType.Information);
                                    sync.UpdateContact(m.GoogleContact, olc);
                                    break;
                                case SyncOption.MergePrompt:
                                    //promp for sync option
                                    if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                        sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                        sync.ConflictResolution != ConflictResolution.SkipAlways)
                                    {
                                        using (var r = new ConflictResolver())
                                        {
                                            sync.ConflictResolution = r.Resolve(m, false);
                                        }
                                    }
                                    switch (sync.ConflictResolution)
                                    {
                                        case ConflictResolution.Skip:
                                        case ConflictResolution.SkipAlways:
                                            Logger.Log(string.Format("User skipped contact ({0}).", m.ToString()), EventType.Information);
                                            sync.SkippedCount++;
                                            break;
                                        case ConflictResolution.OutlookWins:
                                        case ConflictResolution.OutlookWinsAlways:
                                            sync.UpdateContact(olc, m.GoogleContact);
                                            break;
                                        case ConflictResolution.GoogleWins:
                                        case ConflictResolution.GoogleWinsAlways:
                                            sync.UpdateContact(m.GoogleContact, olc);
                                            break;
                                        default:
                                            throw new ApplicationException("Cancelled");
                                    }
                                    break;
                            }
                            return;
                        }

                        //check if outlook contact was updated (with X second tolerance)
                        if (sync.SyncOption != SyncOption.GoogleToOutlookOnly &&
                            (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                             lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                             sync.SyncOption == SyncOption.OutlookToGoogleOnly
                            )
                           )
                        {
                            //outlook contact was changed or changed Google contact will be overwritten

                            if (lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                                sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                                Logger.Log("Google contact has been updated since last sync, but Outlook contact is overwriting Google because of SyncOption " + sync.SyncOption + ": " + m.OutlookContact.FileAs + ".", EventType.Information);

                            sync.UpdateContact(olc, m.GoogleContact);

                            //at the moment use outlook as "master" source of contacts - in the event of a conflict google contact will be overwritten.
                            //TODO: control conflict resolution by SyncOption
                            return;
                        }

                        //check if google contact was updated (with X second tolerance)
                        if (sync.SyncOption != SyncOption.OutlookToGoogleOnly &&
                            (lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                             lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                             sync.SyncOption == SyncOption.GoogleToOutlookOnly
                            )
                           )
                        {
                            //google contact was changed or changed Outlook contact will be overwritten

                            if (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                                sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                                Logger.Log("Outlook contact has been updated since last sync, but Google contact is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + m.OutlookContact.FileAs + ".", EventType.Information);

                            sync.UpdateContact(m.GoogleContact, olc);
                        }
                    }
                    else
                    {
                        //contacts were never synced.
                        //merge contacts.
                        switch (sync.SyncOption)
                        {
                            case SyncOption.MergeOutlookWins:
                            case SyncOption.OutlookToGoogleOnly:
                                //overwrite google contact
                                sync.UpdateContact(olc, m.GoogleContact);
                                break;
                            case SyncOption.MergeGoogleWins:
                            case SyncOption.GoogleToOutlookOnly:
                                //overwrite outlook contact
                                sync.UpdateContact(m.GoogleContact, olc);
                                break;
                            case SyncOption.MergePrompt:
                                //promp for sync option
                                if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                    sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                    sync.ConflictResolution != ConflictResolution.SkipAlways)
                                {
                                    using (var r = new ConflictResolver())
                                    {
                                        sync.ConflictResolution = r.Resolve(m, true);
                                    }
                                }
                                switch (sync.ConflictResolution)
                                {
                                    case ConflictResolution.Skip:
                                    case ConflictResolution.SkipAlways: //Keep both, Google AND Outlook
                                        sync.Contacts.Add(new ContactMatch(m.OutlookContact, null));
                                        sync.Contacts.Add(new ContactMatch(null, m.GoogleContact));
                                        break;
                                    case ConflictResolution.OutlookWins:
                                    case ConflictResolution.OutlookWinsAlways:
                                        sync.UpdateContact(olc, m.GoogleContact);
                                        break;
                                    case ConflictResolution.GoogleWins:
                                    case ConflictResolution.GoogleWinsAlways:
                                        sync.UpdateContact(m.GoogleContact, olc);
                                        break;
                                    default:
                                        throw new ApplicationException("Cancelled");
                                }
                                break;
                        }
                    }
                }
                else
                    throw new ArgumentNullException("ContactMatch has all peers null.");
            }
            catch (ArgumentNullException)
            {
                throw;
            }
            catch (Exception e)
            {
                throw new Exception("Error syncing contact " + (m.OutlookContact != null ? m.OutlookContact.FileAs : m.GoogleContact.Title) + ": " + e.Message, e);
            }
            finally
            {
                if (olc != null && m.OutlookContact != null)
                {
                    m.OutlookContact.Update(olc, sync);
                    Marshal.ReleaseComObject(olc);
                    olc = null;
                }
            }
        }

        private static PhoneNumber FindPhone(string number, ExtensionCollection<PhoneNumber> phones)
        {
            if (string.IsNullOrEmpty(number))
                return null;

            if (phones == null)
                return null;

            foreach (PhoneNumber phone in phones)
            {
                if (phone != null && number.Equals(phone.Value, StringComparison.InvariantCultureIgnoreCase))
                {
                    return phone;
                }
            }

            return null;
        }

        private static EMail FindEmail(string address, ExtensionCollection<EMail> emails)
        {
            if (string.IsNullOrEmpty(address))
                return null;

            foreach (EMail email in emails)
            {
                if (address.Equals(email.Address, StringComparison.InvariantCultureIgnoreCase))
                {
                    return email;
                }
            }

            return null;
        }


        /// <summary>
        /// Adds new Google Groups to the Google account.
        /// </summary>
        /// <param name="sync"></param>
        public static void SyncGroups(ContactsSynchronizer sync)
        {
            foreach (ContactMatch match in sync.Contacts)
            {
                if (match.OutlookContact != null && !string.IsNullOrEmpty(match.OutlookContact.Categories))
                {
                    string[] cats = Utilities.GetOutlookGroups(match.OutlookContact.Categories);
                    
                    foreach (string cat in cats)
                    {
                        var g = sync.GetGoogleGroupByName(cat);
                        if (g == null)
                        {
                            // create group                            
                            g = sync.CreateGroup(cat);
                            g = sync.SaveGoogleGroup(g);
                            sync.GoogleGroups.Add(g);
                        }
                    }
                }
            }
        }
    }

    //internal class List<ContactMatch> : List<ContactMatch>
    //{
    //    public List<ContactMatch>(int capacity) : base(capacity) { }
    //}

    internal class ContactMatch
    {
        public OutlookContactInfo OutlookContact;
        public Contact GoogleContact;
        public readonly List<Contact> AllGoogleContactMatches = new List<Contact>();
        public readonly List<OutlookContactInfo> AllOutlookContactMatches = new List<OutlookContactInfo>();
        //public Contact LastGoogleContact;

        public ContactMatch(OutlookContactInfo outlookContact, Contact googleContact)
        {
            AddOutlookContact(outlookContact);
            AddGoogleContact(googleContact);
        }

        public void AddGoogleContact(Contact googleContact)
        {
            if (googleContact == null)
                return;
            //throw new ArgumentNullException("googleContact must not be null.");

            if (GoogleContact == null)
                GoogleContact = googleContact;

            //this to avoid searching the entire collection. 
            //if last contact it what we are trying to add the we have already added it earlier
            //if (LastGoogleContact == googleContact)
            //    return;

            if (!AllGoogleContactMatches.Contains(googleContact))
                AllGoogleContactMatches.Add(googleContact);

            //LastGoogleContact = googleContact;
        }

        public void AddOutlookContact(OutlookContactInfo outlookContact)
        {
            if (outlookContact == null)
                return;
            //throw new ArgumentNullException("outlookContact must not be null.");

            if (OutlookContact == null)
                OutlookContact = outlookContact;

            //this to avoid searching the entire collection. 
            //if last contact it what we are trying to add the we have already added it earlier
            //if (LastGoogleContact == googleContact)
            //    return;

            if (!AllOutlookContactMatches.Contains(outlookContact))
                AllOutlookContactMatches.Add(outlookContact);

            //LastGoogleContact = googleContact;
        }


        public override string ToString()
        {
            if (OutlookContact != null)
                return GetName(OutlookContact);
            else if (GoogleContact != null)
                return GetName(GoogleContact);
            else
                return string.Empty;
        }


        public static string GetName(OutlookContactInfo outlookContact)
        {
            string name = outlookContact.FileAs;
            if (string.IsNullOrEmpty(name))
                name = outlookContact.FullName;
            if (string.IsNullOrEmpty(name))
                name = outlookContact.Company;
            if (string.IsNullOrEmpty(name))
                name = outlookContact.Email1Address;

            return name;
        }


        public static string GetName(Contact googleContact)
        {
            string name = googleContact.Title;
            if (string.IsNullOrEmpty(name))
                name = googleContact.Name.FullName;
            if (string.IsNullOrEmpty(name) && googleContact.Organizations.Count > 0)
                name = googleContact.Organizations[0].Name;
            if (string.IsNullOrEmpty(name) && googleContact.Emails.Count > 0)
                name = googleContact.Emails[0].Address;

            return name;
        }

        public static string GetSummary(Outlook.ContactItem outlookContact)
        {
            string name = OutlookContactInfo.GetTitleFirstLastAndSuffix(outlookContact);

            string summary = string.Empty;

            if (!string.IsNullOrEmpty(name))
                summary += "Name: " + name.Trim().Replace("  ", " ") + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.FirstName))
                summary += "Firstname: " + outlookContact.FirstName + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.LastName))
                summary += "Lastname: " + outlookContact.LastName + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.CompanyName))
                summary += "Company: " + outlookContact.CompanyName + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Department))
                summary += "Department: " + outlookContact.Department + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.JobTitle))
                summary += "JobTitle: " + outlookContact.JobTitle + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Email1Address))
                summary += "Email1: " + outlookContact.Email1Address + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Email2Address))
                summary += "Email2: " + outlookContact.Email2Address + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Email3Address))
                summary += "Email3: " + outlookContact.Email3Address + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.MobileTelephoneNumber))
                summary += "MobilePhone: " + outlookContact.MobileTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.HomeTelephoneNumber))
                summary += "HomePhone: " + outlookContact.HomeTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Home2TelephoneNumber))
                summary += "HomePhone2: " + outlookContact.Home2TelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.BusinessTelephoneNumber))
                summary += "BusinessPhone: " + outlookContact.BusinessTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Business2TelephoneNumber))
                summary += "BusinessPhone2: " + outlookContact.Business2TelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.OtherTelephoneNumber))
                summary += "OtherPhone: " + outlookContact.OtherTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.HomeAddress))
                summary += "HomeAddress: " + outlookContact.HomeAddress + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.BusinessAddress))
                summary += "BusinessAddress: " + outlookContact.BusinessAddress + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.OtherAddress))
                summary += "OtherAddress: " + outlookContact.OtherAddress + "\r\n";

            return summary;
        }

        public static string GetSummary(Contact googleContact)
        {
            string name = OutlookContactInfo.GetTitleFirstLastAndSuffix(googleContact);

            string summary = string.Empty;

            if (!string.IsNullOrEmpty(name))
                summary += "Name: " + name.Trim().Replace("  ", " ") + "\r\n";
            if (!string.IsNullOrEmpty(googleContact.Name.GivenName))
                summary += "Firstname: " + googleContact.Name.GivenName + "\r\n";
            if (!string.IsNullOrEmpty(googleContact.Name.GivenName))
                summary += "Lastname: " + googleContact.Name.FamilyName + "\r\n";
            for (int i = 0; i < googleContact.Organizations.Count; i++)
            {
                string company = googleContact.Organizations[i].Name;
                string department = googleContact.Organizations[i].Department;
                string jobTitle = googleContact.Organizations[i].JobDescription;
                if (!string.IsNullOrEmpty(company))
                {
                    summary += "Company: " + company + "\r\n";
                }
                if (!string.IsNullOrEmpty(department))
                {
                    summary += "Department: " + department + "\r\n";
                }
                if (!string.IsNullOrEmpty(jobTitle))
                {
                    summary += "JobTitle: " + jobTitle + "\r\n";
                }
            }
            for (int i = 0; i < googleContact.Emails.Count; i++)
            {
                string email = googleContact.Emails[i].Address;
                if (!string.IsNullOrEmpty(email))
                {
                    summary += "Email" + (i + 1) + ": " + email + "\r\n";
                }
            }
            foreach (PhoneNumber phone in googleContact.Phonenumbers)
            {
                if (!string.IsNullOrEmpty(phone.Value))
                {
                    if (phone.Rel == ContactsRelationships.IsMobile)
                        summary += "MobilePhone: ";
                    if (phone.Rel == ContactsRelationships.IsHome)
                        summary += "HomePhone: ";
                    if (phone.Rel == ContactsRelationships.IsWork)
                        summary += "BusinessPhone: ";
                    if (phone.Rel == ContactsRelationships.IsOther)
                        summary += "OtherPhone: ";

                    summary += phone.Value + "\r\n";
                }


            }

            foreach (StructuredPostalAddress address in googleContact.PostalAddresses)
            {
                if (!string.IsNullOrEmpty(address.FormattedAddress))
                {
                    if (address.Rel == ContactsRelationships.IsHome)
                        summary += "HomeAddress: ";
                    if (address.Rel == ContactsRelationships.IsWork)
                        summary += "BusinessAddress: ";
                    if (address.Rel == ContactsRelationships.IsOther)
                        summary += "OtherAddress: ";

                    summary += address.FormattedAddress + "\r\n";
                }
            }

            return summary;
        }
    }



}
