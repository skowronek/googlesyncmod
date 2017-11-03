﻿using System;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    /// <summary>
    /// Holds information about an Outlook contact during processing.
    /// We can not always instantiate an unlimited number of Exchange Outlook objects (policy limitations), 
    /// so instead we copy the info we need for our processing into instances of OutlookContactInfo and only
    /// get the real Outlook.ContactItem objects when needed to communicate with Outlook.
    /// </summary>
    class OutlookContactInfo
    {
        #region Internal classes
        internal class UserPropertiesHolder
        {
            public string GoogleContactId;
            public DateTime? LastSync;
        }
        #endregion

        #region Properties
        public string EntryID { get; set; }
        public string FileAs { get; set; }
        public string FullName { get; set; }
        public string TitleFirstLastAndSuffix { get; set; } //Additional unique identifier
        public string Email1Address { get; set; }
        public string MobileTelephoneNumber { get; set; }
        public string Categories { get; set; }
        public string Company { get; set; }
        public DateTime LastModificationTime { get; set; }
        public UserPropertiesHolder UserProperties { get; set; }
        #endregion

        #region Construction
        private OutlookContactInfo()
        {
            // Not public - we are always constructed from an Outlook.ContactItem (constructor below)
        }

        public OutlookContactInfo(ContactItem item, ContactsSynchronizer sync)
        {
            UserProperties = new UserPropertiesHolder();
            Update(item, sync);
        }
        #endregion

        internal void Update(ContactItem outlookContactItem, ContactsSynchronizer sync)
        {
            EntryID = outlookContactItem.EntryID;
            FileAs = outlookContactItem.FileAs;
            FullName = outlookContactItem.FullName;
            Email1Address = ContactPropertiesUtils.GetOutlookEmailAddress1(outlookContactItem);
            MobileTelephoneNumber = outlookContactItem.MobileTelephoneNumber;
            Categories = outlookContactItem.Categories;
            LastModificationTime = outlookContactItem.LastModificationTime;
            Company = outlookContactItem.CompanyName;
            TitleFirstLastAndSuffix = GetTitleFirstLastAndSuffix(outlookContactItem);

            UserProperties userProperties = outlookContactItem.UserProperties;
            UserProperty prop = userProperties[sync.OutlookPropertyNameId];
            UserProperties.GoogleContactId = prop != null ? string.Copy((string)prop.Value) : null;
            if (prop != null)
                Marshal.ReleaseComObject(prop);

            prop = userProperties[sync.OutlookPropertyNameSynced];
            UserProperties.LastSync = prop != null ? (DateTime)prop.Value : (DateTime?)null;
            if (prop != null)
                Marshal.ReleaseComObject(prop);

            Marshal.ReleaseComObject(userProperties);
        }

        internal ContactItem GetOriginalItemFromOutlook()
        {
            if (EntryID == null)
                throw new ApplicationException("OutlookContactInfo cannot re-create the ContactItem from Outlook because EntryID is null, suggesting that this OutlookContactInfo was not created from an existing Outook contact.");

            ContactItem outlookContactItem = Synchronizer.OutlookNameSpace.GetItemFromID(EntryID) as ContactItem;
            if (outlookContactItem == null)
                throw new ApplicationException("OutlookContactInfo cannot re-create the ContactItem from Outlook because there is no Outlook entry with this EntryID, suggesting that the existing Outook contact may have been deleted.");

            return outlookContactItem;
        }

        internal static string GetTitleFirstLastAndSuffix(ContactItem outlookContactItem)
        {
            return GetTitleFirstLastAndSuffix(outlookContactItem.Title, outlookContactItem.FirstName, outlookContactItem.MiddleName, outlookContactItem.LastName, outlookContactItem.Suffix);
        }

        internal static string GetTitleFirstLastAndSuffix(Google.Contacts.Contact googleContact)
        {
            return GetTitleFirstLastAndSuffix(googleContact.Name.NamePrefix, googleContact.Name.GivenName, googleContact.Name.AdditionalName, googleContact.Name.FamilyName, googleContact.Name.NameSuffix);
        }

        private static string GetTitleFirstLastAndSuffix(string title, string firstName, string middleName, string lastName, string suffix)
        {
            string ret = title + " " + firstName + " " + middleName + " " + lastName + " " + suffix;

            if (string.IsNullOrEmpty(ret.Trim()))
                ret = null;

            return ret;
        }
    }
}
