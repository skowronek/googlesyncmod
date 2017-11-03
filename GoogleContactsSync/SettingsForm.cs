using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections;
using System.Globalization;
using Google.Apis.Util.Store;
using Google.Apis.Calendar.v3.Data;
using System.Threading.Tasks;

namespace GoContactSyncMod
{
    internal partial class SettingsForm : Form
    {
        //Singleton-Object
        #region Singleton Definition

        private static volatile SettingsForm instance;
        private static object syncRoot = new object();

        public static SettingsForm Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new SettingsForm();
                    }
                }
                return instance;
            }
        }
        #endregion

        internal Synchronizer sync;
        private SyncOption syncOption;
        private DateTime lastSync;
        private bool requestClose = false;
        private bool boolShowBalloonTip = true;
        private CancellationTokenSource cancellationTokenSource;

        public const string AppRootKey = @"Software\GoContactSyncMOD";
        public const string RegistrySyncOption = "SyncOption";
        public const string RegistryUsername = "Username";
        public const string RegistryAutoSync = "AutoSync";
        public const string RegistryAutoSyncInterval = "AutoSyncInterval";
        public const string RegistryAutoStart = "AutoStart";
        public const string RegistryReportSyncResult = "ReportSyncResult";
        public const string RegistrySyncDeletion = "SyncDeletion";
        public const string RegistryPromptDeletion = "PromptDeletion";
        public const string RegistrySyncAppointmentsMonthsInPast = "SyncAppointmentsMonthsInPast";
        public const string RegistrySyncAppointmentsMonthsInFuture = "SyncAppointmentsMonthsInFuture";
        public const string RegistrySyncAppointmentsTimezone = "SyncAppointmentsTimezone";
        public const string RegistrySyncAppointments = "SyncAppointments";
        public const string RegistrySyncAppointmentsForceRTF = "SyncAppointmentsForceRTF";
        public const string RegistrySyncContacts = "SyncContacts";
        public const string RegistrySyncContactsForceRTF = "SyncContactsForceRTF";
        public const string RegistryUseFileAs = "UseFileAs";
        public const string RegistryLastSync = "LastSync";
        public const string RegistrySyncContactsFolder = "SyncContactsFolder";
        public const string RegistrySyncAppointmentsFolder = "SyncAppointmentsFolder";
        public const string RegistrySyncAppointmentsGoogleFolder = "SyncAppointmentsGoogleFolder";
        public const string RegistrySyncProfile = "SyncProfile";

        private ProxySettingsForm _proxy = new ProxySettingsForm();

        private string syncContactsFolder = "";
        private string syncAppointmentsFolder = "";
        private string syncAppointmentsGoogleFolder = "";
        private string Timezone = "";

        //private string _syncProfile;
        private static string SyncProfile
        {
            get
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                return (regKeyAppRoot.GetValue(RegistrySyncProfile) != null) ?
                       (string)regKeyAppRoot.GetValue(RegistrySyncProfile) : null;
            }
            set
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                if (value != null)
                {
                    regKeyAppRoot.SetValue(RegistrySyncProfile, value);
                }
            }
        }

        private string ProfileRegistry;
        private bool OutlookFoldersLoaded = false;

        private int executing; // make this static if you want this one-caller-only to
        // all objects instead of a single object

        Thread syncThread;

        //register window for lock/unlock messages of workstation
        //private bool registered = false;

        delegate void TextHandler(string text);
        delegate void SwitchHandler(bool value);
        delegate void IconHandler();
        delegate DialogResult DialogHandler(string text);
        delegate void OnTimeZoneChangesCallback(string timeZone);

        public DialogResult ShowDialog(string text)
        {
            if (InvokeRequired)
            {
                return (DialogResult)Invoke(new DialogHandler(ShowDialog), new object[] { text });
            }
            else
            {
                return MessageBox.Show(this, text, Application.ProductName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            }
        }

        private Icon IconError = Properties.Resources.sync_error;
        private Icon Icon0 = Properties.Resources.sync;
        private Icon Icon30 = Properties.Resources.sync_30;
        private Icon Icon60 = Properties.Resources.sync_60;
        private Icon Icon90 = Properties.Resources.sync_90;
        private Icon Icon120 = Properties.Resources.sync_120;
        private Icon Icon150 = Properties.Resources.sync_150;
        private Icon Icon180 = Properties.Resources.sync_180;
        private Icon Icon210 = Properties.Resources.sync_210;
        private Icon Icon240 = Properties.Resources.sync_240;
        private Icon Icon270 = Properties.Resources.sync_270;
        private Icon Icon300 = Properties.Resources.sync_300;
        private Icon Icon330 = Properties.Resources.sync_330;

        private SettingsForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            cancellationTokenSource = new CancellationTokenSource();
            InitializeComponent();
            Text = Text + " - " + Application.ProductVersion;
            Logger.LogUpdated += new Logger.LogUpdatedHandler(Logger_LogUpdated);
            Logger.Log("Started application " + Application.ProductName + " (" + Application.ProductVersion + ") on " + VersionInformation.GetWindowsVersion() + " and " + OutlookRegistryUtils.GetOutlookVersion(), EventType.Information);
            Logger.Log("Detailed log file created: " + Logger.Folder + "log.txt", EventType.Information);
            ContactsMatcher.NotificationReceived += new ContactsMatcher.NotificationHandler(OnNotificationReceived);
            AppointmentsMatcher.NotificationReceived += new AppointmentsMatcher.NotificationHandler(OnNotificationReceived);
            PopulateSyncOptionBox();

            //temporary remove the listener to avoid to load the settings twice, because it is set from SettingsForm.Designer.cs
            cmbSyncProfile.SelectedIndexChanged -= new EventHandler(cmbSyncProfile_SelectedIndexChanged);
            if (fillSyncProfileItems())
            {
                ProfileRegistry = cmbSyncProfile.Text;
            }
            else
            {
                ProfileRegistry = null;
            }
            LoadSettings(ProfileRegistry);

            //enable the listener
            cmbSyncProfile.SelectedIndexChanged += new EventHandler(cmbSyncProfile_SelectedIndexChanged);

            TimerSwitch(true);
            lastSyncLabel.Text = "Not synced";

            ValidateSyncButton();

            //Register Session Lock Event
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SystemEvents_SessionSwitch);
            //Register Power Mode Event
            SystemEvents.PowerModeChanged += new PowerModeChangedEventHandler(SystemEvents_PowerModeSwitch);
        }

        private void PopulateSyncOptionBox()
        {
            string str;
            for (int i = 0; i < 20; i++)
            {
                str = ((SyncOption)i).ToString();
                if (str == i.ToString())
                    break;

                // format (to add space before capital)
                MatchCollection matches = Regex.Matches(str, "[A-Z]");
                for (int k = 0; k < matches.Count; k++)
                {
                    str = str.Replace(str[matches[k].Index].ToString(), " " + str[matches[k].Index]);
                    matches = Regex.Matches(str, "[A-Z]");
                }
                str = str.Replace("  ", " ");
                // fix start
                str = str.Substring(1);

                syncOptionBox.Items.Add(str);
            }
        }

        private void fillSyncFolderItems()
        {
            if (InvokeRequired)
            {
                Invoke(new InvokeCallback(fillSyncFolderItems));
            }
            else
            {
                lock (syncRoot)
                {
                    if (OutlookFoldersLoaded)
                        return;

                    if (contactFoldersComboBox.DataSource == null || appointmentFoldersComboBox.DataSource == null ||
                        appointmentGoogleFoldersComboBox.DataSource == null && btSyncAppointments.Checked ||
                        contactFoldersComboBox.Items.Count == 0 || appointmentFoldersComboBox.Items.Count == 0 ||
                        appointmentGoogleFoldersComboBox.Items.Count == 0 && btSyncAppointments.Checked)
                    {
                        Logger.Log("Loading Outlook folders...", EventType.Information);

                        contactFoldersComboBox.Visible = btSyncContactsForceRTF.Visible = btSyncContacts.Checked;
                        labelTimezone.Visible = labelMonthsPast.Visible = labelMonthsFuture.Visible = btSyncAppointments.Checked;
                        appointmentFoldersComboBox.Visible = appointmentGoogleFoldersComboBox.Visible = futureMonthInterval.Visible = pastMonthInterval.Visible = appointmentTimezonesComboBox.Visible = btSyncAppointmentsForceRTF.Visible = btSyncAppointments.Checked;
                        cmbSyncProfile.Visible = true;

                        string defaultText = "    --- Select an Outlook folder ---";
                        ArrayList outlookContactFolders = new ArrayList();
                        ArrayList outlookAppointmentFolders = new ArrayList();

                        try
                        {
                            Cursor = Cursors.WaitCursor;
                            SuspendLayout();

                            contactFoldersComboBox.BeginUpdate();
                            appointmentFoldersComboBox.BeginUpdate();
                            contactFoldersComboBox.DataSource = null;
                            appointmentFoldersComboBox.DataSource = null;

                            var folders = Synchronizer.OutlookNameSpace.Folders;
                            for (int i = 1; i <= folders.Count; i++)
                            {
                                Microsoft.Office.Interop.Outlook.MAPIFolder folder = null;
                                try
                                {
                                    folder = folders[i] as Microsoft.Office.Interop.Outlook.MAPIFolder;
                                    GetOutlookMAPIFolders(outlookContactFolders, outlookAppointmentFolders, folder);
                                }
                                catch (Exception e)
                                {
                                    Logger.Log(e, EventType.Debug);
                                    Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
                                }
                                finally
                                {
                                    if (folder != null)
                                        Marshal.ReleaseComObject(folder);
                                }
                            }

                            if (outlookContactFolders != null)
                            {
                                outlookContactFolders.Sort();
                                outlookContactFolders.Insert(0, new OutlookFolder(defaultText, defaultText, false));
                                contactFoldersComboBox.DataSource = outlookContactFolders;
                                contactFoldersComboBox.DisplayMember = "DisplayName";
                                contactFoldersComboBox.ValueMember = "FolderID";
                            }

                            if (outlookAppointmentFolders != null)
                            {
                                outlookAppointmentFolders.Sort();
                                outlookAppointmentFolders.Insert(0, new OutlookFolder(defaultText, defaultText, false));
                                appointmentFoldersComboBox.DataSource = outlookAppointmentFolders;
                                appointmentFoldersComboBox.DisplayMember = "DisplayName";
                                appointmentFoldersComboBox.ValueMember = "FolderID";
                            }

                            contactFoldersComboBox.EndUpdate();
                            appointmentFoldersComboBox.EndUpdate();

                            contactFoldersComboBox.SelectedValue = defaultText;
                            appointmentFoldersComboBox.SelectedValue = defaultText;

                            //If user has not yet selected any folder, select one based on Outlook default folder
                            if (contactFoldersComboBox.SelectedIndex < 1)
                            {
                                foreach (OutlookFolder folder in contactFoldersComboBox.Items)
                                {
                                    if (folder.IsDefaultFolder)
                                    {
                                        contactFoldersComboBox.SelectedValue = folder.FolderID;
                                        break;
                                    }
                                }
                            }

                            //If user has not yet selected any folder, select one based on Outlook default folder
                            if (appointmentFoldersComboBox.SelectedIndex < 1)
                            {
                                foreach (OutlookFolder folder in appointmentFoldersComboBox.Items)
                                {
                                    if (folder.IsDefaultFolder)
                                    {
                                        appointmentFoldersComboBox.SelectedItem = folder;
                                        break;
                                    }
                                }
                            }

                            Logger.Log("Loaded Outlook folders.", EventType.Information);
                        }
                        catch (Exception e)
                        {
                            Logger.Log(e, EventType.Debug);
                            Logger.Log("Error getting available Outlook and Google folders: " + e.Message, EventType.Warning);
                        }
                        finally
                        {
                            Cursor = Cursors.Default;
                            ResumeLayout();
                        }
                    }
                    LoadSettingsFolders(ProfileRegistry);

                    if ((contactFoldersComboBox.SelectedIndex == -1) && (contactFoldersComboBox.Items.Count > 0))
                        contactFoldersComboBox.SelectedIndex = 0;

                    if ((appointmentFoldersComboBox.SelectedIndex == -1) && (appointmentFoldersComboBox.Items.Count > 0))
                        appointmentFoldersComboBox.SelectedIndex = 0;

                    OutlookFoldersLoaded = true;
                }
            }
        }

        public static void GetOutlookMAPIFolders(ArrayList outlookContactFolders, /*ArrayList outlookNoteFolders,*/ ArrayList outlookAppointmentFolders, Microsoft.Office.Interop.Outlook.MAPIFolder folder)
        {
            for (int i = 1; i <= folder.Folders.Count; i++)
            {
                Microsoft.Office.Interop.Outlook.MAPIFolder mapi = null;
                try
                {
                    mapi = folder.Folders[i] as Microsoft.Office.Interop.Outlook.MAPIFolder;
                    if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olContactItem)
                    {
                        bool isDefaultFolder = mapi.EntryID.Equals(Synchronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts).EntryID);
                        outlookContactFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                    }
                    if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
                    {
                        bool isDefaultFolder = mapi.EntryID.Equals(Synchronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar).EntryID);
                        outlookAppointmentFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                    }

                    if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olContactItem ||
                        mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
                        GetOutlookMAPIFolders(outlookContactFolders, outlookAppointmentFolders, mapi);
                }
                finally
                {
                    if (mapi != null)
                        Marshal.ReleaseComObject(mapi);
                }
            }
        }

        private void ClearSettings()
        {
            SetSyncOption(0);
            autoSyncCheckBox.Checked = runAtStartupCheckBox.Checked = reportSyncResultCheckBox.Checked = false;
            autoSyncInterval.Value = 120;
            _proxy.ClearSettings();
        }
        // Fill lists of sync profiles
        private bool fillSyncProfileItems()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            }

            bool vReturn = false;

            cmbSyncProfile.Items.Clear();
            cmbSyncProfile.Items.Add("[Add new profile...]");

            foreach (string subKeyName in regKeyAppRoot.GetSubKeyNames())
            {
                if (!string.IsNullOrEmpty(subKeyName))
                    cmbSyncProfile.Items.Add(subKeyName);
            }

            if (SyncProfile == null)
                SyncProfile = "Default_" + Environment.MachineName;

            if (cmbSyncProfile.Items.Count == 1)
                cmbSyncProfile.Items.Add(SyncProfile);
            else
                vReturn = true;

            cmbSyncProfile.Items.Add("[Configuration manager...]");
            cmbSyncProfile.Text = SyncProfile;

            return vReturn;
        }

        private void LoadSettings(string _profile)
        {
            Logger.Log("Loading settings from registry...", EventType.Information);
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + (_profile != null ? ('\\' + _profile) : ""));

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync" + (_profile != null ? ('\\' + _profile) : ""));
            }

            if (regKeyAppRoot.GetValue(RegistrySyncOption) != null)
            {
                syncOption = (SyncOption)regKeyAppRoot.GetValue(RegistrySyncOption);
                SetSyncOption((int)syncOption);
            }

            if (regKeyAppRoot.GetValue(RegistryUsername) != null)
            {
                UserName.Text = regKeyAppRoot.GetValue(RegistryUsername) as string;
            }

            //temporary remove listener
            autoSyncCheckBox.CheckedChanged -= new EventHandler(autoSyncCheckBox_CheckedChanged);

            ReadRegistryIntoCheckBox(autoSyncCheckBox, regKeyAppRoot.GetValue(RegistryAutoSync));
            ReadRegistryIntoNumber(autoSyncInterval, regKeyAppRoot.GetValue(RegistryAutoSyncInterval));
            ReadRegistryIntoCheckBox(runAtStartupCheckBox, regKeyAppRoot.GetValue(RegistryAutoStart));
            ReadRegistryIntoCheckBox(reportSyncResultCheckBox, regKeyAppRoot.GetValue(RegistryReportSyncResult));
            ReadRegistryIntoCheckBox(btSyncDelete, regKeyAppRoot.GetValue(RegistrySyncDeletion));
            ReadRegistryIntoCheckBox(btPromptDelete, regKeyAppRoot.GetValue(RegistryPromptDeletion));
            ReadRegistryIntoNumber(pastMonthInterval, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInPast));
            ReadRegistryIntoNumber(futureMonthInterval, regKeyAppRoot.GetValue(RegistrySyncAppointmentsMonthsInFuture));
            if (regKeyAppRoot.GetValue(RegistrySyncAppointmentsTimezone) != null)
                appointmentTimezonesComboBox.Text = regKeyAppRoot.GetValue(RegistrySyncAppointmentsTimezone) as string;
            ReadRegistryIntoCheckBox(btSyncAppointments, regKeyAppRoot.GetValue(RegistrySyncAppointments));

            ReadRegistryIntoCheckBox(btSyncContacts, regKeyAppRoot.GetValue(RegistrySyncContacts));
            ReadRegistryIntoCheckBox(chkUseFileAs, regKeyAppRoot.GetValue(RegistryUseFileAs));

            ReadRegistryIntoCheckBox(btSyncContactsForceRTF, regKeyAppRoot.GetValue(RegistrySyncContactsForceRTF));
            ReadRegistryIntoCheckBox(btSyncAppointmentsForceRTF, regKeyAppRoot.GetValue(RegistrySyncAppointmentsForceRTF));

            if (regKeyAppRoot.GetValue(RegistryLastSync) != null)
            {
                try
                {
                    lastSync = new DateTime(Convert.ToInt64(regKeyAppRoot.GetValue(RegistryLastSync)));
                    SetLastSyncText(lastSync.ToString());
                }
                catch (FormatException ex)
                {
                    Logger.Log("LastSyncDate couldn't be read from registry (" + regKeyAppRoot.GetValue(RegistryLastSync) + "): " + ex, EventType.Warning);
                }
            }

            //autoSyncCheckBox_CheckedChanged(null, null);
            btSyncContacts_CheckedChanged(null, null);

            _proxy.LoadSettings(_profile);

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                SaveSettings(_profile);
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Webgear\GOContactSync");
            }

            //enable temporary disabled listener
            autoSyncCheckBox.CheckedChanged += new EventHandler(autoSyncCheckBox_CheckedChanged);
        }

        private static void ReadRegistryIntoCheckBox(CheckBox checkbox, object registryEntry)
        {
            if (registryEntry != null)
            {
                try
                {
                    checkbox.Checked = Convert.ToBoolean(registryEntry);
                }
                catch (FormatException ex)
                {
                    Logger.Log(checkbox.Name + " couldn't be read from registry (" + registryEntry + "), was kept at default (" + checkbox.Checked + "): " + ex, EventType.Warning);
                }
            }
        }

        private static void ReadRegistryIntoNumber(NumericUpDown numericUpDown, object registryEntry)
        {
            if (registryEntry != null)
            {
                decimal interval = Convert.ToDecimal(registryEntry);
                if (interval < numericUpDown.Minimum)
                {
                    numericUpDown.Value = numericUpDown.Minimum;
                    Logger.Log(numericUpDown.Name + " read from registry was below range (" + interval + "), was set to minimum (" + numericUpDown.Minimum + ")", EventType.Warning);
                }
                else if (interval > numericUpDown.Maximum)
                {
                    numericUpDown.Value = numericUpDown.Maximum;
                    Logger.Log(numericUpDown.Name + " read from registry was above range (" + interval + "), was set to maximum (" + numericUpDown.Maximum + ")", EventType.Warning);
                }
                else
                    numericUpDown.Value = interval;
            }
        }

        private void LoadSettingsFolders(string _profile)
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + (_profile != null ? ('\\' + _profile) : ""));

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync" + (_profile != null ? ('\\' + _profile) : ""));
            }

            var regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncContactsFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                foreach (OutlookFolder i in contactFoldersComboBox.Items)
                {
                    if (i.FolderID == regKeyValueStr)
                    {
                        contactFoldersComboBox.SelectedValue = regKeyValueStr;
                        break;
                    }
                }
            }

            regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncAppointmentsFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                foreach (OutlookFolder i in appointmentFoldersComboBox.Items)
                {
                    if (i.FolderID == regKeyValueStr)
                    {
                        appointmentFoldersComboBox.SelectedValue = regKeyValueStr;
                        break;
                    }
                }
            }

            regKeyValueStr = regKeyAppRoot.GetValue(RegistrySyncAppointmentsGoogleFolder) as string;
            if (!string.IsNullOrEmpty(regKeyValueStr))
            {
                if (appointmentGoogleFoldersComboBox.DataSource == null)
                {
                    appointmentFoldersComboBox.BeginUpdate();
                    ArrayList list = new ArrayList();
                    list.Add(new GoogleCalendar(regKeyValueStr, regKeyValueStr, false));
                    appointmentGoogleFoldersComboBox.DataSource = list;
                    appointmentGoogleFoldersComboBox.DisplayMember = "DisplayName";
                    appointmentGoogleFoldersComboBox.ValueMember = "FolderID";
                    appointmentFoldersComboBox.EndUpdate();
                }

                appointmentGoogleFoldersComboBox.SelectedValue = (regKeyValueStr);
            }
        }

        private void SaveSettings()
        {
            SaveSettings(cmbSyncProfile.Text);
        }

        private void SaveSettings(string profile)
        {
            if (!string.IsNullOrEmpty(profile))
            {
                SyncProfile = cmbSyncProfile.Text;
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + profile);
                regKeyAppRoot.SetValue(RegistrySyncOption, (int)syncOption);

                if (!string.IsNullOrEmpty(UserName.Text))
                {
                    regKeyAppRoot.SetValue(RegistryUsername, UserName.Text);
                }
                regKeyAppRoot.SetValue(RegistryAutoSync, autoSyncCheckBox.Checked.ToString());
                regKeyAppRoot.SetValue(RegistryAutoSyncInterval, autoSyncInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistryAutoStart, runAtStartupCheckBox.Checked);
                regKeyAppRoot.SetValue(RegistryReportSyncResult, reportSyncResultCheckBox.Checked);
                regKeyAppRoot.SetValue(RegistrySyncDeletion, btSyncDelete.Checked);
                regKeyAppRoot.SetValue(RegistryPromptDeletion, btPromptDelete.Checked);
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInPast, pastMonthInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsMonthsInFuture, futureMonthInterval.Value.ToString());
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsTimezone, appointmentTimezonesComboBox.Text);
                regKeyAppRoot.SetValue(RegistrySyncAppointments, btSyncAppointments.Checked);
                regKeyAppRoot.SetValue(RegistrySyncAppointmentsForceRTF, btSyncAppointmentsForceRTF.Checked);
                regKeyAppRoot.SetValue(RegistrySyncContacts, btSyncContacts.Checked);
                regKeyAppRoot.SetValue(RegistrySyncContactsForceRTF, btSyncContactsForceRTF.Checked);
                regKeyAppRoot.SetValue(RegistryUseFileAs, chkUseFileAs.Checked);
                regKeyAppRoot.SetValue(RegistryLastSync, lastSync.Ticks);

                _proxy.SaveSettings(cmbSyncProfile.Text);
            }
        }

        private bool ValidSyncFolders
        {
            get
            {
                bool syncContactFolderIsValid = (contactFoldersComboBox.SelectedIndex >= 1 && contactFoldersComboBox.SelectedIndex < contactFoldersComboBox.Items.Count)
                                                || !btSyncContacts.Checked;

                bool syncAppointmentFolderIsValid = (appointmentFoldersComboBox.SelectedIndex >= 1 && appointmentFoldersComboBox.SelectedIndex < appointmentFoldersComboBox.Items.Count)
                        && (appointmentGoogleFoldersComboBox.SelectedIndex == appointmentGoogleFoldersComboBox.Items.Count - 1 || appointmentGoogleFoldersComboBox.SelectedIndex >= 1 && appointmentGoogleFoldersComboBox.SelectedIndex < appointmentGoogleFoldersComboBox.Items.Count)
                                                || !btSyncAppointments.Checked;

                //ToDo: Coloring doesn'T Work for these combos
                //setBgColor(contactFoldersComboBox, syncContactFolderIsValid);
                //setBgColor(noteFoldersComboBox, syncNoteFolderIsValid);
                //setBgColor(appointmentFoldersComboBox, syncAppointmentFolderIsValid);

                return syncContactFolderIsValid && syncAppointmentFolderIsValid;
            }
        }

        private bool ValidSyncContactFolders
        {
            get
            {
                return (contactFoldersComboBox.SelectedIndex >= 1 && contactFoldersComboBox.SelectedIndex < contactFoldersComboBox.Items.Count)
                                                || !btSyncContacts.Checked;
            }
        }

        private bool ValidSyncAppointmentFolders
        {
            get
            {
                return (appointmentFoldersComboBox.SelectedIndex >= 1 && appointmentFoldersComboBox.SelectedIndex < appointmentFoldersComboBox.Items.Count)
                        && (appointmentGoogleFoldersComboBox.SelectedIndex == appointmentGoogleFoldersComboBox.Items.Count - 1 || appointmentGoogleFoldersComboBox.SelectedIndex >= 1 && appointmentGoogleFoldersComboBox.SelectedIndex < appointmentGoogleFoldersComboBox.Items.Count)
                                                || !btSyncAppointments.Checked;
            }
        }

        private bool ValidCredentials
        {
            get
            {
                bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
                bool syncProfileIsValid = (cmbSyncProfile.SelectedIndex > 0 && cmbSyncProfile.SelectedIndex < cmbSyncProfile.Items.Count - 1);


                setBgColor(UserName, userNameIsValid);
                setBgColor(cmbSyncProfile, syncProfileIsValid);

                if (!userNameIsValid)
                    toolTip.SetToolTip(UserName, "User is of wrong format, should be full Google Mail address, e.g. user@googelmail.com");
                else
                    toolTip.SetToolTip(UserName, string.Empty);

                return userNameIsValid &&

                       syncProfileIsValid;
            }
        }

        private static void setBgColor(Control box, bool isValid)
        {
            if (!isValid)
                box.BackColor = Color.LightPink;
            else
                box.BackColor = Color.LightGreen;
        }

        private void syncButton_Click(object sender, EventArgs e)
        {
            Sync();
        }

        private void Sync()
        {
            try
            {
                if (!ValidCredentials)
                    //return;
                    throw new Exception("E-Mail address and or Sync-Profile is incomplete or incorrect - Maybe a typo or no selection...");

                fillSyncFolderItems();

                if (!ValidSyncContactFolders)
                {
                    Logger.Log(@"contactFoldersComboBox.SelectedIndex: " + contactFoldersComboBox.SelectedIndex, EventType.Debug);
                    Logger.Log(@"contactFoldersComboBox.Items.Count: " + contactFoldersComboBox.Items.Count, EventType.Debug);
                    throw new Exception("At least one Outlook contact folder is not selected or invalid!");
                }

                if (!ValidSyncAppointmentFolders)
                {
                    Logger.Log(@"appointmentFoldersComboBox.SelectedIndex: " + appointmentFoldersComboBox.SelectedIndex, EventType.Debug);
                    Logger.Log(@"appointmentFoldersComboBox.Items.Count: " + appointmentFoldersComboBox.Items.Count, EventType.Debug);
                    Logger.Log(@"appointmentGoogleFoldersComboBox.SelectedIndex: " + appointmentGoogleFoldersComboBox.SelectedIndex, EventType.Debug);
                    Logger.Log(@"appointmentGoogleFoldersComboBox.Items.Count: " + appointmentGoogleFoldersComboBox.Items.Count, EventType.Debug);
                    throw new Exception("At least one Outlook appointment folder is not selected or invalid!");
                }

                //IconTimerSwitch(true);
                ThreadStart starter = new ThreadStart(Sync_ThreadStarter);
                syncThread = new Thread(starter);
                syncThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                syncThread.CurrentUICulture = new CultureInfo("en-US");
                syncThread.Start();

                //if new version on sourceforge.net website than print an information to the log
                CheckVersion();

                // wait for thread to start
                for (int i = 0; !syncThread.IsAlive && i < 10; i++)
                    Thread.Sleep(1000);//DoNothing, until the thread was started, but only wait maximum 10 seconds
            }
            catch (Exception ex)
            {
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(ex);
            }
        }

        [STAThread]
        private async void Sync_ThreadStarter()
        {
            //==>Instead of lock, use Interlocked to exit the code, if already another thread is calling the same
            bool won = false;

            try
            {
                won = Interlocked.CompareExchange(ref executing, 1, 0) == 0;
                if (won)
                {
                    TimerSwitch(false);

                    //if the contacts folder has changed ==> Reset matches (to not delete contacts on the one or other side)                
                    RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + SyncProfile);
                    string oldSyncContactsFolder = regKeyAppRoot.GetValue(RegistrySyncContactsFolder) as string;
                    string oldSyncAppointmentsFolder = regKeyAppRoot.GetValue(RegistrySyncAppointmentsFolder) as string;
                    string oldSyncAppointmentsGoogleFolder = regKeyAppRoot.GetValue(RegistrySyncAppointmentsGoogleFolder) as string;

                    //only reset contacts if ContactsFolder changed
                    //and only reset appointments, if either OutlookAppointmentsFolder changed (without changing Google at the same time) or GoogleAppointmentsFolder changed (without changing Outlook at the same time) (not chosen before means not changed)
                    bool syncContacts = !string.IsNullOrEmpty(oldSyncContactsFolder) && !oldSyncContactsFolder.Equals(syncContactsFolder) && btSyncContacts.Checked;
                    bool syncAppointments = !string.IsNullOrEmpty(oldSyncAppointmentsFolder) && !oldSyncAppointmentsFolder.Equals(syncAppointmentsFolder) && btSyncAppointments.Checked;
                    bool syncGoogleAppointments = !string.IsNullOrEmpty(syncAppointmentsGoogleFolder) && !syncAppointmentsGoogleFolder.Equals(oldSyncAppointmentsGoogleFolder) && btSyncAppointments.Checked;
                    if (syncContacts || syncAppointments && !syncGoogleAppointments || !syncAppointments && syncGoogleAppointments)
                    {
                        bool r = await ResetMatches(syncContacts, syncAppointments);
                        if (!r)
                            throw new Exception("Reset required but cancelled by user");
                    }

                    //Then save the Contacts Folders used at last sync
                    if (btSyncContacts.Checked)
                        regKeyAppRoot.SetValue(RegistrySyncContactsFolder, syncContactsFolder);

                    if (btSyncAppointments.Checked)
                    {
                        regKeyAppRoot.SetValue(RegistrySyncAppointmentsFolder, syncAppointmentsFolder);
                        if (string.IsNullOrEmpty(syncAppointmentsGoogleFolder) && !string.IsNullOrEmpty(oldSyncAppointmentsGoogleFolder))
                            syncAppointmentsGoogleFolder = oldSyncAppointmentsGoogleFolder;
                        if (!string.IsNullOrEmpty(syncAppointmentsGoogleFolder))
                            regKeyAppRoot.SetValue(RegistrySyncAppointmentsGoogleFolder, syncAppointmentsGoogleFolder);
                    }

                    SetLastSyncText("Syncing...");
                    notifyIcon.Text = Application.ProductName + "\nSyncing...";
                    //System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
                    //notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));                    
                    IconTimerSwitch(true);

                    SetFormEnabled(false);

                    if (sync == null)
                    {
                        sync = new Synchronizer();
                        sync.contactsSynchronizer.DuplicatesFound += new ContactsSynchronizer.DuplicatesFoundHandler(OnDuplicatesFound);
                        sync.appointmentsSynchronizer.ErrorEncountered += new AppointmentsSynchronizer.ErrorNotificationHandler(OnErrorEncountered);
                        sync.contactsSynchronizer.ErrorEncountered += new ContactsSynchronizer.ErrorNotificationHandler(OnErrorEncountered);
                        sync.appointmentsSynchronizer.TimeZoneChanges += new AppointmentsSynchronizer.TimeZoneNotificationHandler(OnTimeZoneChanges);
                    }

                    Logger.ClearLog();
                    SetSyncConsoleText("");
                    Logger.Log("Sync started (" + SyncProfile + ").", EventType.Information);
                    //SetSyncConsoleText(Logger.GetText());
                    sync.SyncProfile = SyncProfile;
                    ContactsSynchronizer.SyncContactsFolder = syncContactsFolder;
                    AppointmentsSynchronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
                    AppointmentsSynchronizer.SyncAppointmentsGoogleFolder = syncAppointmentsGoogleFolder;
                    AppointmentsSynchronizer.TimeMin = DateTime.Now.AddMonths(-Convert.ToUInt16(pastMonthInterval.Value));
                    AppointmentsSynchronizer.TimeMax = DateTime.Now.AddMonths(Convert.ToUInt16(futureMonthInterval.Value));
                    AppointmentsSynchronizer.Timezone = Timezone;

                    sync.SyncOption = syncOption;
                    sync.SyncDelete = btSyncDelete.Checked;
                    sync.PromptDelete = btPromptDelete.Checked && btSyncDelete.Checked;
                    sync.contactsSynchronizer.UseFileAs = chkUseFileAs.Checked;
                    sync.SyncContacts = btSyncContacts.Checked;
                    sync.SyncAppointments = btSyncAppointments.Checked;
                    AppointmentsSynchronizer.SyncAppointmentsForceRTF = btSyncAppointmentsForceRTF.Checked;
                    ContactsSynchronizer.SyncContactsForceRTF = btSyncContactsForceRTF.Checked;

                    if (!sync.SyncContacts && !sync.SyncAppointments)
                    {
                        SetLastSyncText("Sync failed.");
                        notifyIcon.Text = Application.ProductName + "\nSync failed";

                        string messageText = "Neither contacts nor appointments are switched on for syncing. Please choose at least one option. Sync aborted!";
                        //    Logger.Log(messageText, EventType.Error);
                        //    ShowForm();
                        //    ShowBalloonToolTip("Error", messageText, ToolTipIcon.Error, 5000, true);
                        //    return;
                        //}

                        //if (sync.SyncAppointments && Syncronizer.Timezone == "")
                        //{
                        //    string messageText = "Please set your timezone before syncing your appointments! Sync aborted!";
                        Logger.Log(messageText, EventType.Error);
                        ShowForm();
                        ShowBalloonToolTip("Error", messageText, ToolTipIcon.Error, 5000, true);
                        return;
                    }

                    sync.LoginToGoogle(UserName.Text);
                    sync.LoginToOutlook();

                    sync.Sync();

                    lastSync = DateTime.Now;
                    SetLastSyncText("Last synced at " + lastSync.ToString());

                    var message = string.Format("Sync complete. Synced: {1} out of {0}. Deleted: {2}. Skipped: {3}. Errors: {4}.", sync.TotalCount, sync.SyncedCount + sync.appointmentsSynchronizer.SyncedCount + sync.contactsSynchronizer.SyncedCount, sync.DeletedCount + sync.appointmentsSynchronizer.DeletedCount + sync.contactsSynchronizer.DeletedCount, sync.SkippedCount + sync.appointmentsSynchronizer.SkippedCount + sync.contactsSynchronizer.SkippedCount, sync.ErrorCount + sync.appointmentsSynchronizer.ErrorCount + sync.contactsSynchronizer.ErrorCount);
                    Logger.Log(message, EventType.Information);

                    if (reportSyncResultCheckBox.Checked)
                    {
                        /*
                        notifyIcon.BalloonTipTitle = Application.ProductName;
                        notifyIcon.BalloonTipText = string.Format("{0}. {1}", DateTime.Now, message);
                        */
                        ToolTipIcon icon;
                        if (sync.ErrorCount + sync.appointmentsSynchronizer.ErrorCount + sync.contactsSynchronizer.ErrorCount > 0)
                            icon = ToolTipIcon.Error;
                        else if (sync.SkippedCount + sync.appointmentsSynchronizer.SkippedCount + sync.contactsSynchronizer.SkippedCount > 0)
                            icon = ToolTipIcon.Warning;
                        else
                            icon = ToolTipIcon.Info;
                        /*notifyIcon.ShowBalloonTip(5000);
                        */
                        ShowBalloonToolTip(Application.ProductName,
                            string.Format("{0}. {1}", DateTime.Now, message),
                            icon,
                            5000, false);

                    }
                    string toolTip = string.Format("{0}\nLast sync: {1}", Application.ProductName, DateTime.Now.ToString("dd.MM. HH:mm"));
                    if (sync.ErrorCount + sync.appointmentsSynchronizer.ErrorCount + sync.contactsSynchronizer.ErrorCount + sync.SkippedCount + sync.appointmentsSynchronizer.SkippedCount + sync.contactsSynchronizer.SkippedCount > 0)
                        toolTip += string.Format("\nWarnings: {0}.", sync.ErrorCount + sync.appointmentsSynchronizer.ErrorCount + sync.contactsSynchronizer.ErrorCount + sync.SkippedCount + sync.appointmentsSynchronizer.SkippedCount + sync.contactsSynchronizer.SkippedCount);
                    if (toolTip.Length >= 64)
                        toolTip = toolTip.Substring(0, 63);
                    notifyIcon.Text = toolTip;
                }
            }
            catch (Google.GData.Client.GDataRequestException ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                //string responseString = (null != ex.InnerException) ? ex.ResponseString : ex.Message;

                if (ex.InnerException is System.Net.WebException)
                {
                    string message = "Cannot connect to Google, please check for available internet connection and proxy settings if applicable: " + ex.InnerException.Message + "\r\n" + ex.ResponseString;
                    Logger.Log(message, EventType.Warning);
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }
            catch (Google.GData.Client.InvalidCredentialsException)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                string message = "The credentials (Google Account username and/or password) are invalid, please correct them in the settings form before you sync again";
                Logger.Log(message, EventType.Error);
                ShowForm();
                ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
            }
            catch (Exception ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";
                Logger.Log(ex, EventType.Debug);
                if (ex is COMException)
                {
                    string message = "Outlook exception, please assure that Outlook is running and not closed when syncing";
                    Logger.Log(message + ": " + ex.Message + "\r\n" + ex.StackTrace, EventType.Warning);
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }
            finally
            {
                if (won)
                {
                    Interlocked.Exchange(ref executing, 0);
                    lastSync = DateTime.Now;
                    TimerSwitch(true);
                    SetFormEnabled(true);
                    if (sync != null)
                    {
                        sync.LogoffOutlook();
                        sync.LogoffGoogle();
                        sync = null;
                    }
                    IconTimerSwitch(false);
                }
            }
        }

        public void ShowBalloonToolTip(string title, string message, ToolTipIcon icon, int timeout, bool error)
        {
            //if user is active on workstation
            if (boolShowBalloonTip)
            {
                notifyIcon.BalloonTipTitle = title;
                notifyIcon.BalloonTipText = message;
                notifyIcon.BalloonTipIcon = icon;
                notifyIcon.ShowBalloonTip(timeout);
            }

            string iconText = title + ": " + message;
            if (!string.IsNullOrEmpty(iconText))
                notifyIcon.Text = (iconText).Substring(0, iconText.Length >= 63 ? 63 : iconText.Length);

            if (error)
                notifyIcon.Icon = IconError;
        }

        void Logger_LogUpdated(string Message)
        {
            AppendSyncConsoleText(Message);
        }

        void OnErrorEncountered(string title, Exception ex, EventType eventType)
        {
            // do not show ErrorHandler, as there may be multiple exceptions that would nag the user
            Logger.Log(ex.ToString(), EventType.Error);
            string message = string.Format("Error Saving Contact: {0}.\nPlease report complete ErrorMessage from Log to the Tracker\nat https://sourceforge.net/tracker/?group_id=369321", ex.Message);
            ShowBalloonToolTip(title, message, ToolTipIcon.Error, 5000, true);
            /*notifyIcon.BalloonTipTitle = title;
            notifyIcon.BalloonTipText = message;
            notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
            notifyIcon.ShowBalloonTip(5000);*/
        }

        void OnTimeZoneChanges(string timeZone)
        {
            if (appointmentTimezonesComboBox.InvokeRequired)
            {
                OnTimeZoneChangesCallback d = new OnTimeZoneChangesCallback(OnTimeZoneChanges);
                Invoke(d, new object[] { timeZone });
            }
            else
            {
                appointmentTimezonesComboBox.Text = timeZone;
            }
            Timezone = timeZone;
            AppointmentsSynchronizer.Timezone = timeZone;
        }

        void OnDuplicatesFound(string title, string message)
        {
            Logger.Log(message, EventType.Warning);
            ShowBalloonToolTip(title, message, ToolTipIcon.Warning, 5000, true);
            /*
			notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
			notifyIcon.ShowBalloonTip(5000);
             */
        }

        void OnNotificationReceived(string message)
        {
            SetLastSyncText(message);
        }

        public void SetFormEnabled(bool enabled)
        {
            if (InvokeRequired)
            {
                SwitchHandler h = new SwitchHandler(SetFormEnabled);
                Invoke(h, new object[] { enabled });
            }
            else
            {
                resetMatchesLinkLabel.Enabled = enabled;
                settingsGroupBox.Enabled = enabled;
                syncButton.Enabled = enabled;
                cancelButton.Enabled = !enabled;
            }
        }
        public void SetLastSyncText(string text)
        {
            if (InvokeRequired)
            {
                TextHandler h = new TextHandler(SetLastSyncText);
                Invoke(h, new object[] { text });
            }
            else
            {
                lastSyncLabel.Text = text;
            }
        }

        public void SetSyncConsoleText(string text)
        {
            if (InvokeRequired)
            {
                TextHandler h = new TextHandler(SetSyncConsoleText);
                Invoke(h, new object[] { text });
            }
            else
            {
                syncConsole.Text = text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }

        }
        public void AppendSyncConsoleText(string text)
        {
            if (InvokeRequired)
            {
                TextHandler h = new TextHandler(AppendSyncConsoleText);
                Invoke(h, new object[] { text });
            }
            else
            {
                syncConsole.Text += text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }
        }
        public void TimerSwitch(bool value)
        {
            if (InvokeRequired)
            {
                SwitchHandler h = new SwitchHandler(TimerSwitch);
                Invoke(h, new object[] { value });
            }
            else
            {
                //If PC resumes or unlocks or is started, give him 5 minutes to recover everything before the sync starts
                if (lastSync <= DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0))
                    lastSync = DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
                autoSyncInterval.Enabled = autoSyncCheckBox.Checked && value;
                syncTimer.Enabled = autoSyncCheckBox.Checked && value;
                nextSyncLabel.Visible = autoSyncCheckBox.Checked && value;
            }
        }

        protected override void WndProc(ref Message m)
        {
            //Logger.Log(m.Msg, EventType.Information);
            switch (m.Msg)
            {
                //System shutdown
                case NativeMethods.WM_QUERYENDSESSION:
                    requestClose = true;
                    break;
                /*case NativeMethods.WM_WTSSESSION_CHANGE:
                    {
                        int value = m.WParam.ToInt32();
                        //User Session locked
                        if (value == NativeMethods.WTS_SESSION_LOCK)
                        {
                            Console.WriteLine("Session Lock",EventType.Information);
                            //OnSessionLock();
                            boolShowBalloonTip = false; // Do something when locked
                        }
                        //User Session unlocked
                        else if (value == NativeMethods.WTS_SESSION_UNLOCK)
                        {
                            Console.WriteLine("Session Unlock", EventType.Information);
                            //OnSessionUnlock();
                            boolShowBalloonTip = true; // Do something when unlocked
                            TimerSwitch(true);
                        }
                     break;
                    }
                
                
                case NativeMethods.WM_POWERBROADCAST:
                    {
                        if (m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMEAUTOMATIC ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMECRITICAL ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMESTANDBY ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMESUSPEND ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBYFAILED ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBYFAILED)
                        {                            
                            TimerSwitch(true);
                        }
                        else if (m.WParam.ToInt32() == NativeMethods.PBT_APMSUSPEND ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMSTANDBY ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBY ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSUSPEND)
                        {
                            TimerSwitch(false);
                        }
                            

                        break;
                    }*/
                default:
                    break;
            }
            //Show Window from Tray
            if (m.Msg == NativeMethods.WM_GCSM_SHOWME)
                ShowForm();
            base.WndProc(ref m);
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!requestClose)
            {
                SaveSettings();
                e.Cancel = true;
            }
            HideForm();
        }

        private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                cancellationTokenSource.Cancel();

                if (sync != null)
                    sync.LogoffOutlook();

                Logger.Log("Closed application.", EventType.Information);
                Logger.Close();

                SaveSettings();

                //unregister event handler
                SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
                SystemEvents.PowerModeChanged -= SystemEvents_PowerModeSwitch;

                notifyIcon.Dispose();
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }
        }

        private void syncOptionBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Application.DoEvents();
                int index = syncOptionBox.SelectedIndex;
                if (index == -1)
                    return;

                SetSyncOption(index);
            }
            catch (Exception ex)
            {
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(ex);
            }
        }
        private void SetSyncOption(int index)
        {
            syncOption = (SyncOption)index;
            for (int i = 0; i < syncOptionBox.Items.Count; i++)
            {
                if (i == index)
                    syncOptionBox.SetItemCheckState(i, CheckState.Checked);
                else
                    syncOptionBox.SetItemCheckState(i, CheckState.Unchecked);
            }
        }

        private void SettingsForm_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
                Hide();
        }

        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ShowForm();
        }

        private void autoSyncCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            lastSync = DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            TimerSwitch(true);
        }

        private void syncTimer_Tick(object sender, EventArgs e)
        {
            TimeSpan syncTime = DateTime.Now - lastSync;
            TimeSpan limit = new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            if (syncTime < limit)
            {
                TimeSpan diff = limit - syncTime;
                string str = "Next sync in";
                if (diff.Hours != 0)
                    str += " " + diff.Hours + " h";
                if (diff.Minutes != 0 || diff.Hours != 0)
                    str += " " + diff.Minutes + " min";
                if (diff.Seconds != 0)
                    str += " " + diff.Seconds + " s";
                nextSyncLabel.Text = str;
            }
            else
            {
                Sync();
            }
        }

        private async void resetMatchesLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // force deactivation to show up
            Application.DoEvents();
            try
            {
                cancelButton.Enabled = false; //Cancel is only working for sync currently, not for reset
                await ResetMatches(btSyncContacts.Checked, btSyncAppointments.Checked);//ToDo: Google.Documents API Replaced by Google.Drive API on 21-Apr-2015
            }
            catch (Exception ex)
            {
                SetLastSyncText("Reset Matches failed");
                Logger.Log("Reset Matches failed", EventType.Error);
                ErrorHandler.Handle(ex);
            }
            finally
            {
                lastSync = DateTime.Now;
                TimerSwitch(true);
                SetFormEnabled(true);
                hideButton.Enabled = true;
                if (sync != null)
                {
                    sync.LogoffOutlook();
                    sync.LogoffGoogle();
                    sync = null;
                }
            }
        }

        private async Task<bool> ResetMatches(bool syncContacts, bool syncAppointments)
        {
            TimerSwitch(false);

            SetLastSyncText("Resetting matches...");
            notifyIcon.Text = Application.ProductName + "\nResetting matches...";

            SetFormEnabled(false);

            if (sync == null)
            {
                sync = new Synchronizer();
            }

            Logger.ClearLog();
            SetSyncConsoleText("");
            Logger.Log("Reset Matches started  (" + SyncProfile + ").", EventType.Information);

            sync.SyncContacts = syncContacts;
            sync.SyncAppointments = syncAppointments;

            ContactsSynchronizer.SyncContactsFolder = syncContactsFolder;
            AppointmentsSynchronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
            AppointmentsSynchronizer.SyncAppointmentsGoogleFolder = syncAppointmentsGoogleFolder;
            sync.SyncProfile = SyncProfile;

            sync.LoginToGoogle(UserName.Text);
            sync.LoginToOutlook();

            if (sync.SyncAppointments)
            {
                bool deleteOutlookAppointments = false;
                bool deleteGoogleAppointments = false;

                switch (ShowDialog("Do you want to delete all Outlook Calendar entries?"))
                {
                    case DialogResult.Yes: deleteOutlookAppointments = true; break;
                    case DialogResult.No: deleteOutlookAppointments = false; break;
                    default: return false;
                }
                switch (ShowDialog("Do you want to delete all Google Calendar entries?"))
                {
                    case DialogResult.Yes: deleteGoogleAppointments = true; break;
                    case DialogResult.No: deleteGoogleAppointments = false; break;
                    default: return false;
                }

                Logger.Log("Resetting Google appointment matches...", EventType.Information);
                try
                {
                    await sync.appointmentsSynchronizer.ResetGoogleAppointmentMatches(deleteGoogleAppointments, cancellationTokenSource.Token);
                    sync.appointmentsSynchronizer.LoadAppointments();
                    sync.appointmentsSynchronizer.ResetOutlookAppointmentMatches(deleteOutlookAppointments);
                }
                catch (TaskCanceledException)
                {
                    Logger.Log("Task cancelled by user.", EventType.Information);
                    sync.appointmentsSynchronizer.LoadAppointments();
                }
            }

            if (sync.SyncContacts)
            {
                sync.contactsSynchronizer.LoadContacts();
                sync.contactsSynchronizer.ResetContactMatches();
            }

            lastSync = DateTime.Now;
            SetLastSyncText("Matches reset at " + lastSync.ToString());
            Logger.Log("Matches reset.", EventType.Information);

            return true;
        }

        public delegate DialogResult InvokeConflict(ConflictResolverForm conflictResolverForm);

        public DialogResult ShowConflictDialog(ConflictResolverForm conflictResolverForm)
        {
            if (InvokeRequired)
            {
                return (DialogResult)Invoke(new InvokeConflict(ShowConflictDialog), new object[] { conflictResolverForm });
            }
            else
            {
                DialogResult res = conflictResolverForm.ShowDialog(this);
                notifyIcon.Icon = Icon0;
                return res;
            }
        }

        public delegate DialogResult InvokeDeleteTooManyPropertiesForm(DeleteTooManyPropertiesForm form);

        public DialogResult ShowDeleteTooManyPropertiesForm(DeleteTooManyPropertiesForm form)
        {
            if (InvokeRequired)
            {
                return (DialogResult)Invoke(new InvokeDeleteTooManyPropertiesForm(ShowDeleteTooManyPropertiesForm), new object[] { form });
            }
            else
            {
                return form.ShowDialog(this);
            }
        }

        public delegate DialogResult InvokeDeleteTooBigPropertiesForm(DeleteTooBigPropertiesForm form);

        public DialogResult ShowDeleteTooBigPropertiesForm(DeleteTooBigPropertiesForm form)
        {
            if (InvokeRequired)
            {
                return (DialogResult)Invoke(new InvokeDeleteTooBigPropertiesForm(ShowDeleteTooBigPropertiesForm), new object[] { form });
            }
            else
            {
                return form.ShowDialog(this);
            }
        }

        public delegate DialogResult InvokeDeleteDuplicatedPropertiesForm(DeleteDuplicatedPropertiesForm form);

        public DialogResult ShowDeleteDuplicatedPropertiesForm(DeleteDuplicatedPropertiesForm form)
        {
            if (InvokeRequired)
            {
                return (DialogResult)Invoke(new InvokeDeleteDuplicatedPropertiesForm(ShowDeleteDuplicatedPropertiesForm), new object[] { form });
            }
            else
            {
                return form.ShowDialog(this);
            }
        }

        private delegate void InvokeCallback();

        private void ShowForm()
        {
            if (InvokeRequired)
            {
                Invoke(new InvokeCallback(ShowForm));
            }
            else
            {
                FormWindowState oldState = WindowState;

                Show();
                Activate();
                WindowState = FormWindowState.Normal;

                using (var filter = new OleMessageFilter())
                {
                    fillSyncFolderItems();
                }

                if (oldState != WindowState)
                    CheckVersion();
            }
        }

        private async void CheckVersion()
        {
            if (!NewVersionLinkLabel.Visible)
            {//Only check once, if new version is available

                try
                {
                    Cursor = Cursors.WaitCursor;
                    SuspendLayout();
                    //check for new version
                    if (NewVersionLinkLabel.LinkColor != Color.Red && await VersionInformation.isNewVersionAvailable(cancellationTokenSource.Token))
                    {
                        NewVersionLinkLabel.Visible = true;
                        NewVersionLinkLabel.LinkColor = Color.Red;
                        NewVersionLinkLabel.Text = "New Version of GCSM available on sf.net!";
                        notifyIcon.BalloonTipClicked += notifyIcon_BalloonTipClickedDownloadNewVersion;
                        ShowBalloonToolTip("New version available", "Click here to download", ToolTipIcon.Info, 20000, false);
                    }
                    NewVersionLinkLabel.Visible = true;
                }
                finally
                {
                    Cursor = Cursors.Default;
                    ResumeLayout();
                }
            }
        }

        private void notifyIcon_BalloonTipClickedDownloadNewVersion(object sender, System.EventArgs e)
        {
            Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
            notifyIcon.BalloonTipClicked -= notifyIcon_BalloonTipClickedDownloadNewVersion;
        }

        private void HideForm()
        {
            WindowState = FormWindowState.Minimized;
            Hide();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm();
            Activate();
        }
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            HideForm();
        }
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            requestClose = true;
            Close();
        }
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            using (AboutBox about = new AboutBox())
            {
                about.Show();
            }
        }
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Sync();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(UserName.Text) ||
                string.IsNullOrEmpty(cmbSyncProfile.Text) /*||
                string.IsNullOrEmpty(contactFoldersComboBox.Text)*/ )
            {
                // this is the first load, show form
                ShowForm();
                UserName.Focus();
                ShowBalloonToolTip(Application.ProductName,
                        "Application started and visible in your PC's system tray, click on this balloon or the icon below to open the settings form and enter your Google credentials there.",
                        ToolTipIcon.Info,
                        5000, false);
            }
            else
                HideForm();
        }

        private void runAtStartupCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            string regKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
            try
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(regKey);
                if (runAtStartupCheckBox.Checked)
                {
                    // add to registry
                    regKeyAppRoot.SetValue("GoogleContactSync", "\"" + Application.ExecutablePath + "\"");
                }
                else
                {
                    // remove from registry
                    regKeyAppRoot.DeleteValue("GoogleContactSync");
                }
            }
            catch (Exception ex)
            {
                //if we can't write to that key, disable it... 
                runAtStartupCheckBox.Checked = false;
                runAtStartupCheckBox.Enabled = false;
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(new Exception(("Error saving 'Run program at startup' settings into Registry key '" + regKey + "' Error: " + ex.Message), ex));
            }
        }

        private void UserName_TextChanged(object sender, EventArgs e)
        {
            ValidateSyncButton();
        }

        private void ValidateSyncButton()
        {
            syncButton.Enabled = ValidCredentials && ValidSyncFolders;
        }

        private void Donate_Click(object sender, EventArgs e)
        {
            Process.Start("https://sourceforge.net/project/project_donations.php?group_id=369321");
        }

        private void Donate_MouseEnter(object sender, EventArgs e)
        {
            Donate.BackColor = Color.LightGray;
        }

        private void Donate_MouseLeave(object sender, EventArgs e)
        {
            Donate.BackColor = Color.Transparent;
        }

        private void hideButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void proxySettingsLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_proxy != null) _proxy.ShowDialog(this);
        }

        private void SettingsForm_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            ShowHelp();
        }

        private void SettingsForm_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            ShowHelp();
        }

        private static void ShowHelp()
        {
            // go to the page showing the help and howto instructions
            Process.Start("http://googlesyncmod.sourceforge.net/");
        }

        private void btSyncContacts_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on appointments for syncing now).", "No sync switched on");
                btSyncAppointments.Checked = true;
            }
            contactFoldersComboBox.Visible = btSyncContacts.Checked;
            btSyncContactsForceRTF.Visible = btSyncContacts.Checked;
        }

        private void btSyncAppointments_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on contacts for syncing now).", "No sync switched on");
                btSyncContacts.Checked = true;
            }
            appointmentFoldersComboBox.Visible = appointmentGoogleFoldersComboBox.Visible = btSyncAppointments.Checked;
            labelTimezone.Visible = labelMonthsPast.Visible = labelMonthsFuture.Visible = btSyncAppointments.Checked;
            pastMonthInterval.Visible = futureMonthInterval.Visible = appointmentTimezonesComboBox.Visible = btSyncAppointments.Checked;
            btSyncAppointmentsForceRTF.Visible = btSyncAppointments.Checked;
        }

        private void cmbSyncProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if ((0 == comboBox.SelectedIndex) || (comboBox.SelectedIndex == (comboBox.Items.Count - 1)))
            {
                using (ConfigurationManagerForm _configs = new ConfigurationManagerForm())
                {
                    if (0 == comboBox.SelectedIndex && _configs != null)
                    {
                        SyncProfile = _configs.AddProfile();
                        ClearSettings();
                    }

                    if (comboBox.SelectedIndex == (comboBox.Items.Count - 1) && _configs != null)
                        _configs.ShowDialog(this);
                }
                fillSyncProfileItems();

                comboBox.Text = SyncProfile;
                SaveSettings();
            }
            if (comboBox.SelectedIndex < 0)
                MessageBox.Show("Please select Sync Profile.", "No sync switched on");
            else
            {
                //ClearSettings();
                LoadSettings(comboBox.Text);
                SyncProfile = comboBox.Text;
            }

            ValidateSyncButton();
        }

        private void contacFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Contacts folder you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is OutlookFolder)
            {
                syncContactsFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + ((OutlookFolder)comboBox.SelectedItem).DisplayName);
            }
            else
            {
                syncContactsFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }
            ValidateSyncButton();
        }



        private void appointmentFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Appointments folder you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is OutlookFolder)
            {
                syncAppointmentsFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + ((OutlookFolder)comboBox.SelectedItem).DisplayName);
            }
            else
            {
                syncAppointmentsFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }

            ValidateSyncButton();
        }

        private void appointmentGoogleFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Google Calendar you want to sync";
            var comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex >= 0 && comboBox.SelectedIndex < comboBox.Items.Count && comboBox.SelectedItem is GoogleCalendar)
            {
                syncAppointmentsGoogleFolder = comboBox.SelectedValue.ToString();
                toolTip.SetToolTip(comboBox, message + ":\r\n" + ((GoogleCalendar)comboBox.SelectedItem).DisplayName);
            }
            else
            {
                syncAppointmentsGoogleFolder = "";
                toolTip.SetToolTip(comboBox, message);
            }

            ValidateSyncButton();
        }

        private void btSyncDelete_CheckedChanged(object sender, EventArgs e)
        {
            btPromptDelete.Visible = btSyncDelete.Checked;
            btPromptDelete.Checked = btSyncDelete.Checked;
        }

        private void pictureBoxExit_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Do you really want to exit " + Application.ProductName + "? This will also stop the service performing automatic synchronizaton in the background. If you only want to hide the settings form, use the 'Hide' Button instead.", "Exit " + Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                CancelButton_Click(sender, EventArgs.Empty); //Close running thread
                requestClose = true;
                Close();
            }
        }

        private void SystemEvents_PowerModeSwitch(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Suspend)
            {
                TimerSwitch(false);
            }
            else if (e.Mode == PowerModes.Resume)
            {
                TimerSwitch(true);
            }
        }

        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                boolShowBalloonTip = false;
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                boolShowBalloonTip = true;
                TimerSwitch(true);
            }
        }

        private void autoSyncInterval_ValueChanged(object sender, EventArgs e)
        {
            TimerSwitch(true);
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            cancellationTokenSource.Cancel();
            KillSyncThread();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand, ControlThread = true)]
        private void KillSyncThread()
        {
            if (syncThread != null && syncThread.IsAlive)
                syncThread.Abort();
        }

        #region syncing icon
        public void IconTimerSwitch(bool value)
        {
            if (InvokeRequired)
            {
                SwitchHandler h = new SwitchHandler(IconTimerSwitch);
                Invoke(h, new object[] { value });
            }
            else
            {
                if (value) //Reset Icon to default icon as starting point for the syncing icon
                    notifyIcon.Icon = Icon0;
                iconTimer.Enabled = value;
            }
        }

        private void iconTimer_Tick(object sender, EventArgs e)
        {
            showNextIcon();
        }

        private void showNextIcon()
        {
            if (InvokeRequired)
            {
                IconHandler h = new IconHandler(showNextIcon);
                Invoke(h, new object[] { });
            }
            else
                notifyIcon.Icon = GetNextIcon(notifyIcon.Icon); ;
        }

        private Icon GetNextIcon(Icon currentIcon)
        {
            if (currentIcon == IconError) //Don't change the icon anymore, once an error occurred
                return IconError;
            if (currentIcon == Icon30)
                return Icon60;
            else if (currentIcon == Icon60)
                return Icon90;
            else if (currentIcon == Icon90)
                return Icon120;
            else if (currentIcon == Icon120)
                return Icon150;
            else if (currentIcon == Icon150)
                return Icon180;
            else if (currentIcon == Icon180)
                return Icon210;
            else if (currentIcon == Icon210)
                return Icon240;
            else if (currentIcon == Icon240)
                return Icon270;
            else if (currentIcon == Icon270)
                return Icon300;
            else if (currentIcon == Icon300)
                return Icon330;
            else if (currentIcon == Icon330)
                return Icon0;
            else
                return Icon30;
        }
        #endregion

        //private void futureMonthTextBox_Validating(object sender, CancelEventArgs e)
        //{
        //    ushort value;
        //    if (!ushort.TryParse(futureMonthTextBoxOld.Text, out value))
        //    {
        //        MessageBox.Show("only positive integer numbers or 0 (i.e. all) allowed");
        //        futureMonthTextBoxOld.Text = "0";
        //        e.Cancel = true;
        //    }

        //}

        //private void pastMonthTextBox_Validating(object sender, CancelEventArgs e)
        //{
        //    ushort value;
        //    if (!ushort.TryParse(pastMonthTextBoxOld.Text, out value))
        //    {
        //        MessageBox.Show("only positive integer numbers or 0 (i.e. all) allowed");
        //        pastMonthTextBoxOld.Text = "1";
        //        e.Cancel = true;
        //    }
        //}

        private void appointmentTimezonesComboBox_TextChanged(object sender, EventArgs e)
        {
            Timezone = appointmentTimezonesComboBox.Text;
        }

        private void linkLabelRevokeAuthentication_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Logger.Log("Trying to remove Authentication...", EventType.Information);
                FileDataStore fDS = new FileDataStore(Logger.AuthFolder, true);
                fDS.ClearAsync();
                Logger.Log("Removed Authentication...", EventType.Information);
            }
            catch (Exception ex)
            {
                Logger.Log(ex.ToString(), EventType.Error);
            }
        }

        private void appointmentGoogleFoldersComboBox_Enter(object sender, EventArgs e)
        {
            if (appointmentGoogleFoldersComboBox.DataSource == null ||
                appointmentGoogleFoldersComboBox.Items.Count <= 1)
            {
                Logger.Log("Loading Google Calendars...", EventType.Information);
                ArrayList googleAppointmentFolders = new ArrayList();

                appointmentGoogleFoldersComboBox.BeginUpdate();
                //this.appointmentGoogleFoldersComboBox.DataSource = null;

                Logger.Log("Loading Google Appointments folder...", EventType.Information);
                string defaultText = "    --- Select a Google Appointment folder ---";

                if (sync == null)
                    sync = new Synchronizer();

                sync.SyncAppointments = btSyncAppointments.Checked;
                sync.LoginToGoogle(UserName.Text);
                foreach (CalendarListEntry calendar in sync.appointmentsSynchronizer.calendarList)
                {
                    googleAppointmentFolders.Add(new GoogleCalendar(calendar.Summary, calendar.Id, calendar.Primary.HasValue ? calendar.Primary.Value : false));
                }

                if (googleAppointmentFolders != null) // && googleAppointmentFolders.Count > 0)
                {
                    googleAppointmentFolders.Sort();
                    googleAppointmentFolders.Insert(0, new GoogleCalendar(defaultText, defaultText, false));
                    appointmentGoogleFoldersComboBox.DataSource = googleAppointmentFolders;
                    appointmentGoogleFoldersComboBox.DisplayMember = "DisplayName";
                    appointmentGoogleFoldersComboBox.ValueMember = "FolderID";
                }
                appointmentGoogleFoldersComboBox.EndUpdate();
                appointmentGoogleFoldersComboBox.SelectedValue = defaultText;

                //Select Default Folder per Default
                foreach (GoogleCalendar folder in appointmentGoogleFoldersComboBox.Items)
                    if (folder.IsDefaultFolder)
                    {
                        appointmentGoogleFoldersComboBox.SelectedItem = folder;
                        break;
                    }
                Logger.Log("Loaded Google Calendars.", EventType.Information);
            }
        }

        private void autoSyncInterval_Enter(object sender, EventArgs e)
        {
            syncTimer.Enabled = false;
        }

        private void autoSyncInterval_Leave(object sender, EventArgs e)
        {
            //if (autoSyncInterval.Value == null)
            //{  //ToDo: Doesn'T work, if user deleted it, the Value is kept
            //    MessageBox.Show("No empty value allowed, set to minimum value: " + autoSyncInterval.Minimum);
            //    autoSyncInterval.Value = autoSyncInterval.Minimum;
            //}
            syncTimer.Enabled = true;
        }

        private void NewVersionLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (((LinkLabel)sender).LinkColor == Color.Red)
            {
                Logger.Log("Process Start for https://sourceforge.net/projects/googlesyncmod/files/latest/download", EventType.Debug);
                Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
            }
            else
            {
                Logger.Log("Process Start for https://sourceforge.net/projects/googlesyncmod/", EventType.Debug);
                Process.Start("https://sourceforge.net/projects/googlesyncmod/");
            }
        }
    }
}