﻿using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Net;

namespace GoContactSyncMod
{
    partial class ProxySettingsForm : Form
    {
        private static IWebProxy _systemProxy = new System.Net.WebProxy();

        private void Form_Changed(object sender, EventArgs e)
        {
            FormSettings();
        }

        public ProxySettingsForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();
            if (null == _systemProxy)
                _systemProxy = WebRequest.DefaultWebProxy;
#if debug
            cbUseGlobalSettings.Visible = true;
#else
            cbUseGlobalSettings.Checked = true;
            cbUseGlobalSettings.Visible = false;
            LoadSettings(null);
#endif
        }

        private static void setBgColor(TextBox box, bool isValid)
        {
            if (box.Enabled)
            {
                if (!isValid)
                    box.BackColor = Color.LightPink;
                else
                    box.BackColor = Color.LightGreen;
            }
        }

        private bool ValidCredentials
        {
            get
            {
                bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\\\/\@\'\%\._\+\s\-]+)$", RegexOptions.IgnoreCase);
                bool passwordIsValid = !string.IsNullOrEmpty(Password.Text.Trim());
                bool AddressIsValid = Regex.IsMatch(Address.Text, @"^(?'url'[\w\d#@%;$()~_?\-\\\.&]+)$", RegexOptions.IgnoreCase);
                bool PortIsValid = Regex.IsMatch(Port.Text, @"^(?'port'[0-9]{2,6})$", RegexOptions.IgnoreCase);

                setBgColor(UserName, userNameIsValid);
                setBgColor(Password, passwordIsValid);
                setBgColor(Address, AddressIsValid);
                setBgColor(Port, PortIsValid);
                return (userNameIsValid && passwordIsValid || !Authorization.Checked) && AddressIsValid && PortIsValid || SystemProxy.Checked;
            }
        }

        private void FormSettings()
        {
            Address.Enabled = CustomProxy.Checked;
            Port.Enabled = CustomProxy.Checked;
            Authorization.Enabled = CustomProxy.Checked;
            UserName.Enabled = CustomProxy.Checked && Authorization.Checked;
            Password.Enabled = CustomProxy.Checked && Authorization.Checked;

            bool isValid = ValidCredentials;
        }

        public void ProxySet()
        {
            if (CustomProxy.Checked && !string.IsNullOrEmpty(Address.Text))
            {
                try
                {
                    System.Net.WebProxy myProxy = new System.Net.WebProxy(Address.Text);
                    if (!string.IsNullOrEmpty(Port.Text))
                        myProxy = new System.Net.WebProxy(Address.Text, Convert.ToInt16(Port.Text));
                    myProxy.BypassProxyOnLocal = true;
                    myProxy.UseDefaultCredentials = true;

                    if (Authorization.Checked)
                    {
                        myProxy.Credentials = new System.Net.NetworkCredential(UserName.Text, Password.Text);
                    }
                    WebRequest.DefaultWebProxy = myProxy;
                }
                catch (Exception ex)
                {
                    ErrorHandler.Handle(ex);
                }
            }
            else // to do set defaul system proxy
                WebRequest.DefaultWebProxy = _systemProxy;
        }

        public void LoadSettings(string _profile)
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey);

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            }

            if (null != regKeyAppRoot.GetValue("UseGlobalProxySettings"))
                cbUseGlobalSettings.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("UseGlobalProxySettings"));

            regKeyAppRoot = Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey + (_profile != null && !cbUseGlobalSettings.Checked ? ('\\' + _profile) : ""));

            //only for downside compliance reasons: load old registry settings first and save them later on in new structure
            if (Registry.CurrentUser.OpenSubKey(@"Software\Webgear\GOContactSync") != null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync" + (_profile != null && !cbUseGlobalSettings.Checked ? ('\\' + _profile) : ""));
            }

            if (regKeyAppRoot.GetValue("ProxyUsage") != null)
            {
                if (Convert.ToBoolean(regKeyAppRoot.GetValue("ProxyUsage")))
                {
                    CustomProxy.Checked = true;
                    SystemProxy.Checked = !CustomProxy.Checked;

                    if (regKeyAppRoot.GetValue("ProxyURL") != null)
                        Address.Text = (string)regKeyAppRoot.GetValue("ProxyURL");

                    if (regKeyAppRoot.GetValue("ProxyPort") != null)
                        Port.Text = (string)regKeyAppRoot.GetValue("ProxyPort");

                    if (Convert.ToBoolean(regKeyAppRoot.GetValue("ProxyAuth")))
                    {
                        Authorization.Checked = true;

                        if (regKeyAppRoot.GetValue("ProxyUsername") != null)
                        {
                            UserName.Text = regKeyAppRoot.GetValue("ProxyUsername") as string;
                            if (regKeyAppRoot.GetValue("ProxyPassword") != null)
                                Password.Text = Encryption.DecryptPassword(UserName.Text, regKeyAppRoot.GetValue("ProxyPassword") as string);
                        }
                    }
                }
            }

            FormSettings();
            ProxySet();
        }
        public void ClearSettings()
        {
            if (!cbUseGlobalSettings.Checked)
            {
                SystemProxy.Checked = true;
                CustomProxy.Checked = Authorization.Checked = !SystemProxy.Checked;
                Address.Text = Port.Text = UserName.Text = Password.Text;
            }
        }

        public void SaveSettings(string _profile)
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey);

            regKeyAppRoot.SetValue("UseGlobalProxySettings", cbUseGlobalSettings.Checked);

            regKeyAppRoot = Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey + (_profile != null && !cbUseGlobalSettings.Checked ? ('\\' + _profile) : ""));
            regKeyAppRoot.SetValue("ProxyUsage", CustomProxy.Checked);

            if (CustomProxy.Checked)
            {

                if (!string.IsNullOrEmpty(Address.Text))
                {
                    regKeyAppRoot.SetValue("ProxyURL", Address.Text);
                    if (!string.IsNullOrEmpty(Port.Text))
                        regKeyAppRoot.SetValue("ProxyPort", Port.Text);
                }

                regKeyAppRoot.SetValue("ProxyAuth", Authorization.Checked);
                if (Authorization.Checked)
                {
                    if (!string.IsNullOrEmpty(UserName.Text))
                    {
                        regKeyAppRoot.SetValue("ProxyUsername", UserName.Text);
                        if (!string.IsNullOrEmpty(Password.Text))
                            regKeyAppRoot.SetValue("ProxyPassword", Encryption.EncryptPassword(UserName.Text, Password.Text));
                    }
                }
            }
        }


        private void cancelButton_Click(object sender, EventArgs e)
        {
            if (!ValidCredentials)
                SystemProxy.Checked = true;
            Hide();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (!ValidCredentials)
                return;
#if !debug
            SaveSettings(null);
#endif
            ProxySet();
            Hide();
        }

    }
}

