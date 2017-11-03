# +++ NEWS +++ NEWS +++ NEWS +++

## Version [3.10.14] - 27.11.2016

### SVN commits

**r790 - r799:

- FIX: Do not scan for duplicates during matching, duplicates were removed during load [bug: #954]
- FIX: Fixed NullReferenceException in Office version checking
- FIX: Handle DPI resolutions [bug: #953]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.13] - 20.11.2016

### SVN commits

**r787 - r789:

- FIX: Avoid further matching in case appointments are already matched [bug: #951]

## Version [3.10.12] - 19.11.2016

### SVN commits

**r781 - r786:

- FIX: Check if value of extended property field is null [bug: #950]

## Version [3.10.11] - 18.11.2016 

### SVN commits

**r765 - r780:

- FIX: Check if UserDefinedFields field is null [bug: #948]
- FIX: Check if FormattedAddress field is null [bug: #945]
- FIX: Workaround for issue in Google client libraries [bug: #866]
- IMPROVEMENT: Do not warn about skipping empty contact if this is distribution list
- FIX: Implement OleMessageFilter to handle RPC_E_CALL_REJECTED errors [bug #939]
- FIX: Unhide label with status text [bug #942]
- FIX: Retry in case ProtocolViolationException exception during Google contact save [bug #903]
- FIX: Do not add custom field to folder [bug #651]

## Version [3.10.10] - 12.11.2016

### SVN commits

**r744 - r764:

- FIX: Preserve FileAs format when updating existing Outlook contact [bug #543]
- FIX: Change mapping between Outlook and Gmail for email types (do not use display name from Outlook) [bug #932]
- FIX: Added more logging [bugs #843, #897]
- FIX: Handle contacts with duplicated extended properties [bug #655]
- FIX: Handle contacts with too big extended properties [bug #895]
- FIX: Handle contacts with more than 10 extended properties [bug #900]
- FIX: Do not synchronize from Outlook phone numbers with only white spaces [bug #629]
- FIX: Add dummy values to contact user properties or contact extended properties [bugs #634, #886]
- UPDATE: Update NuGet packages to latest version. 

## Version [3.10.9] - 21.10.2016 

### SVN commits

**r737 - r743:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Fixed regression in selecting folders.

## Version [3.10.8] - 18.10.2016 

### SVN commits

**r727 - r736:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Add more logging for exceptions during accessing user properties [bug #651]
- FIX: Additional check to avoid access violation [bug #567]
- FIX: Fixed regression from 3.10.7, logon to MAPI using selected folder not default one.

## Version [3.10.7] - 15.10.2016 

### SVN commits

**r711 - r726:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Some users have emty time zone at Google, in such situation try to use what is set in GUI [bug #878]
- FIX: Do not throw exception in case there is problem with registry
- FIX: Layout changes for high DPI setups
- FIX: Handle situation when Outlook folder is invalid, logon to Outlook using default folders
- FIX: Handle situation when previously selected Outlook folder became invalid (for example was deleted in Outlook)

## Version [3.10.6] - 07.10.2016 

### SVN commits

**r705 - r710:

- FIX: Logon to MAPI in case of exception [bug #871]
- FIX: Select the first item in folder combo [bug #871]

## Version [3.10.5] - 06.10.2016 

### SVN commits

**r695 - r704:

- FIX: Add more logging [bugs #871, #877, #878, #879]
- FIX: Improved logging for COMExceptions [bug #871]
- FIX: In case folder was set not correctly, switch to default one [bug #871]
- FIX: Avoid exception if version information cannot be read 

## Version [3.10.4] - 04.10.2016 

### SVN commits

**r689 - r694:

- FIX: Release memory while scanning Outlook items [bug #874]
- FIX: Added more detailed logging [bug #871]
- FIX: Handle situation when bitness is not set in registry [bug #876]

## Version [3.10.3] - 03.10.2016 

### SVN commits

**r672 - r688:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Corrected time zone mapping between Google and Outlook
- FIX: Added more logging [bugs: #863, #870]
- FIX: Added AutoGenerateBindingRedirects

## Version [3.10.2] - 29.09.2016 

### SVN commits

**r660 - r671:

- UPDATE: Update NuGet packages to latest version. 
- FIX: Set time zone for recurrent appointments.
- FIX: Added more logging [bugs: #870, #871]

## Version [3.10.1] - 25.09.2016 

### SVN commits

**r610 - r659:

- UPDATE: Update NuGet packages to latest version 
- FIX: update detection routine to fetch information about the latest version [bugs: #795, #826, #845, #853]
- FIX: Ignore exceptions when retrieving windows version, put more diagnostic to log in case of exception [bug: #849]
- FIX: Remove duplicates from Outlook: two different Outlook appointments pointing to the same Google appointment. [bug: #614]
- FIX: Force Outlook to set country in formatted address string. [bug: #850]
- FIX: Clear Google reminders in case Outlook appointment has no reminders. [bug: #599]
- FIX: Synchronize time zones. [bugs: #533, #654, #813, #851, #852, #856]

## Version [3.10.0] - 14.06.2016

### SVN commits

**r586 - r609**:

- CHANGE: Retargetted to .NET 4.5,  as a result Windows XP is not supported anymore, minimum requirement is Windows Vista SP2.
- UPDATE: Update NuGet packages to latest version (new version of Google client libraries require .NET 4.5)
- FIX: ResetMatches rewritten to use BatchRequest functionality [bugs: #673, #738, #796, #799, #806, #836]
- FIX: Warning in exception handler to indicate appointment which triggered error (feature request: #148)
- FIX: Log instead of Error Handler to avoid multiple Windows
- FIX: Added more info about raised exceptions

## Version [3.9.15] - 11.03.2016

### SVN commits

**r582 - r583**:

- FIX: merged back AppointmentSync to use ForceRTF
- UPDATE: Removed GoogleDocuments 2nd level authentication, because no notes sync possible currently (no need to provide GCSM access to GoogleDocuments)

## Version [3.9.14] - 24.02.2016

### SVN commits

**r572 - r579**:

- FIX: handle busy/free/tentative status by transparency, see <https://sourceforge.net/p/googlesyncmod/bugs/463/>
- FIX: implemented ForceRTF checkbox
- UPDATE: update NuGet packages to latest version
- UPDATE: tooltip (UserName)
- FIX: Added more diagnostics for problems with Outlook installation [bugs:#785]
- FIX: changed copy to clipboard code to prevent HRESULT: CLIPBRD_E_CANT_OPEN, see [bugs:#749]
- UPDATE: field label User -> E-Mail in UI FIX: changed error typo

## Version [3.9.13] - 01.11.2015

### SVN commits

**r567 - r571**:

```
- FIX: [bugs:#780]
- UPDATE: nuget packages
- IMPROVEMENT: detect version: Outlook 2016
- IMPROVEMENT: log windows version: name, architecture, number
- IMPROVEMENT: do not copy the interop dll to output dir
- IMPROVEMENT: do not include interop into setup
- CHANGE: target type to AnyCPU
- CHANGE: remove Office.dll (not necessary)
- IMPROVEMENT: Added notes how to repair VS2013 installation after modifying machine.config for UnitTests
- prepared new setup
```

## Version [3.9.12] - 16.10.2015

### SVN commits

**r563 - r566**:

```
- Reverted change from 3.9.11: Referenced Outlook 2013 Interop API and copied it locally
turned out, also not runnable with Outlook 2016
and has issues with Older Office 2010 and 2007 installations
```

## Version [3.9.11] - 15.10.2015

### SVN commits

**r558 - r562**:

```
- FIX: Workaround, to not overwrite tentative/free Calendar items, see [bugs:#709]
- FIX: [bugs:#731]
- UPDATE: nuget packages
- FIX: don't load old registry settings to avoid profile errors
- FIX: Remove recurrence from slave, if removed from master
- FIX: Extended ListSeparator for GoogleGroups
- FIX: handle exception when saving Outlook appointment fails (log warning instead of stop and throw error)
```

## Version [3.9.10] - 16.05.2015

### SVN commits

**r555 - r557**:

```
- FIX: Remove recurrence from slave, if removed from master
- FIX: Extended ListSeparator for GoogleGroups
- FIX: handle exception when saving Outlook appointment fails (log warning instead of stop and throw error)
```

## Version [3.9.9] - 12.05.2015

### SVN commits

**r552 - r553**

```
- FIX: Improved GUI behavior, if CheckVersion fails (e.g. because of missing internet connection or wrong proxy settings)
- FIX: added America/Phoenix to the timezone Dropdown
```

## Version [3.9.8] - 04.05.2015

### SVN commits

**r546 - r550**

```
- FIX: stopped duplicating Group combinations and adding them to Google, [see](https://sourceforge.net/p/googlesyncmod/bugs/691/)
- FIX: avoid "Forbidden" error message, if calender item cannot be changed by Google account, [see](https://sourceforge.net/p/googlesyncmod/bugs/696/)
- FIX: removed debug update detection code
- UPDATE: Google.Apis.Calendar.v3
- FIX: moving "Copy to Clipboard" back to own STA-Thread
- FIX: ballon tooltip for update was always shown (svn commit error)
```

## Version [3.9.7] - 21.04.2015

### SVN commits

**r542 - r544**

```
- FIX: Removed Notes Sync, because not supported by Google anymore
- FIX: Handle null values in Registry Profiles, [see](http://sourceforge.net/p/googlesyncmod/bugs/675/)
```

**Free Open Source Software, Hell Yeah!**

## Version [3.9.6] - 15.04.2015

### SVN commits

**r536 - r541**

- **IMPROVEMENT**: adjusted error text color
- **IMPROVEMENT**: Made Timezone selection a dropdown combobox to enable users to add their own timezone, if needed (e.g. America/Arizona)
- **IMPROVEMENT**: check for latest downloadable version at sf.net
- **IMPROVEMENT**: check for update on start
- **IMPROVEMENT**: added new error dialog for user with clickable links
- **FIX**: renamed Folder OutlookAPI to MicrosoftAPI
- **FIX**: <https://sourceforge.net/p/googlesyncmod/bugs/700/>
- **CHANGE**: small fixes and changes to the Error Dialog

**Free Open Source Software, Hell Yeah!**

## Version [3.9.5] - 10.04.2015

### SVN commits

**r535**

```
- **FIX**: Fix errors when reading registry into checkbox or number textbox, see https://sourceforge.net/p/googlesyncmod/bugs/667/
https://sourceforge.net/p/googlesyncmod/bugs/695/
https://sourceforge.net/p/googlesyncmod/support-requests/354/, and others
- **FIX**: Invalid recurrence pattern for yearly events, see
https://sourceforge.net/p/googlesyncmod/support-requests/324/
https://sourceforge.net/p/googlesyncmod/support-requests/363/
https://sourceforge.net/p/googlesyncmod/support-requests/344/
- **IMPROVEMENT**: Swtiched to number textboxes for the months range
```

**Free Open Source Software, Hell Yeah!**

## Version [3.9.4] - 07.04.2015

### SVN commits

**r529 - r534**

- **FIX**: persist GoogleCalendar setting into Registry, see <https://sourceforge.net/p/googlesyncmod/bugs/685/> <https://sourceforge.net/p/googlesyncmod/bugs/684/>
- **FIX**: FIX: more spelling corrections
- **FIX**: spelling/typos corrections [bugs:#662] - UPD: nuget packages

**Free Open Source Software, Hell Yeah!**

## Version [3.9.3] - 04.04.2015

### SVN commits

**r514 - r528**

- **FIX**: fixed Google Exception when syncing appointments accepted on Google side (sent by different Organizer on Google), see <http://sourceforge.net/p/googlesyncmod/bugs/532/>
- **FIX**: not show delete conflict resoultion, if syncDelete is switched off or GoogleToOutlookOnly or OutlookToGoogleOnly
- **FIX**: fixed some issues with GoogleCalendar choice
- **FIX**: fixed some NullPointerExceptions

- **IMPROVEMENT**: Added Google Calendar Selection for appointment sync

- **IMPROVEMENT**: set culture for main-thread and SyncThread to English for english-style exception messages which are not handled by Errorhandler.cs

**Free Open Source Software, Hell Yeah!**

[3.9.3] <http://sourceforge.net/projects/googlesyncmod/files/Releases/3.9.3/SetupGCSM-3.9.3.msi/download>

## Version [3.9.2] - 27.12.2014

### SVN commits

**r511 - r513**

- **FIX**: Switched from Debugging to Release, prepared setup 3.9.2
- **FIX**: Handle AccessViolation exceptions to avoid crashes when accessing RTF Body

**Free Open Source Software, Hell Yeah!**

## Version [3.9.1] - 27.12.2014

### SVN commits

**r491 - r510**

- **FIX**: Handle Google Contact Photos wiht oAuth2 AccessToken
- **FIX**: small text changes in error dialog (added "hint message")
- **FIX**: moved client_secrets.json to Resources + added paths
- **FIX**: upgraded UnitTests and made them compilable
- **FIX**: Proxy Port was not used, because of missing exclamation mark before the null check
- **FIX**: bugfixes for Calendar sync
- **FIX**: replaced ClientLoginAuthenticator by OAuth2 Version and enabled Notes sync again
- **FIX**: removed 5 minutes minimum timespan again (doesn't make sense for 2 syncs, would make sense between changes of Outlook items, but this we cannot control
- **FIX**: Instead of deleting the registry settings, copy it from old WebGear structure ...
- **FIX**: copy error message to clipboard see [bugs:#542]

- **CHANGE**: search only .net 4.0 full profile as startup condition

- **CHANGE**: changed Auth-Class

  ```
          removed password field
          added possibility to delete user auth tokens
          changed auth folder
          changed registry settings tree name
          remove old settings-tree
  ```

- **CHANGE**: use own OAuth-Broker

  ```
          added own implementation of OAuth2-Helper class to append user (parameter: login_hint) to authorization url
          add user email to authorization uri
  ```

- **CHANGE**: removed build setting for old GoogleAPIDir

- **IMPROVEMENT**: simplified code

  ```
               rename class file - small code cleanup
  ```

- **IMPROVEMENT**: Authentication between GCSM and Google is done with OAuth2 - no password needed anymore
- **IMPROVEMENT**: changed layout and added labels for appointment fields

  ```
               set timezone before appointment sync! see [feature-requests:#112]
  ```

- **IMPROVEMENT**: setting culture for error messages to english

**Free Open Source Software, Hell Yeah!**

## Version 3.9.0

FIX: Got UnitTests running and confirmed pass results, to create setup for version 3.9.0 FIX: crash with .NET4.0 because of AccessViolationException when accessing RTFBoxy <http://sourceforge.net/p/googlesyncmod/bugs/528> FIX: Make use of Timezone settings for recurring events optional FIX: small text changes in error dialog (added "hint message") FIX: moved client_secrets.json to Resources FIX: upgraded UnitTests and made them compilable FIX: log and auth token are now written to System.Environment.SpecialFolder.ApplicationData + -NET 4.0 is now prerequisite IMPROVEMENT: added Google.Apis.Calendar.v3 and replaced v2

Contact Sync Mod, Version 3.8.6 Switched off Calandar sync, because v2 API was switched off Created last .NET 2.0 setup for version 3.8.6 (without CalendarSync

- fixed newline spelling in Error Dialog
- disable Checkbox "Run program at startup" if we can't write to hive (HKCU)
- Unload Outlook after version detection FIX: check, if Proxy settings are valid
- release outlook COM explicitly
- show Outlook Logoff in log windows
- remove old windows version detection code

Contact Sync Mod, Version 3.8.5 FIX: Handle invalid characters in syncprofile FIX: Also enable recreating appointment from Outlook to Google, if Google appointment was deleted and Outlook has multiple participants FIX: also sync 0 minutes reminder

Contact Sync Mod, Version 3.8.4 FIX: debug instead of warning, if AllDay/Start/End cannot be updated for a recurrence FIX: Don't show error messge, if appointment/contact body is the same and must not be updated

Contact Sync Mod, Version 3.8.3 Improvement: Added some info to setting errors (Google credentials and not selected folder), and added a dummy entry to the Outlook folder comboboxes to highlight, that a selection is necessary FIX: Show text, not class in Error message for recurrence FIX: Changed RTF error message to Debug FIX: Try/Catch exception when converting RTF to plain text, because some users reported memory exception since 3.8.2 release and changed error message to Debug INSTALL: added version detection for Windows 8.1 and Windows Server 2012 R2

- fixed detect of windows version
- remove "old" unmanaged calls
- use WMI to detect version

Contact Sync Mod, Version 3.8.2 IMPROVEMENT: Not overwrite RTF in Outlook contact or appointment bode FIX: recurrence exception during more than one day but not allday events are synced properly now FIX: Sensitivity can only be changed for single appointments or recurrence master

Contact Sync Mod, Version 3.8.1 FIX: sync reminder for newly created recurrence AppointmentSync IMPROVEMENT: sync private flag FIX: don't use allday property to find OriginalDate FIX: Sync deleted appointment recurrences

Contact Sync Mod, Version 3.8.0 IMPROVEMENT: Upgraded development environment from VS2010 to VS2012 and migrated setup from vdproj to wix ATTENTION: To install 3.8.0 you will have to uninstall old GCSM versions first, because the new setup (based on wix) is not compatible with the old one (based on vdproj) FIX: Save OutlookAppointment 2 times, because sometimes ComException is thrown FIX: Cleaned up some duplicate timezone entries FIX: handle Exception when permission denied for recurrences

Contact Sync Mod, Version 3.7.3 FIX: Handle error when Google contact group is not existing FIX: Handle appointments with multiple participants (ConflictResolver)

Contact Sync Mod, Version 3.7.2 FIX: don't update or delete Outlook appointments with more than 1 recipient (i.e. has been sent to participants) <https://sourceforge.net/p/googlesyncmod/support-requests/272/> FIX: Also consider changed recurrence exceptions on Google Side

Contact Sync Mod, Version 3.7.1 IMPROVEMENT: Added Timezone Combobox for Recurrent Events FIX: Fixed some pilot issues with the first appointment sync

Contact Sync Mod, Version 3.7.0 IMPROVEMENT: Added Calendar Appointments Sync

Contact Sync Mod, Version 3.6.1 FIX: Renamed automization by automation FIX: stop time, when Error is handled, to avoid multiple error message popping up

Contact Sync Mod, Version 3.6.0 IMPROVEMENT: Added icons to show syncing progress by rotating icon in notification area IMPROVEMENT: upgraded to Google Data API 2.2.0 IMPROVEMENT: linked notifyIcon.Icon to global properties' resources IMPROVEMENT: centralized all images and icon into Resources folder and replaced embedded images by link to file

Contact Sync Mod, Version 3.5.25 FIX: issue reported regarding sync folders always set to default: <https://sourceforge.net/p/googlesyncmod/bugs/436/> FIX: NullPointerException when resolving deleted GoogleNote to update again from Outlook

Contact Sync Mod, Version 3.5.24 IMPROVEMENT: Added CancelButton to cancel a running sync thread FIX: DoEvents to handle AsyncUpload of Google notes FIX: suspend timer, if user changes the time interval (to prevent running the sync instantly e.g. if removing the interval) FIX: little code cleanup FIX: add Outlook 2013 internal version number for detection FIX: removed obsolete debug-code

Contact Sync Mod, Version 3.5.23 IMPROVEMENT: Added new Icon with exclamation mark for warning/error situations FIX: show conflict in icon text and balloon, and keep conflict dialog on Top, see <http://sourceforge.net/p/googlesyncmod/support-requests/184/> FIX: Allow Outlook notes without subject (create untitled Google document) FIX: Wait up to 10 seconds until thread is alive (instead of endless loop)

Contact Sync Mod, Version 3.5.22 IMPROVEMENT: Replaced lock by Interlocked to exit sync thread if already another one is running IMPROVEMENT: fillSyncFolderItems only when needed (e.g. showing the GUI or start syncing or reset matches). IMPROVEMENT: Changed the start sync interval from 90 seconds to 5 minutes to give the PC more time to startup

Contact Sync Mod, Version 3.5.21 FIX: Fixed the issue, if Google username had an invalid character for Outlook properties <https://sourceforge.net/tracker/?func=detail&atid=1539126&aid=3598515&group_id=369321> <https://sourceforge.net/tracker/?func=detail&aid=3590035&group_id=369321&atid=1539126> FIX: Assign relationship, if no EmailDisplayName exists IMPROVEMENT: Added possibility to delete Google contact without unique property FIX: docked right splitContainer panel of ConflictResolverForm to fill full panel

Contact Sync Mod, Version 3.5.20 IMPROVEMENT: Improved INSTALL PROCESS

```
- added POSTBUILDEVENT to add version of Variable Productversion (vdproj) automatically to installer (msi) file after successful build only change the version string in the setup project and all other is done
- changed standard setup filename
```

IMPROVEMENT: added to error message to use the latest version (with url) before reporting a error to the tracker IMPROVEMENT: Added Exit-Button between hide button (Tracker ID: 3578131) FIX: Delete Google Note categories first before reassigning them (has been fixed also on Google Drive now, when updating a document, it doesn't lose the categories anymore) FIX: Updated Email Display Name

Contact Sync Mod, Version 3.5.19 IMPROVEMENT: Added Note Category sync FIX: Google Notes folder link is removed from updated note => Move note to Notes folder again after update IMPROVEMENT: added class VersionInformation (detect Outlook-Version and Operating-System-Version)

Contact Sync Mod, Version 3.5.18 FIX: added log message, if EmailDisplayName is different, because Outlook cannot set it manually FIX: switched to x86 compilation (tested with Any CPU and 64 bit, no real performance improvement), therefore x86 will be the most compatible way FIX: Preserve Email Display Name if address not changed, see also <https://sourceforge.net/tracker/index.php?func=detail&aid=3575688&group_id=369321&atid=1539129> FIX: removed Cleanup algorithm to get rid of duplicate primary phone numbers FIX: Handle unauthorized access exception when saving 'run program at startup' setting to registry, see also <https://sourceforge.net/tracker/?func=detail&aid=3560905&group_id=369321&atid=1539126> FIX: Fixed null addresses at emails

Contact Sync Mod, Version 3.5.17 FIX: applied proper tooltips to the checkboxes, see <https://sourceforge.net/tracker/?func=detail&atid=1539126&aid=3559759&group_id=369321> FIX: UI Spelling and Grammar Corrections - ID: 3559753 FIX: fixed problem when saving Google Photo, see <https://sourceforge.net/tracker/?func=detail&aid=3555588&group_id=369321&atid=1539126>

Contact Sync Mod, Version 3.5.16 FIX: fixed bug when deleting a contact on GoogleSide (Precondition failed error) FIX: fixed some typos and label sizes in ConflictResolverForm FIX: Also handle InvalidCastException when loggin into Outlook IMPROVEMENT: changed some variable declarations to var FIX: Skip empty OutlookNote to avoid Nullpointer Reference Exception FIX: fixed IM sync, not to add the address again and again, until the storage of this field exceeds on Google side FIX: fixed saving contacts and notes folder to registry, if empty before

Contact Sync Mod, Version 3.5.15 FIX: increased TimeTolerance to 120 seconds to avoid resync after ResetMatches FIX: added UseFileAs feature also for updating existing contacts IMPROVEMENT: applied "UseFileAs" setting also for syncing from Google to Outlook (to allow Outlook to choose FileAs as configured) IMPROVEMENT: replaced radiobuttons rbUseFileAs and rbUseFullName by Checkbox chkUseFileAs and moved it from bottom to the settings groupBox

Contact Sync Mod, Version 3.5.14 FIX: NullPointerException when syncing notes, see <https://sourceforge.net/tracker/index.php?func=detail&aid=3522539&group_id=369321&atid=1539126> IMPROVEMENT: Added setting to choose between Outlook FileAs and FullName

Contact Sync Mod, Version 3.5.13 IMPROVEMENT: added tooltips to Username and Password if input is wrong IMPROVEMENT: put contacts and notes folder combobox in different lines to enable resizing them Improvement: Migrated to Google Data API 2.0 Imporvement: switched to ResumableUploader for GoogleNotes FIX: Changed layer order of checkboxes to overcome hiding them, if Windows is showing a bigger font

Contact Sync Mod, Version 3.5.12 IMPROVEMENT: Implemented GUI to match Duplicates and added feature to keep both (Google and Outlook entry) FIX: Only show warning, if an OutlookFolder couldn't be opened and try to open next one

Contact Sync Mod, Version 3.5.11 FIX: Also create Outlook Contact and Note items in the selected folder (not default folder)

Contact Sync Mod, Version 3.5.10 FIX: Only check log file size, if log file size already exists

Contact Sync Mod, Version 3.5.9 IMPROVEMENT: create new logfile, once 1MB has been exceeded (move to backup before) Improvement: Added ConflictResolutions to perform selected actions for all following itmes IMPROVEMENT: Enable the user to configure multipole sync profiles, e.g. to sync with multiple gmail accounts IMPROVEMENT: Enable the user to choose Outlook folder IMPROVEMENT: Added language sync FIX: Remove Google Note directly from root folder IMPROVEMENT: No ErrorHandle when neither notes nor contacts are selected ==> Show BalloonTooltip and Form instead Improvement: Added ComException special handling for not reachable RPC, e.g. if Outlook was closed during sync Improvement: Added SwitchTimer to Unlock PC message FIX: Improved error handling, especially when invalid credentials=> show the settings form Improvement: handle Standby/Hibernate and Resume windows messages to suspend timer for 90 seconds after resume

Contact Sync Mod, Version 3.5.8 FIX: validation mask of proxy user name (by #3469442) FIX: handled OleAut Date exceptions when updating birthday IMPROVEMENT: open Settings GUI of running GCSM process when starting new instance (instead of error message, that a process is already running) FIX: validation mask of proxy uri (by #3458192) IMPROVEMENT: ResetMatch when deleting an entry (to avoid deleting it again, if restored from Outlook recycle bin)

Contact Sync Mod, Version 3.5.7 IMPROVEMENT: made OutlookApplication and Namespace static IMPROVEMENT: added balloon after first run, see <https://sourceforge.net/tracker/?func=detail&aid=3429308&group_id=369321&atid=1539126> FIX: Delete temporary note file before creating a new one FIX: Reset OutlookGoogleNoteId after note has been deleted on Google side before recreated by Upload (new GoogleNoteId) FIX: Set bypass proxy local resource in new proxy mask FIX: set for use default credentials for auth. in new proxy mask

Contact Sync Mod, Version 3.5.6 IMPROVEMENT: added proxy config mask and proxy authentication (in addition to use App.config) IMPROVEMENT: finished Notes sync feature IMPROVEMENT: Switched to new Google API 1.9 (Previous: 1.8) FIX: Added CreateOutlookInstance to OutlookNameSpace property, to avoid NullReferenceExceptions FIX: Removed characters not allowed for Outlook user property names: []_# FIX: handled exception when updating Birthday and anniversary with invalid date, see <https://sourceforge.net/tracker/?func=detail&aid=3397921&group_id=369321&atid=1>

Contact Sync Mod, Version 3.5.5 FIX: set _sync.SyncContacts properly when resetting matches (fixes <https://sourceforge.net/tracker/index.php?func=detail&aid=3403819&group_id=369321&atid=1539126>)

Contact Sync Mod, Version 3.5.4 IMPROVEMENT: added pdb file to installation to get some more information, when users report bugs IMPROVEMENT: Added also email to not require FullName IMPROVEMENT: Added company as unique property, if FullName is emptyFullName

See also Feature Request <https://sourceforge.net/tracker/index.php?func=detail&aid=3297935&group_id=369321&atid=1539126> FIX: handled exception when updating Birthday and anniversary with invalid date, see <https://sourceforge.net/tracker/?func=detail&aid=3397921&group_id=369321&atid=1539126> FIX: Handle Nullpointerexception when Release Marshall Objects at GetOutlookItems, maybe this helps to fix the Nullpointer Exceptions in LoadOutlookContacts

Contact Sync Mod, Version 3.5.3 Improvement: Upgraded to Google Data API 1.8 FIX: Handle Nullpointerexception when Release Marshall Objects

Contact Sync Mod, Version 3.5.1

FIX: Handle AccessViolation Exception when trying to get Email address from Exchange Email

Contact Sync Mod, Version 3.5

FIX: Moved NotificationReceived to constructor to not handle this event redundantly FIX: moved assert of TestSyncPhoto above the UpdateContact line FIX: Added log message when skipping a faulty Outlook Contact FIX: fixed number of current match (i not i+1) because of 1 based array Fix: set SyncDelete at every SyncStart to avoid "Skipped deletion" warnings, though Sync Deletion checkbox was checked Improvement: Support large Exchange contact lists, get SMTP email when Exchange returns X500 addresses, use running Outlook instance if present.

CHANGE 1: Support a large number of contacts on Exchange server without hitting the policy limitation of max number of contacts that can be processed simultaneously.

CHANGE 2: Enhancement request 3156687: Properly get the SMTP email address of Exchange contacts when Exchange returns X500 addresses.

CHANGE 3: Try to contact a running Outlook application before trying to launch a new one. Should make the program work in any situation, whether Outlook is running or not.

OTHER SMALL FIXES:

- Never re-throw an exception using "throw ex". Just use "throw". (preserves stack trace)
- Handle an invalid photo on a Google contact (skip the photo).

IMPROVEMENT: added EnableLaunchApplication to start GOContactSyncMod after installation as PostBuildEvent Improvement: added progress notifications (which contact is currently syncing or matching) Improvement: Sync also State and PostOfficeBox, see Tracker item <https://sourceforge.net/tracker/?func=detail&aid=3276467&group_id=369321&atid=1539126> Improvement: Avoid MatchContacts when just resetting matches (Performance improvement)

[3.9.1]: http://sourceforge.net/projects/googlesyncmod/files/Releases/3.9.1/SetupGCSM-3.9.1.msi/download
[3.9.2]: http://sourceforge.net/projects/googlesyncmod/files/Releases/3.9.2/SetupGCSM-3.9.2.msi/download
