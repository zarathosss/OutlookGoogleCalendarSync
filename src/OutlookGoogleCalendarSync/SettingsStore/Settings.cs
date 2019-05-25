using log4net;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// The main Settings class.
    /// </summary>

    [DataContract]
    public class Settings {
        private static readonly ILog log = LogManager.GetLogger(typeof(Settings));

        private static String configFilename = "settings.xml";
        public static String ConfigFilename {
            get { return configFilename; }
        }
        /// <summary>
        /// Absolute path to config file, eg C:\foo\bar\settings.xml
        /// </summary>
        public static String ConfigFile {
            get { return Path.Combine(Program.WorkingFilesDirectory, ConfigFilename); }
        }

        public static void InitialiseConfigFile(String filename, String directory = null) {
            if (!string.IsNullOrEmpty(filename)) configFilename = filename;
            Program.WorkingFilesDirectory = directory;

            if (string.IsNullOrEmpty(Program.WorkingFilesDirectory)) {
                if (Program.IsInstalled || File.Exists(Path.Combine(Program.RoamingProfileOGCS, ConfigFilename)))
                    Program.WorkingFilesDirectory = Program.RoamingProfileOGCS;
                else
                    Program.WorkingFilesDirectory = System.Windows.Forms.Application.StartupPath;
            }

            if (!File.Exists(ConfigFile)) {
                log.Info("No settings.xml file found in " + Program.WorkingFilesDirectory);
                Settings.Instance.Save(ConfigFile);
                log.Info("New blank template created.");
                if (!Program.IsInstalled)
                    XMLManager.ExportElement("Portable", true, ConfigFile);
            }

            log.Info("Running OGCS from " + System.Windows.Forms.Application.ExecutablePath);
        }

        private static Settings instance;
        //Settings saved immediately
        private Boolean apiLimit_inEffect;
        private DateTime apiLimit_lastHit;
        private DateTime lastSyncDate;
        private Int32 completedSyncs;
        private Boolean portable;
        private Boolean alphaReleases;
        private String version;
        private Boolean donor;
        private Boolean hideSplashScreen;
        private Boolean suppressSocialPopup;

        public Settings() {
            setDefaults();
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        private void setDefaults() {
            //Default values
            assignedClientIdentifier = "";
            assignedClientSecret = "";
            PersonalClientIdentifier = "";
            PersonalClientSecret = "";

            apiLimit_inEffect = false;
            apiLimit_lastHit = DateTime.Parse("01-Jan-2000");
            GaccountEmail = "";

            Calendar = new SettingsStore.Calendar();

            MuteClickSounds = false;
            ShowBubbleTooltipWhenSyncing = true;
            StartOnStartup = false;
            StartupDelay = 0;
            StartInTray = false;
            MinimiseToTray = false;
            MinimiseNotClose = false;
            ShowBubbleWhenMinimising = true;

            CreateCSVFiles = false;
            LoggingLevel = "DEBUG";
            portable = false;
            Proxy = new SettingsStore.Proxy();

            alphaReleases = !System.Windows.Forms.Application.ProductVersion.EndsWith("0.0");
            SkipVersion = null;
            Subscribed = DateTime.Parse("01-Jan-2000");
            donor = false;
            hideSplashScreen = false;
            suppressSocialPopup = false;
            
            lastSyncDate = new DateTime(0);
            completedSyncs = 0;
            VerboseOutput = true;
        }

        public static Boolean InstanceInitialiased() {
            return (instance != null);
        }

        public static Settings Instance {
            get {
                if (instance == null) instance = new Settings();
                return instance;
            }
            set {
                instance = value;
            }
        }

        #region Google
        private String assignedClientIdentifier;
        [DataMember] public String AssignedClientIdentifier {
            get { return assignedClientIdentifier; }
            set {
                assignedClientIdentifier = value.Trim();
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("AssignedClientIdentifier", value.Trim(), ConfigFile);
            }
        }
        private String assignedClientSecret;
        [DataMember] public String AssignedClientSecret {
            get { return assignedClientSecret; }
            set {
                assignedClientSecret = value.Trim();
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("AssignedClientSecret", value.Trim(), ConfigFile);
            }
        }
        private String personalClientIdentifier;
        private String personalClientSecret;
        [DataMember] public String PersonalClientIdentifier {
            get { return personalClientIdentifier; }
            set { personalClientIdentifier = value.Trim(); }
        }
        [DataMember] public String PersonalClientSecret {
            get { return personalClientSecret; }
            set { personalClientSecret = value.Trim(); }
        }
        public Boolean UsingPersonalAPIkeys() {
            return !string.IsNullOrEmpty(PersonalClientIdentifier) && !string.IsNullOrEmpty(PersonalClientSecret);
        }
        [DataMember] public Boolean APIlimit_inEffect {
            get { return apiLimit_inEffect; }
            set {
                apiLimit_inEffect = value;
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("APIlimit_inEffect", value, ConfigFile);
            }
        }
        [DataMember] public DateTime APIlimit_lastHit {
            get { return apiLimit_lastHit; }
            set {
                apiLimit_lastHit = value;
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("APIlimit_lastHit", value, ConfigFile);
            }
        }
        [DataMember] public String GaccountEmail { get; set; }
        public String GaccountEmail_masked() {
            if (string.IsNullOrWhiteSpace(GaccountEmail)) return "<null>";
            return EmailAddress.MaskAddress(GaccountEmail);
        }
        #endregion
        #region App behaviour
        [DataMember] public bool HideSplashScreen {
            get { return hideSplashScreen; }
            set {
                if (!Settings.Instance.Loading() && hideSplashScreen != value) {
                    XMLManager.ExportElement("HideSplashScreen", value, ConfigFile);
                    if (Forms.Main.Instance != null) Forms.Main.Instance.cbHideSplash.Checked = value;
                }
                hideSplashScreen = value;
            }
        }

        [DataMember] public bool SuppressSocialPopup {
            get { return suppressSocialPopup; }
            set {
                if (!Settings.Instance.Loading() && suppressSocialPopup != value) {
                    XMLManager.ExportElement("SuppressSocialPopup", value, ConfigFile);
                    if (Forms.Main.Instance != null) Forms.Main.Instance.cbSuppressSocialPopup.Checked = value;
                }
                suppressSocialPopup = value;
            }
        }
        [DataMember] public bool ShowBubbleTooltipWhenSyncing { get; set; }
        [DataMember] public bool StartOnStartup { get; set; }
        [DataMember] public Int32 StartupDelay { get; set; }
        [DataMember] public bool StartInTray { get; set; }
        [DataMember] public bool MinimiseToTray { get; set; }
        [DataMember] public bool MinimiseNotClose { get; set; }
        [DataMember] public bool ShowBubbleWhenMinimising { get; set; }
        [DataMember] public bool Portable {
            get { return portable; }
            set {
                portable = value;
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("Portable", value, ConfigFile);
            }
        }

        [DataMember] public bool CreateCSVFiles { get; set; }
        [DataMember] public String LoggingLevel { get; set; }
        private bool? cloudLogging;
        [DataMember] public bool? CloudLogging {
            get { return cloudLogging; }
            set {
                cloudLogging = value;
                GoogleOgcs.ErrorReporting.SetThreshold(value ?? false);
            }
        }
        //Proxy
        [DataMember] public SettingsStore.Proxy Proxy { get; set; }
        [DataMember] public SettingsStore.Calendar Calendar { get; set; }
        #endregion
        #region About
        [DataMember] public string Version {
            get { return version; }
            set {
                if (version != null && version != value) {
                    XMLManager.ExportElement("Version", value, ConfigFile);
                }
                version = value;
            }
        }
        [DataMember] public bool AlphaReleases {
            get { return alphaReleases; }
            set {
                alphaReleases = value;
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("AlphaReleases", value, ConfigFile);
            }
        }
        public Boolean UserIsBenefactor() {
            return Subscribed != DateTime.Parse("01-Jan-2000") || donor;
        }
        [DataMember] public DateTime Subscribed { get; set; }
        [DataMember] public Boolean Donor {
            get { return donor; }
            set {
                donor = value;
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("Donor", value, ConfigFile);
            }
        }
        #endregion

        [DataMember] public DateTime LastSyncDate {
            get { return lastSyncDate; }
            set {
                lastSyncDate = value;
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("LastSyncDate", value, ConfigFile);
            }
        }
        [DataMember] public Int32 CompletedSyncs {
            get { return completedSyncs; }
            set {
                completedSyncs = value;
                if (!Settings.Instance.Loading()) XMLManager.ExportElement("CompletedSyncs", value, ConfigFile);
            }
        }
        [DataMember] public bool VerboseOutput { get; set; }
        [DataMember] public bool MuteClickSounds { get; set; }
        [DataMember] public String SkipVersion { get; set; }

        private static Boolean isLoaded = false;
        public static Boolean IsLoaded {
            get { return isLoaded; }
        }

        public static void Load(String XMLfile = null) {
            try {
                Settings.Instance = XMLManager.Import<Settings>(XMLfile ?? ConfigFile);
                log.Fine("User settings loaded.");
                Settings.isLoaded = true;
            } catch (ApplicationException ex) {
                log.Error(ex.Message);
                ResetFile(XMLfile);
                try {
                    Settings.Instance = XMLManager.Import<Settings>(XMLfile ?? ConfigFile);
                    log.Debug("User settings loaded successfully this time.");
                } catch (System.Exception ex2) {
                    log.Error("Still failed to load settings!");
                    OGCSexception.Analyse(ex2);
                }
            }
        }

        public static void ResetFile(String XMLfile = null) {
            System.Windows.Forms.MessageBox.Show("Your OGCS settings appear to be corrupt and will have to be reset.",
                    "Corrupt OGCS Settings", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            log.Warn("Resetting settings.xml file to defaults.");
            System.IO.File.Delete(XMLfile ?? ConfigFile);
            Settings.Instance.Save(XMLfile ?? ConfigFile);
        }

        public void Save(String XMLfile = null) {
            log.Info("Saving settings.");
            XMLManager.Export(this, XMLfile ?? ConfigFile);
        }

        public Boolean Loading() {
            StackTrace stackTrace = new StackTrace();
            foreach (StackFrame frame in stackTrace.GetFrames().Reverse()) {
                if (new String[] {"Load","isNewVersion"}.Contains(frame.GetMethod().Name)) {
                    return true;
                }
            }
            return false;
        }

        public void LogSettings() {
            log.Info(ConfigFile);
            log.Info("OUTLOOK SETTINGS:-");
            log.Info("  Service: "+ Calendar.OutlookService.ToString());
            if (Calendar.OutlookService == OutlookOgcs.Calendar.Service.SharedCalendar) {
                log.Info("  Shared Calendar: " + Calendar.SharedCalendar);
            } else {
                log.Info("  Mailbox/FolderStore Name: " + Calendar.MailboxName);
            }
            log.Info("  Calendar: "+ (Calendar.UseOutlookCalendar.Name=="Calendar"?"Default ":"") + Calendar.UseOutlookCalendar.Name);
            log.Info("  Category Filter: " + Calendar.CategoriesRestrictBy.ToString());
            log.Info("  Categories: " + String.Join(",", Calendar.Categories.ToArray()));
            log.Info("  Only Responded Invites: " + Calendar.OnlyRespondedInvites);
            log.Info("  Filter String: " + Calendar.OutlookDateFormat);
            log.Info("  GAL Blocked: " + Calendar.OutlookGalBlocked);
            
            log.Info("GOOGLE SETTINGS:-");
            log.Info("  Calendar: " + Calendar.UseGoogleCalendar.Name);
            log.Info("  Personal API Keys: " + UsingPersonalAPIkeys());
            log.Info("    Client Identifier: " + PersonalClientIdentifier);
            log.Info("    Client Secret: " + (PersonalClientSecret.Length < 5
                ? "".PadLeft(PersonalClientSecret.Length, '*')
                : PersonalClientSecret.Substring(0, PersonalClientSecret.Length - 5).PadRight(5, '*')));
            log.Info("  API attendee limit in effect: " + APIlimit_inEffect);
            log.Info("  API attendee limit last reached: " + APIlimit_lastHit);
            log.Info("  Assigned API key: " + AssignedClientIdentifier);
            log.Info("  Cloak Email: " + Calendar.CloakEmail);
        
            log.Info("SYNC OPTIONS:-");
            log.Info(" How");
            log.Info("  SyncDirection: "+ Calendar.SyncDirection.Name);
            log.Info("  MergeItems: " + Calendar.MergeItems);
            log.Info("  DisableDelete: " + Calendar.DisableDelete);
            log.Info("  ConfirmOnDelete: " + Calendar.ConfirmOnDelete);
            log.Info("  SetEntriesPrivate: " + Calendar.SetEntriesPrivate);
            log.Info("  SetEntriesAvailable: " + Calendar.SetEntriesAvailable);
            log.Info("  SetEntriesColour: " + Calendar.SetEntriesColour + (Calendar.SetEntriesColour ? "; " + Calendar.SetEntriesColourValue + "; \"" + Calendar.SetEntriesColourName + "\"" : ""));
            if ((Calendar.SetEntriesPrivate || Calendar.SetEntriesAvailable || Calendar.SetEntriesColour) && Calendar.SyncDirection == Sync.Direction.Bidirectional) {
                log.Info("    TargetCalendar: " + Calendar.TargetCalendar.Name);
                log.Info("    CreatedItemsOnly: " + Calendar.CreatedItemsOnly);
            }
            log.Info("  Obfuscate Words: " + Calendar.Obfuscation.Enabled);
            if (Calendar.Obfuscation.Enabled) {
                if (Settings.Instance.Calendar.Obfuscation.FindReplace.Count == 0) log.Info("    No regex defined.");
                else {
                    foreach (FindReplace findReplace in Settings.Instance.Calendar.Obfuscation.FindReplace) {
                        log.Info("    '" + findReplace.find + "' -> '" + findReplace.replace + "'");
                    }
                }
            }
            log.Info(" When");
            log.Info("  DaysInThePast: "+ Calendar.DaysInThePast);
            log.Info("  DaysInTheFuture:" + Calendar.DaysInTheFuture);
            log.Info("  SyncInterval: " + Calendar.SyncInterval);
            log.Info("  SyncIntervalUnit: " + Calendar.SyncIntervalUnit);
            log.Info("  Push Changes: " + Calendar.OutlookPush);
            log.Info(" What");
            log.Info("  AddLocation: " + Calendar.AddLocation);
            log.Info("  AddDescription: " + Calendar.AddDescription + "; OnlyToGoogle: " + Calendar.AddDescription_OnlyToGoogle);
            log.Info("  AddAttendees: " + Calendar.AddAttendees);
            log.Info("  AddColours: " + Calendar.AddColours);
            log.Info("  AddReminders: " + Calendar.AddReminders);
            log.Info("    UseGoogleDefaultReminder: " + Calendar.UseGoogleDefaultReminder);
            log.Info("    UseOutlookDefaultReminder: " + Calendar.UseOutlookDefaultReminder);
            log.Info("    ReminderDND: " + Calendar.ReminderDND + " (" + Calendar.ReminderDNDstart.ToString("HH:mm") + "-" + Calendar.ReminderDNDend.ToString("HH:mm") + ")");
            
            log.Info("PROXY:-");
            log.Info("  Type: " + Proxy.Type);
            if (Proxy.BrowserUserAgent != Proxy.DefaultBrowserAgent)
                log.Info("  Browser Agent: " + Proxy.BrowserUserAgent);
            if (Proxy.Type == "Custom") {
                log.Info("  Server Name: " + Proxy.ServerName);
                log.Info("  Port: " + Proxy.Port.ToString());
                log.Info("  Authentication Required: " + Proxy.AuthenticationRequired);
                log.Info("  UserName: " + Proxy.UserName);
                log.Info("  Password: " + (string.IsNullOrEmpty(Proxy.Password) ? "" : "*********"));
            } 
        
            log.Info("APPLICATION BEHAVIOUR:-");
            log.Info("  ShowBubbleTooltipWhenSyncing: " + ShowBubbleTooltipWhenSyncing);
            log.Info("  StartOnStartup: " + StartOnStartup + "; DelayedStartup: "+ StartupDelay.ToString());
            log.Info("  HideSplashScreen: " + (UserIsBenefactor() ? HideSplashScreen.ToString() : "N/A"));
            log.Info("  SuppressSocialPopup: " + (UserIsBenefactor() ? SuppressSocialPopup.ToString() : "N/A"));
            log.Info("  StartInTray: " + StartInTray);
            log.Info("  MinimiseToTray: " + MinimiseToTray);
            log.Info("  MinimiseNotClose: " + MinimiseNotClose);
            log.Info("  ShowBubbleWhenMinimising: " + ShowBubbleWhenMinimising);
            log.Info("  Portable: " + Portable);
            log.Info("  CreateCSVFiles: " + CreateCSVFiles);

            log.Info("  VerboseOutput: " + VerboseOutput);
            log.Info("  MuteClickSounds: " + MuteClickSounds);
            //To pick up from settings.xml file:
            //((log4net.Repository.Hierarchy.Hierarchy)log.Logger.Repository).Root.Level.Name);
            log.Info("  Logging Level: "+ LoggingLevel);
            log.Info("  Error Reporting: " + CloudLogging ?? "Undefined");

            log.Info("ABOUT:-");
            log.Info("  Alpha Releases: " + alphaReleases);
            log.Info("  Skip Version: " + SkipVersion);
            log.Info("  Subscribed: " + Subscribed.ToString("dd-MMM-yyyy"));
            log.Info("  Timezone Database: " + TimezoneDB.Instance.Version);
            
            log.Info("ENVIRONMENT:-");
            log.Info("  Current Locale: " + System.Globalization.CultureInfo.CurrentCulture.Name);
            log.Info("  Short Date Format: "+ System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
            log.Info("  Short Time Format: "+ System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern);
            log.Info("  Completed Syncs: "+ CompletedSyncs);
        }

        public static void configureLoggingLevel(string logLevel) {
            log.Info("Logging level configured to '" + logLevel + "'");
            ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).Root.Level = log.Logger.Repository.LevelMap[logLevel];
            ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).RaiseConfigurationChanged(EventArgs.Empty);
        }
    }
}
