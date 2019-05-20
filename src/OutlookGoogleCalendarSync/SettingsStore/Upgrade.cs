using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Xml.Linq;

namespace OutlookGoogleCalendarSync.SettingsStore {
    public static class Upgrade {
        private static readonly ILog log = LogManager.GetLogger(typeof(Upgrade));

        //OGCS releases that require the settings XML to be upgraded
        private const Int32 multipleCalendars = 2070901; //v2.7.9.1;

        private static String currentVersion;


        public static void Check() {
            currentVersion = XMLManager.ImportElement("Version", Settings.ConfigFile);

            while (upgradePerformed()) {
            }
        }

        private static Boolean upgradePerformed() {
            Int32 settingsNum = Program.VersionToInt(currentVersion);

            try {
                if (settingsNum < multipleCalendars) {
                    upgradeToMultiCalendar();
                    settingsNum = multipleCalendars;
                    return true;
                } else
                    return false;
            } catch {
                log.Warn("Upgrade(s) didn't complete successfully. The user will likely need to reset their settings.");
                return false;
            }
        }

        private static void backupSettingsFile() {
            try {
                log.Info("Backing up '" + Settings.ConfigFile + "' for v" + currentVersion);
                String backupFile = System.Text.RegularExpressions.Regex.Replace(Settings.ConfigFile, @"(\.\w+)$", "-v" + currentVersion + "$1");
                File.Copy(Settings.ConfigFile, backupFile);
                log.Info(backupFile + " created.");
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to create backup settings file", ex);
            }
        }

        private static void upgradeToMultiCalendar() {
            backupSettingsFile();

            XDocument xml = null;
            try {
                xml = XDocument.Load(Settings.ConfigFile);
                XElement settingsElement = XMLManager.GetElement("Settings", xml);
                XElement calendarsElement = XMLManager.AddElement("Calendars", settingsElement);
                XElement calendarElement = XMLManager.AddElement("Calendar", calendarsElement);

                XMLManager.MoveElement("DaysInTheFuture", settingsElement, calendarElement);

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Problem encountered whilst upgrading " + Settings.ConfigFilename, ex);
                throw ex;
            } finally {
                if (xml != null) {
                    xml.Root.Sort();
                    try {
                        xml.Save(Settings.ConfigFile);
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse("Could not save upgraded settings file " + Settings.ConfigFile, ex);
                        throw ex;
                    }
                }
            }
        }
    }
}
