using Google.Apis.Calendar.v3.Data;
using log4net;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;


namespace OutlookGoogleCalendarSync.GoogleOgcs {
    public class EventColour {
        public class Palette {
            public String Id { get; }
            public String HexValue { get; }
            public Color RgbValue { get; }
            public String Name { get {
                    String name = "";
                    try {
                        name = names[Id];
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse(ex);
                        name = HexValue;
                    }
                    return name;
                }
            }

            public static Palette NullPalette = new Palette(null, null, System.Drawing.Color.Transparent);

            public Palette(String id, String hexValue, Color rgbValue) {
                this.Id = id;
                this.HexValue = hexValue;
                this.RgbValue = rgbValue;
            }

            public override String ToString() {
                return "ID: " + Id + "; HexValue: " + HexValue + "; RgbValue: " + RgbValue +"; Name: "+ Name;
            }

            private static Dictionary<String,String> names = new Dictionary<String, String> {
                {"1", "Tomato"},
                { "2", "Flamingo" },
                { "3", "Tangerine" },
                { "4", "Banana" },
                { "5", "Sage" },
                { "6", "Peacock" },
                { "7", "Blueberry" },
                { "8", "Lavendar" },
                { "9", "Grape" },
                { "10", "Graphite" },
                { "11", "Basil" },
                { "Custom", "Calendar Default" }
            };

            public static String GetColourId(String name) {
                String id = null;
                try {
                    id = names.First(n => n.Value == name).Key;
                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Could not find colour ID for '" + name + "'.", ex);
                }
                return id;
            }

            public static String GetColourName(String id) {
                String name = null;
                try {
                    name = names[id];
                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Could not find colour name for '" + id + "'.", ex);
                }
                return name;
            }
        }

        private static readonly ILog log = LogManager.GetLogger(typeof(EventColour));
        private List<Palette> calendarPalette;
        private List<Palette> eventPalette;
        /// <summary>
        /// All event colours, including currently used calendar "custom" colour
        /// </summary>
        public List<Palette> ActivePalette {
            get {
                List<Palette> activePalette = new List<Palette>();

                //Palette currentCal = calendarPalette.Find(p => p.Id == Settings.Instance.UseGoogleCalendar.ColourId);
                Palette currentCal = null;
                foreach (Palette cal in calendarPalette) {
                    if (cal.Id == Settings.Instance.UseGoogleCalendar.ColourId) {
                        currentCal = cal;
                        break;
                    }
                }
                activePalette.Add(new Palette("Custom", currentCal.HexValue, currentCal.RgbValue));

                activePalette.AddRange(eventPalette);
                return activePalette;
            }
        }

        public EventColour() { }

        /// <summary>
        /// Retrieve calendar's Event colours from Google
        /// </summary>
        public void Get() {
            log.Debug("Retrieving calendar Event colours.");
            Colors colours = null;
            calendarPalette = new List<Palette>();
            eventPalette = new List<Palette>();
            try {
                colours = GoogleOgcs.Calendar.Instance.Service.Colors.Get().Execute();
            } catch (System.Exception ex) {
                log.Error("Failed retrieving calendar Event colours.");
                OGCSexception.Analyse(ex);
                return;
            }

            if (colours == null) log.Warn("No colours found!");
            else log.Debug(colours.Event__.Count() + " event colours and "+ colours.Calendar.Count() +" calendars (with a colour) found.");
            
            foreach (KeyValuePair<String, ColorDefinition> colour in colours.Event__) {
                eventPalette.Add(new Palette(colour.Key, colour.Value.Background, OutlookOgcs.Categories.Map.RgbColour(colour.Value.Background)));
            }
            foreach (KeyValuePair<String, ColorDefinition> colour in colours.Calendar) {
                calendarPalette.Add(new Palette(colour.Key, colour.Value.Background, OutlookOgcs.Categories.Map.RgbColour(colour.Value.Background)));
            }
        }

        /// <summary>
        /// Get the Google Palette from its Google ID
        /// </summary>
        /// <param name="colourId">Google ID</param>
        public Palette GetColour(String colourId) {
            Palette gColour = this.ActivePalette.Where(x => x.Id == colourId).FirstOrDefault();
            if (gColour != null)
                return gColour;
            else
                return Palette.NullPalette;
        }

        /// <summary>
        /// Find the closest colour palette offered by Google.
        /// </summary>
        /// <param name="colour">The colour to search with.</param>
        public Palette GetClosestColour(Color baseColour) {
            try {
                var colourDistance = ActivePalette.Select(x => new { Value = x, Diff = GetDiff(x.RgbValue, baseColour) }).ToList();
                var minDistance = colourDistance.Min(x => x.Diff);
                return colourDistance.Find(x => x.Diff == minDistance).Value;
            } catch (System.Exception ex) {
                log.Warn("Failed to get closest Event colour for " + baseColour.Name);
                OGCSexception.Analyse(ex);
                return Palette.NullPalette;
            }
        }

        public static int GetDiff(Color colour, Color baseColour) {
            int a = colour.A - baseColour.A,
                r = colour.R - baseColour.R,
                g = colour.G - baseColour.G,
                b = colour.B - baseColour.B;
            return (a * a) + (r * r) + (g * g) + (b * b);
        }
    }
}
