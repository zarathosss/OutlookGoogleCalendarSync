using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class ColourMap : Form {

        private static readonly ILog log = LogManager.GetLogger(typeof(ColourMap));
        public static Extensions.OutlookColourPicker OutlookComboBox = new Extensions.OutlookColourPicker();
        public static Extensions.GoogleColourPicker GoogleComboBox = new Extensions.GoogleColourPicker();
        
        public ColourMap() {
            OutlookComboBox = new Extensions.OutlookColourPicker();
            OutlookComboBox.AddCategoryColours();
            GoogleComboBox = new Extensions.GoogleColourPicker();
            GoogleComboBox.AddPaletteColours();

            InitializeComponent();
            initialiseDataGridView();
            colourGridView.AllowUserToAddRows = true;
        }

        private void initialiseDataGridView() {
            try {
                log.Info("Opening colour mapping window.");
                
                //loadConfig();

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }
        /*
        private void loadConfig() {
            try {
                colourGridView.AllowUserToAddRows = true;
                if (Settings.Instance.TimezoneMaps.Count > 0) colourGridView.Rows.Clear();
                foreach (KeyValuePair<String, String> tzMap in Settings.Instance.TimezoneMaps) {
                    addRow(tzMap.Key, tzMap.Value);
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Populating gridview cells from Settings.", ex);
            }
        }

        private void addRow(String organiserTz, String systemTz) {
            int lastRow = 0;
            try {
                lastRow = colourGridView.Rows.GetLastRow(DataGridViewElementStates.None);
                Object currentValue = colourGridView.Rows[lastRow].Cells["OrganiserTz"].Value;
                if (currentValue != null && currentValue.ToString() != "") {
                    lastRow++;
                    colourGridView.Rows.Insert(lastRow);
                }
                colourGridView.Rows[lastRow].Cells["OrganiserTz"].Value = organiserTz;
                colourGridView.Rows[lastRow].Cells["SystemTz"].Value = systemTz;

                colourGridView.CurrentCell = colourGridView.Rows[lastRow].Cells[1];
                colourGridView.NotifyCurrentCellDirty(true);
                colourGridView.NotifyCurrentCellDirty(false);

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Adding timezone map row #" + lastRow, ex);
            }
        }
        */
        public static TimeZoneInfo GetSystemTimezone(String organiserTz, System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> sysTZ) {
            TimeZoneInfo tzi = null;
            /*if (Settings.Instance.TimezoneMaps.ContainsKey(organiserTz)) {
                tzi = sysTZ.FirstOrDefault(t => t.Id == Settings.Instance.TimezoneMaps[organiserTz]);
                if (tzi != null) {
                    log.Debug("Using custom timezone mapping ID '" + tzi.Id + "' for '" + organiserTz + "'");
                    return tzi;
                } else log.Warn("Failed to convert custom timezone mapping to any available system timezone.");
            }*/
            return tzi;            
        }

        #region EVENTS
        private void btSave_Click(object sender, EventArgs e) {
            /*try {
                Settings.Instance.TimezoneMaps.Clear();
                foreach (DataGridViewRow row in colourGridView.Rows) {
                    if (row.Cells[0].Value == null || row.Cells[0].Value.ToString().Trim() == "") continue;
                    try {
                        Settings.Instance.TimezoneMaps.Add(row.Cells[0].Value.ToString(), row.Cells[1].Value.ToString());
                    } catch (System.ArgumentException ex) {
                        if (OGCSexception.GetErrorCode(ex) == "0x80070057") {
                            //An item with the same key has already been added
                        } else throw;
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not save timezone mappings to Settings.", ex);
            } finally {
                this.Close();
            }*/
        }

        private void colourGridView_DataError(object sender, DataGridViewDataErrorEventArgs e) {
            /*log.Error(e.Context.ToString());
            if (e.Exception.HResult == -2147024809) { //DataGridViewComboBoxCell value is not valid.
                DataGridViewCell cell = colourGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                log.Warn("Cell[" + cell.RowIndex + "][" + cell.ColumnIndex + "] has invalid value of '" + cell.Value + "'. Removing.");
                cell.OwningRow.Cells[0].Value = null;
                cell.OwningRow.Cells[1].Value = null;
            } else {
                try {
                    DataGridViewCell cell = colourGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    log.Debug("Cell[" + cell.RowIndex + "][" + cell.ColumnIndex + "] caused error.");
                } catch {
                } finally {
                    OGCSexception.Analyse("Bad cell value in timezone data grid.", e.Exception);
                }
            }*/
        }

        private void colourGridView_CellClick(object sender, DataGridViewCellEventArgs e) {
            if (!this.Visible) return;

            Boolean validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            //Check to make sure the cell clicked is the cell containing the combobox 
            if (validClick && colourGridView.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn) {
                colourGridView.BeginEdit(true);
                ((ComboBox)colourGridView.EditingControl).DroppedDown = true;
            }
        }
        
        private void colourGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e) {
            try {
                if (e.Control is ComboBox) {
                    ComboBox cb = e.Control as ComboBox;
                    cb.DrawMode = DrawMode.OwnerDrawFixed;
                    cb.SelectedIndexChanged -= colourGridView_SelectedIndexChanged;
                    cb.SelectedIndexChanged += colourGridView_SelectedIndexChanged;
                    if (cb is Extensions.OutlookColourCombobox) {
                        cb.DrawItem -= OutlookComboBox.ColourPicker_DrawItem;
                        cb.DrawItem += OutlookComboBox.ColourPicker_DrawItem;
                        OutlookComboBox.ColourPicker_DrawItem(sender, null);
                    } else if (cb is Extensions.GoogleColourCombobox) {
                        cb.DrawItem -= GoogleComboBox.ColourPicker_DrawItem;
                        cb.DrawItem += GoogleComboBox.ColourPicker_DrawItem;
                        GoogleComboBox.ColourPicker_DrawItem(sender, null);
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }
        #endregion


        private void colourGridView_SelectedIndexChanged(object sender, EventArgs e) {
            //((ComboBox)sender).BackColor = System.Drawing.Color.Red; // (System.Drawing.Color)((ComboBox)sender).SelectedItem;
            //colourGridView.CurrentCell.ba
        }

        private void colourGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e) {
            log.Debug("colourGridView_CurrentCellDirtyStateChanged");
            //colourGridView.CurrentCell.Style.BackColor = System.Drawing.Color.Blue;
            DataGridViewColumn col = colourGridView.Columns[colourGridView.CurrentCell.ColumnIndex];
            //if (col is DataGridViewComboBoxColumn) {
            //    colourGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //    colourGridView.EndEdit();
            //}
            //colourGridView.celEditingControl
            
        }

        private void colourGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e) {
            log.Debug("CellEndEdit");
            //colourGridView.CurrentCell.Style.BackColor = System.Drawing.Color.Blue;
        }

        private void colourGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            log.Debug("CellFormatting");
            //colourGridView.CurrentCell
            
        }

        private void colourGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e) {
            log.Debug("colourGridView_CellValueChanged");
        }

        private void colourGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e) {
            log.Debug("colourGridView_CellPainting "+ e.RowIndex +":"+ e.ColumnIndex);
            //e.PaintBackground(e.ClipBounds, true);
            //e.PaintContent(e.ClipBounds);
            //e.CellStyle.BackColor = System.Drawing.Color.Red;
            //e.Handled = true;
        }
    }
}
